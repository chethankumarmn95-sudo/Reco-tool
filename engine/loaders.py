"""
loaders.py
----------
Turns an uploaded file (xlsx or csv) into a clean pandas DataFrame.

Why this file exists on its own:
Every source system (Shopify, Delhivery, Gokwik, etc.) exports data slightly
differently - sometimes as .xlsx, sometimes .csv, sometimes with a blank first
row. Keeping "how do I read a file" in one place means the rest of the engine
never has to worry about file formats again.
"""

import pandas as pd


def load_table(file, sheet_name=None):
    """
    Load a single table from an uploaded file.

    file: a file path (str) or an uploaded-file object (e.g. from Streamlit)
    sheet_name: which sheet to read if it's an Excel workbook with multiple
                sheets. If None, reads the first/only sheet.

    Returns a pandas DataFrame with column names stripped of extra spaces.
    """
    name = getattr(file, "name", str(file))

    if name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    elif name.lower().endswith(".xls"):
        # 2026-08-31 fix - see load_table_smart() below / views/
        # page_upload.py's _read_all_sheets() for the full story: pandas'
        # default .xls reader (xlrd) failed to read a real, valid
        # Shiprocket export; python-calamine reads it correctly.
        try:
            df = pd.read_excel(file, sheet_name=sheet_name if sheet_name else 0, engine="calamine")
        except Exception:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name if sheet_name else 0, engine="xlrd")
    else:
        # sheet_name=0 means "first sheet" when the caller doesn't specify one
        df = pd.read_excel(file, sheet_name=sheet_name if sheet_name else 0)

    # Clean up column names - trailing/leading spaces cause silent bugs
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_table_smart(file, expected_signature=None, max_header_rows_to_try=40):
    """
    Like load_table(), but if the real header row isn't row 1 - e.g. someone
    inserted a title row or a blank row above the actual columns when
    exporting - this tries reading with the header a few rows further down
    until the expected columns actually show up.

    expected_signature: a list of column specs (each a str or list of
    aliases) that should be present once the header is found correctly.
    If None, just returns the plain row-1 read (no detection needed).

    max_header_rows_to_try defaults to 40, not a handful - real Indian bank
    statement exports (the main source of this problem in practice) commonly
    carry a 10-20+ row letterhead block (account holder address, IFSC/MICR,
    nomination status, etc.) before the actual transaction header row, and
    that block's length varies bank to bank. Trying a few dozen extra
    skiprows is cheap (each attempt just re-reads the file), so it's better
    to default generously here than to have every new bank format need a
    one-off code change to raise this number.

    Returns (df, header_row_used). header_row_used is 0 if no shifting was
    needed, or None if no row within max_header_rows_to_try worked - in
    that case the plain row-1 read is returned so the caller can still show
    a clear error naming what's missing.
    """
    name = getattr(file, "name", str(file))
    is_csv = name.lower().endswith(".csv")
    is_legacy_xls = name.lower().endswith(".xls")

    def _read(skip):
        if is_csv:
            file.seek(0)
            return pd.read_csv(file, skiprows=skip)
        file.seek(0)
        if is_legacy_xls:
            # 2026-08-31 fix (client-reported - see views/page_upload.py's
            # _read_all_sheets docstring for the full story): pandas'
            # default reader for legacy .xls is xlrd, which a real courier
            # export (Shiprocket's own COD Remittance file) failed to read
            # at all with a low-level "Workbook corruption" error, despite
            # the file opening fine in both Excel and LibreOffice.
            # python-calamine reads the same file correctly - tried first
            # here, with xlrd as a fallback rather than a hard requirement.
            try:
                return pd.read_excel(file, sheet_name=0, skiprows=skip, engine="calamine")
            except Exception:
                file.seek(0)
                return pd.read_excel(file, sheet_name=0, skiprows=skip, engine="xlrd")
        return pd.read_excel(file, sheet_name=0, skiprows=skip)

    df0 = _read(0)
    df0.columns = [str(c).strip() for c in df0.columns]

    if not expected_signature:
        return df0, 0

    def _matches(df):
        for spec in expected_signature:
            if resolve_col(df, spec) is None:
                return False
        return True

    if _matches(df0):
        return df0, 0

    for skip in range(1, max_header_rows_to_try + 1):
        try:
            df = _read(skip)
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        if _matches(df):
            return df, skip

    # Nothing worked - return the original so the caller's normal validation
    # error still fires with a sensible message.
    return df0, None


def normalize_order_id(series):
    """
    Order IDs show up inconsistently across systems: '#12345', '12345',
    12345 (as a number), ' #12345 '. This makes them all comparable by
    stripping to digits-only text. Used as the join key everywhere.
    """
    return (
        series.astype(str)
        .str.strip()
        .str.replace("#", "", regex=False)
        .str.replace(r"\.0$", "", regex=True)  # in case Excel read it as float
    )


def _normalize_col_name(name):
    return str(name).strip().lower().replace(" ", "").replace("_", "").replace(".", "")


def resolve_col(df, col_spec):
    """
    Finds the actual column in df matching col_spec, which can be a single
    column name (str) or a list of acceptable aliases - e.g. the Order ID
    column might be called "Name", "Order ID", "Shopify Order ID", "Order
    Number", or "Shopify Order Number" depending on the export. Matching is
    case/whitespace/punctuation-insensitive so small naming differences
    between exports don't break the pipeline.

    Returns the actual column name from df, or None if nothing matched.
    """
    candidates = col_spec if isinstance(col_spec, list) else [col_spec]
    lookup = {_normalize_col_name(c): c for c in df.columns}
    for candidate in candidates:
        key = _normalize_col_name(candidate)
        if key in lookup:
            return lookup[key]
    return None


def resolve_col_or_raise(df, col_spec, label=""):
    resolved = resolve_col(df, col_spec)
    if resolved is None:
        candidates = col_spec if isinstance(col_spec, list) else [col_spec]
        raise KeyError(
            f"Could not find any of the expected column(s) {candidates} "
            f"in the {label} file. Actual columns: {list(df.columns)}"
        )
    return resolved


def signature_match_count(df, required_specs):
    """How many of required_specs (each a column name or list of aliases,
    per resolve_col) actually resolve against df's columns. The building
    block behind every "does this sheet/file look like report X" check in
    this app (see find_sheet_by_columns below and engine.report_detection)
    - deliberately just counts header-name matches, never looks at a file
    name, sheet name, or sheet position."""
    if df is None or not required_specs:
        return 0
    return sum(1 for spec in required_specs if resolve_col(df, spec) is not None)


def find_sheet_by_columns(sheets, required_specs, min_required=None):
    """
    Finds which sheet in `sheets` (a {sheet_name: DataFrame} dict, e.g.
    from pd.read_excel(file, sheet_name=None)) actually HAS the data a
    report needs - identified purely by its column headers, never by the
    sheet's name or its position in the workbook. This is what makes a
    multi-sheet source (e.g. Shiprocket's COD export) robust to the
    platform renaming a sheet, reordering sheets, or changing
    capitalisation/spacing in a sheet name - none of that is looked at
    here at all.

    required_specs: the column specs (each a str or list of aliases) a
    sheet needs to plausibly BE the report being looked for.

    min_required: how many of required_specs must resolve for a sheet to
    count as a match at all. Defaults to requiring ALL of them (the
    strict case: "this must be exactly this report's sheet"). Pass a
    smaller number for a looser "close enough" match.

    When more than one sheet matches, the one matching the MOST columns
    wins; a tie is broken by whichever has more rows (a real data sheet
    is far more likely to be larger than an incidental summary/pivot
    sheet that happens to share a couple of column names).

    Returns (sheet_name, DataFrame) for the best match, or (None, None)
    if no sheet in the workbook matches well enough.
    """
    if min_required is None:
        min_required = len(required_specs)
    best = None
    for name, df in sheets.items():
        matched = signature_match_count(df, required_specs)
        if matched >= min_required:
            rank = (matched, len(df))
            if best is None or rank > best[0]:
                best = (rank, name, df)
    if best is None:
        return None, None
    return best[1], best[2]
