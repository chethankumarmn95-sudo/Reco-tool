"""
formatting.py
-------------
Three jobs:
1. indian_number() - formats big rupee values the way Indian accounting
   actually reads them (93,45,550 not 9,345,550) so the summary cards don't
   need to truncate.
2. style_workbook() - takes a plain pandas-written Excel file and makes it
   look like a workbook a controller actually produced: bold header row,
   frozen header, auto-sized columns, currency formatting, and the Query
   column colour-flagged in red where it isn't "Okk".
3. add_dashboard_sheet() - builds the front-page 'Dashboard' sheet: a
   management-facing MIS summary (KPI cards, reconciled-vs-unreconciled
   charts, a key metrics table, a trend chart) rather than a plain data
   sheet - see the "Dashboard build" section below.
"""

import copy
import datetime as dt
import io
import pandas as pd
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .settlement_pending import (
    NOT_REFLECTING_LABEL, COD_REPORT_ABSENT_QUERY_FRAGMENT, BROAD_NOT_REFLECTING_FRAGMENT,
)


def indian_number(value, decimals=0):
    """
    93,45,550 style grouping (lakh/crore) instead of 9,345,550 (western).
    Falls back gracefully for small numbers and negatives.
    """
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)

    sign = "-" if value < 0 else ""
    value = abs(value)
    whole = int(value)
    frac = value - whole

    s = str(whole)
    if len(s) <= 3:
        grouped = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        grouped = ",".join(parts) + "," + last3

    if decimals:
        grouped += f".{round(frac, decimals):.{decimals}f}".split(".")[1]

    return sign + grouped


def ensure_valid_sheet_names(sheet_data):
    """
    Excel hard-caps worksheet names at 31 characters and disallows the
    characters : \\ / ? * [ ] - openpyxl only warns (or in some cases says
    nothing at all) when a name breaks either rule rather than raising, so
    a workbook can write and save "successfully" here and still show
    Excel's own "we found a problem with some content, do you want us to
    recover as much as we can?" dialog the moment the file is opened for
    real - a silent corruption from this codebase's point of view, since
    nothing here ever sees an exception.

    Renames anything over the limit / with invalid characters right
    before writing (dropping invalid characters, then truncating), and
    keeps names unique if two long names happen to truncate to the same
    thing. Call this on the sheet_data dict right before the df.to_excel()
    loop, and pass the SAME (possibly renamed) dict on to style_workbook()
    afterwards - it looks sheets up by these exact names.
    """
    invalid_chars = set(':\\/?*[]')
    seen = set()
    fixed = {}
    for name, df in sheet_data.items():
        safe = "".join(c for c in str(name) if c not in invalid_chars).strip() or "Sheet"
        safe = safe[:31]
        base, n = safe, 2
        while safe in seen:
            suffix = f" ({n})"
            safe = base[: 31 - len(suffix)] + suffix
            n += 1
        seen.add(safe)
        fixed[safe] = df
    return fixed


# ---------------------------------------------------------------------------
# Large-sheet write performance (the "Amazon Settlement file has lakhs of
# rows, downloading the result takes too long" fix)
# ---------------------------------------------------------------------------
# Two separate thresholds, because they fix two separate cost centres:
#
# FAST_WRITE_ROW_THRESHOLD - above this many rows, skip pandas' own
# to_excel() writer (which builds one ExcelCell object per cell via its
# generic cross-engine formatter before handing it to openpyxl) in favour
# of writing straight into the openpyxl worksheet via ws.append() per row
# (see write_sheet_fast() below) - openpyxl accepts numpy/pandas scalar
# types (numpy.float64, pandas.Timestamp, NaN, NaT, None) directly and
# coerces them exactly the same way pandas' own writer would (verified by
# round-tripping a mixed-dtype frame through both paths and comparing the
# cell values/types read back) - this is purely a faster code path to the
# same result, not a behaviour change. Matches MAX_STYLED_ROWS above, since
# a sheet already skips the expensive per-cell styling pass at that size -
# consistent "this sheet is in bulk-data territory" threshold used
# everywhere in this module.
#
# EXCEL_SHEET_ROW_LIMIT - above this many rows, a sheet is EXCLUDED from
# the .xlsx entirely and offered as a separate CSV download instead (see
# write_workbook_sheets() below). Every pure-Python Excel writer (openpyxl
# AND xlsxwriter both benchmarked - see the perf investigation behind this
# fix) costs roughly the same, unavoidable ~0.1-0.15 milliseconds PER CELL
# once you're writing real workbook XML, AND openpyxl's own wb.save() step
# (serializing every cell object it built, across every sheet) scales with
# TOTAL workbook cell count, not per-sheet - measured directly at ~0.13ms/
# row (append + save combined) for a realistic ~15-column sheet. There is
# no code-level trick that makes a multi-hundred-thousand-row sheet fast to
# embed in an .xlsx, because that cost lives in the file format/library,
# not in this engine's own logic. A CSV, by contrast, writes the exact same
# lakhs-of-rows data in low single-digit seconds (no per-cell XML/style
# object overhead at all) with zero data loss - every column, every row,
# still fully present, just delivered as .csv instead of as one more Excel
# sheet when embedding it as Excel would blow the "reconciliation result
# available for download within a minute" target on its own.
#
# 75,000 (not a rounder/larger number) is deliberately conservative: at the
# ~0.13ms/row measured cost, a single 75,000-row sheet costs roughly 10s to
# write, leaving headroom in the ~60s target for the reconciliation compute
# itself, the Dashboard/chart build, and more than one large-ish sheet in
# the same workbook (a client's file can easily produce two independently
# large sheets - e.g. a large order-wise detail AND a large settlement
# ledger - in the same run). This only ever pushes out genuinely bulk data
# - every compact, one-row-per-order/settlement summary sheet stays well
# under this limit even for a full financial year's data (tens of
# thousands of orders is normal; 75,000+ rows in ANY one sheet, order-wise
# or line-item-level, is the genuinely extreme case this guards against).
FAST_WRITE_ROW_THRESHOLD = 20_000
EXCEL_SHEET_ROW_LIMIT = 75_000


def write_sheet_fast(writer, df, sheet_name):
    """
    Writes df into a new sheet on writer's underlying openpyxl workbook via
    ws.append() per row instead of pandas' own (slower, at real scale)
    to_excel() cell-by-cell writer - see FAST_WRITE_ROW_THRESHOLD's
    docstring above for why/when this is used, and for the
    correctness note (openpyxl's own value coercion is used either way,
    this only changes how each cell VALUE gets to the worksheet, not what
    value it ends up with).

    Registers the new worksheet into writer.sheets too, so every existing
    caller that looks a sheet up by name afterwards (style_workbook(),
    link_waterfall_formulas(), etc.) keeps working exactly as if
    df.to_excel(writer, sheet_name=sheet_name, index=False) had been
    called instead.
    """
    ws = writer.book.create_sheet(sheet_name)
    writer.sheets[sheet_name] = ws
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(row)
    return ws


def estimate_seconds_for_row_counts(row_counts):
    """
    Upfront time estimate for write_workbook_sheets() below, shown to the
    user BEFORE the write starts (the "This may take approximately X
    seconds/minutes" message the client asked for) - so the wait is never
    a blank spinner with no sense of how long it'll be.

    Takes a plain list of row counts rather than the sheet_data dict
    itself, so the Reports page can estimate BEFORE it has actually
    assembled every output sheet (sheet_data is only built partway
    through _build_workbook()/_build_amazon_workbook()) - the caller just
    passes the row counts of the few known-large input frames (e.g.
    reco_df, the Amazon expense ledger) that dominate the real write time.

    Uses the same per-row throughput figures this write path was
    benchmarked against (see write_sheet_fast/EXCEL_SHEET_ROW_LIMIT's
    docstrings): ~0.13ms/row for the normal/fast Excel-write paths
    (measured combined append+save cost), ~0.0054ms/row for a CSV
    overflow write (~5.4s per million rows) - then scaled by a single
    CALIBRATION_MULTIPLIER (4.0x) fitted against real, measured
    _build_workbook()/_build_amazon_workbook() runs (not just the raw
    write_workbook_sheets() cost this formula is otherwise based on).

    That multiplier exists because the real end-to-end build cost isn't
    just "write these N rows" - it also includes deriving the extra
    sheets (Amazon's order-wise/settlement-wise Expense Ledger pivots,
    the Exceptions sheet, etc.), running style_workbook()'s formatting
    over every sheet, and building the Dashboard sheet's charts - none of
    which this row-count-only estimate can see in advance. Measured
    directly (3 calibration runs each, DTC and Amazon, at 20k/80k/150k
    orders): the actual build time was consistently ~3.2-3.4x this
    formula's raw per-row figure whenever nothing overflowed to CSV yet
    (the worst case, since every sheet then pays the slower Excel-write
    path) - so build time is NOT monotonic in row count: a 20,000-order
    file that stays entirely below EXCEL_SHEET_ROW_LIMIT can take longer
    to build than an 80,000-order file where the largest sheets cross the
    threshold and get diverted to the much-faster CSV path instead. 4.0x
    was picked to safely cover the worst (no-overflow) case with a little
    headroom, which means it will OVER-estimate once overflow kicks in -
    intentional per this function's own "over-estimate, not under"
    philosophy below.

    Deliberately a slight over-estimate rather than an exact prediction -
    promising "3 seconds" and taking 4 reads as broken; promising "12
    seconds" and taking 3 reads as a pleasant surprise.
    """
    EXCEL_MS_PER_ROW = 0.13
    CSV_MS_PER_ROW = 0.0054
    PER_SHEET_OVERHEAD_S = 0.15
    CALIBRATION_MULTIPLIER = 4.0
    total_seconds = 0.0
    for n_rows in row_counts:
        total_seconds += PER_SHEET_OVERHEAD_S
        if n_rows > EXCEL_SHEET_ROW_LIMIT:
            total_seconds += (n_rows * CSV_MS_PER_ROW) / 1000.0
        else:
            total_seconds += (n_rows * EXCEL_MS_PER_ROW) / 1000.0
    return total_seconds * CALIBRATION_MULTIPLIER


def estimate_workbook_write_seconds(sheet_data):
    """Convenience wrapper around estimate_seconds_for_row_counts() for
    when the full sheet_data dict is already assembled."""
    return estimate_seconds_for_row_counts(len(df) for df in sheet_data.values())


def write_workbook_sheets(writer, sheet_data, oversized_note_col_width=100, progress_callback=None):
    """
    Shared "write every sheet in sheet_data into writer" step for both
    _build_workbook() (DTC) and _build_amazon_workbook() (Amazon) in
    views/page_reports.py - replaces the old `for name, df in
    sheet_data.items(): df.to_excel(...)` loop with one that stays fast
    and stays within a practical file size regardless of how large a
    client's raw export is:

      - A normal-sized sheet (<= FAST_WRITE_ROW_THRESHOLD rows) is written
        exactly as before, via df.to_excel().
      - A large sheet (> FAST_WRITE_ROW_THRESHOLD, <= EXCEL_SHEET_ROW_LIMIT
        rows) is written via write_sheet_fast() above - same content,
        faster path.
      - A VERY large sheet (> EXCEL_SHEET_ROW_LIMIT rows - in practice,
        only a full financial year's line-item-level export, e.g. the
        Amazon Expense Ledger's "Full Detail" sheet) is NOT embedded in
        the workbook at all. A small placeholder sheet explaining why
        (and how many rows/columns are in the full data) is written in
        its place, and the sheet's full DataFrame is returned as a CSV
        byte-string in the `overflow_csvs` dict so the caller (Reports
        page) can offer it as a separate download button. See
        EXCEL_SHEET_ROW_LIMIT's docstring above for why this is a size
        trade-off, not a data-loss one - nothing in the ledger is
        dropped, only ITS DELIVERY FORMAT changes when embedding it as
        an Excel sheet would make the whole report miss the "ready
        within a minute" target on its own.

    progress_callback(sheet_index, total_sheets, sheet_name, n_rows), if
    given, is called once BEFORE each sheet is written - lets the caller
    (Reports page) drive a real, honest progress bar/status line off
    actual work completed, rather than a spinner that gives no sense of
    whether the app is still working or has hung. Optional and purely
    cosmetic - never affects what gets written.

    Returns overflow_csvs: {sheet_name: csv_bytes} for every sheet that
    was too large to embed - empty dict when every sheet fit.
    """
    overflow_csvs = {}
    total_sheets = len(sheet_data)
    for i, (sheet_name, df) in enumerate(list(sheet_data.items())):
        n_rows = len(df)
        if progress_callback is not None:
            progress_callback(i, total_sheets, sheet_name, n_rows)
        if n_rows > EXCEL_SHEET_ROW_LIMIT:
            csv_buf = io.StringIO()
            df.to_csv(csv_buf, index=False)
            overflow_csvs[sheet_name] = csv_buf.getvalue().encode("utf-8")

            note_df = pd.DataFrame([{
                "Note": (
                    f"This sheet's full data ({n_rows:,} rows x {len(df.columns)} columns) was "
                    "too large to include directly in this Excel workbook without making the "
                    "whole report slow to generate and download. It has NOT been left out - "
                    "every row is available in the separate CSV file provided alongside this "
                    "download (same name as this sheet). Every other sheet in this workbook "
                    "(order-wise and settlement-wise summaries, the Waterfall, etc.) already "
                    "includes the full, correctly totalled figures - this only affects the "
                    "individual line-item-level trace/audit detail."
                ),
            }])
            ws = writer.book.create_sheet(sheet_name)
            writer.sheets[sheet_name] = ws
            ws.append(list(note_df.columns))
            for row in note_df.itertuples(index=False, name=None):
                ws.append(row)
            ws.column_dimensions["A"].width = oversized_note_col_width
            sheet_data[sheet_name] = note_df  # so style_workbook() styles the note sheet, not the huge original
        elif n_rows > FAST_WRITE_ROW_THRESHOLD:
            write_sheet_fast(writer, df, sheet_name)
        else:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return overflow_csvs


def strip_tz(df):
    """
    Excel (and openpyxl, the engine pandas uses to write it) has no concept
    of a timezone-aware datetime at all - not even as a fallback, it's a
    hard ValueError - so any column that ended up timezone-aware has to be
    made naive before a df.to_excel() call, or the whole report download
    crashes. This happens whenever a source file's own date/time column
    carries an explicit UTC/IST-style offset (e.g. Shopify's order export
    timestamps often look like "2026-03-05T10:00:00+05:30") - pandas then
    infers a timezone-aware dtype for that column right from the read, and
    it stays tz-aware through every merge/concat/groupby all the way to
    the final Reco working / Order Lookup / Bank Reco sheets.

    Drops the tz label rather than converting to UTC first - i.e. keeps
    the same wall-clock date/time the source file showed (10:00:00 stays
    10:00:00), it just stops being tagged with an offset, since that's
    what someone reading "order date" in the report actually wants to see.

    Call this on every DataFrame right before df.to_excel() - safe to call
    on frames with no tz-aware columns at all (returns them unchanged).
    """
    if df is None or getattr(df, "empty", True):
        return df
    df = df.copy()
    for col in df.columns:
        s = df[col]
        if isinstance(s.dtype, pd.DatetimeTZDtype):
            df[col] = s.dt.tz_localize(None)
        elif s.dtype == "object":
            # Rarer case: an object-dtype column holding individual
            # datetime/Timestamp values built up one row at a time (not
            # through one vectorized pd.to_datetime call), where only
            # some of them happen to carry tzinfo.
            has_tz_aware = s.map(
                lambda v: isinstance(v, dt.datetime) and v.tzinfo is not None
            ).any()
            if has_tz_aware:
                df[col] = s.map(
                    lambda v: v.replace(tzinfo=None) if isinstance(v, dt.datetime) and v.tzinfo is not None else v
                )
    return df


HEADER_FILL = PatternFill(start_color="1A1A2E", end_color="1A1A2E", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FLAG_FILL = PatternFill(start_color="FDE2E1", end_color="FDE2E1", fill_type="solid")
THIN_BORDER = Border(*(Side(style="thin", color="DDDDDD"),) * 4)

MONEY_COLUMNS = {
    "subtotal", "shipping", "taxes", "total", "receipt_amount", "total_deduction",
    "refund_amount", "diff", "settlement_amount", "order_value", "receipt",
    "deduction", "settlement", "exposure", "Order value", "Pending amount", "Receipt amount",
    # Exceptions sheet (engine.lookup.build_exception_detail) column names
    "Order Value", "Receipt Amount", "Pending Amount",
    # Gateway Settlement overall (engine.settlement.gateway_settlement_overall) -
    # client-reported 2026-08-30, items 5/6, replaces the old single
    # "Gateway Settlement" sheet's gateway_settlement_summary() output.
    "Settlement Done", "Deductions", "Pending for Settlement", "Total (Settled + Pending)",
    "Settlement Done - This Period (Same Month)", "Settlement Done - This Period (Next Month)",
    "Settlement Done - Other Period", "Order ID Not Found - Settled",
    "Total Settlement Done", "PG Deductions Reco period", "PG Deductions other period",
    # UTR-level bank reconciliation (engine.bank.bank_reconciliation_by_utr)
    "Amount (this period settled this period)", "Amount (this period transaction Settled Subsequent Period)",
    "Settled (subsequent period transaction subsequent period)", "Settled (previous period transaction settled this period)",
    "Order ID Not Found - Settled During Reco Period", "Order ID Not Found - Settled After Reco Period",
    "Bank Credit", "Difference",
    # Settlement Pending Report (engine.settlement_pending)
    "Order Amount", "Gateway Amount", "Settlement Amount", "Amount Pending", "Exception Amount",
    # Gateway Recon Recoperiod (engine.settlement.gateway_recon_by_period) -
    # client-reported 2026-08-30, items 5/6, replaces the old
    # "Gateway Reconciliation Health" sheet (engine.settlement_pending.
    # reconciliation_health_by_gateway).
    "Received in Bank", "Pending Settlement", "Pending Bank Matching", "Exception", "Total",
    "Gateway deduction", "Refund", "Part prepaid and part post paid",
    "Order Value partially not received", "Check", "Diff",
    # Amazon/marketplace channel (engine.amazon_reco / engine.amazon_consolidator /
    # engine.amazon_bank / engine.amazon_invoice_check) - Reports page workbook
    "Amount", "invoice_amount", "net_sales", "order_deductions", "order_credits", "net_payout",
    "taxable", "tax_amount", "principal_amount", "shipping_amount", "tcs_amount",
    "settlement_amount", "bank_amount", "batch_amount", "amount_as_parsed",
    "amount_as_per_amazon", "difference", "Settlement Total (Rs)", "Bank Receipts (Rs)",
    "Variance (Rs)",
    # "Reco working" sheet's own display header for settlement_amount
    # (client-reported 2026-08-30 rename - see views/page_reports.py's
    # _build_workbook()).
    "Bank credit",
    # "Open queries" sheet's own display headers (client-reported
    # 2026-08-30, item 4 - see engine/summary.py::open_queries() and
    # views/page_reports.py's _build_workbook() rename).
    "Recipt", "Bank receipt", "bank_receipt",
}

# Client-reported 2026-09-06 (round 18): "Bank Date" on the "Bank Reco
# (UTR-wise)" sheet (engine.bank.bank_reconciliation_by_utr()'s own
# output column) was showing as "YYYY-MM-DD-HH:mm:ss" - openpyxl's own
# default number format for a datetime cell (this column was never
# given an explicit one before) - rather than a clean date. Scoped to
# just this one column name (confirmed via grep to appear nowhere else
# in the workbook under this exact header - the Reco working/Order
# Lookup sheets carry the same value under the differently-named
# "Payment Date (Bank Date)" column instead, untouched here) rather than
# reformatting every date column in every sheet, which wasn't asked for
# and risks changing a format the client hasn't flagged as wrong.
DATE_COLUMNS = {
    "Bank Date": "DD-MM-YYYY",
    # 2026-09-06 (round 19, client-reported item 5): new column on the
    # "Settlement Pending Detail" sheet only (engine.settlement_pending.
    # build_settlement_pending_report()) - safe to add here unscoped since
    # the name is brand new and doesn't collide with any existing sheet.
    "Report Period End Date": "DD-MM-YYYY",
    # 2026-09-06 (round 20, client-reported item 1): "Payment Date (Bank
    # Date)" on the "Reco working" sheet (engine.bank.
    # build_order_level_utr_detail()'s own column - the same underlying
    # "Bank Date" value as above, just merged onto reco_df under this
    # different header) had no explicit number format at all, so it showed
    # openpyxl's default datetime rendering instead of a clean date. Client
    # asked specifically for "DD-MMM-YYYY" here (e.g. "06-Sep-2026") - a
    # different format string than "Bank Date"'s own "DD-MM-YYYY" above, so
    # kept as its own dict entry rather than reusing that one. Safe to add
    # unscoped: this exact column (same name, same value, same meaning) is
    # the only place it appears - also passed through onto the "Order
    # Lookup" sheet (engine.lookup.build_order_lookup()) where the same
    # clean format is equally correct, not an unrelated column that would
    # be wrongly reformatted.
    "Payment Date (Bank Date)": "DD-MMM-YYYY",
    # 2026-09-06 (round 20, client-reported items 2/3): "Refund Date" -
    # same "Reco working"/"Order Lookup" columns, same requested format.
    # The VALUE bug (wrong/blank dates) is fixed separately at the source
    # in engine/consolidator.py::normalize_gateway_df()'s _parse_date_col()
    # (dayfirst=True) - this entry only controls how the (now-correct)
    # value is displayed.
    "Refund Date": "DD-MMM-YYYY",
}

# 2026-09-06 (round 19, client-reported item 4): unlike "Report Period End
# Date" above, "Order Date" is NOT safe to add to the unscoped DATE_COLUMNS
# dict - engine.lookup.build_exception_detail() also has a column literally
# named "Order Date" on the "Exceptions" sheet, and the client only
# reported this problem on "Settlement Pending Detail" - reformatting the
# Exceptions sheet's own Order Date too would be an undisclosed, unasked-
# for change. Keyed by sheet TITLE (ws.title, already available to
# style_sheet() below) so only the one sheet that actually had its "Order
# Date" column converted to a real, tz-aware-then-stripped datetime (see
# that function's own docstring) gets the clean format; the Exceptions
# sheet's "Order Date" - still reco_df's raw, unconverted "created_at" -
# is untouched, exactly as before.
SHEET_SCOPED_DATE_COLUMNS = {
    "Settlement Pending Detail": {"Order Date": "DD-MM-YYYY"},
}

# 2026-09-06 (round 19, client-reported item 5): "Days Pending" on the
# Settlement Pending Detail sheet is now a live formula (Report Period End
# Date minus Order Date - see build_settlement_pending_report()) whose
# result is a plain day COUNT, not a currency figure or a date - without an
# explicit format Excel can sometimes auto-render a date-subtraction
# formula's result using a date format instead of a number. Name is unique
# to this one sheet (confirmed via grep), so left unscoped like Bank Date.
INTEGER_COLUMNS = {
    "Days Pending": "0",
}


# Sheets at or below this many rows get full per-cell styling (border +
# currency number format on every cell). Above it, that per-cell loop is
# skipped - header styling/freeze panes/column widths still apply, just not
# the cosmetic per-cell border+format pass. This isn't a cosmetic nice-to-
# have: openpyxl styling is a real Python-object-per-cell operation, and a
# report sheet with 100k+ rows (e.g. the Amazon Expense Ledger, kept at full
# transaction-line grain per the client's explicit "don't merge fee
# categories" requirement - see engine/amazon_consolidator.py) turns an
# O(rows x cols) loop into minutes of wall-clock time - confirmed by timing
# it directly: the ~159k-row Expense Ledger sheet alone took over 2 minutes
# and was still running, which is exactly the "downloadable report...
# nothing is coming" symptom reported against this page.
MAX_STYLED_ROWS = 20_000


def style_sheet(ws, df, query_col_name="query", header_color=None):
    """
    Apply header styling, column widths, currency format, banded rows, and
    query/status highlighting to one worksheet that pandas has already
    written. header_color (hex string, no '#') lets style_workbook() give
    each sheet its own accent colour - defaults to navy so any existing
    caller that doesn't pass one keeps the original look. Resolved lazily
    (not as a default parameter value) so this can reference the palette
    constants defined later in this module without a NameError at import
    time - see the "Dashboard build" section below for NAVY/FONT_NAME/etc.
    """
    header_color = header_color or NAVY
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")

    ws.sheet_view.showGridLines = False  # look like a finished report, not a raw worksheet

    # Header row
    ws.row_dimensions[1].height = 22
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"

    # Column widths always apply - cheap (only samples the first 200 rows).
    for col_idx, col_name in enumerate(df.columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).head(200)])
        ws.column_dimensions[letter].width = min(max(max_len + 3, 10), 45)

    # Per-cell border + standardised font + currency formatting + banded
    # rows (alternating tint for readability) - see MAX_STYLED_ROWS above
    # for why this whole pass is skipped on very large sheets.
    if len(df) <= MAX_STYLED_ROWS:
        band_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        body_font = Font(name=FONT_NAME, size=10)
        sheet_date_columns = SHEET_SCOPED_DATE_COLUMNS.get(ws.title, {})
        for row_idx in range(2, len(df) + 2):
            is_band = (row_idx % 2 == 0)
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = THIN_BORDER
                cell.font = body_font
                if is_band:
                    cell.fill = band_fill
                if col_name in MONEY_COLUMNS:
                    cell.number_format = "#,##0.00"
                elif col_name in DATE_COLUMNS:
                    cell.number_format = DATE_COLUMNS[col_name]
                elif col_name in sheet_date_columns:
                    cell.number_format = sheet_date_columns[col_name]
                elif col_name in INTEGER_COLUMNS:
                    cell.number_format = INTEGER_COLUMNS[col_name]

    # Highlight the query/status column wherever it flags an exception -
    # works for the "query" column (Reco working), "Reconciliation status"
    # (Order Lookup), "Reconciliation Status" (Exceptions sheet), "Remarks"
    # (UTR-level bank reconciliation - engine.bank bank_reconciliation_by_utr),
    # and "Reconciliation Category" (engine.bank classify_order_bank_status),
    # whichever is present.
    # Values are either a literal set (exact match) or a callable taking
    # the cell's own value and returning True/False - "Remarks" (UTR-level
    # bank reconciliation - engine.bank.bank_reconciliation_by_utr) needs
    # the latter: since 2026-08-23 (client-reported, matched against her
    # own corrected working) a tied-out row's Remarks is a composed phrase
    # string (e.g. "Same-period txn settled same period | Ties to bank
    # credit", or several such phrases joined with " + ") rather than the
    # single literal word "Matched" - a plain "Matched" is now only ever a
    # rare degenerate fallback (see _classify_utr_remark's docstring), so
    # an exact-match set here would wrongly flag nearly every reconciled
    # row as an exception.
    ok_values_by_col = {
        "query": {"Okk"},
        # "Query" (capital Q, client-reported 2026-08-30 header rename on
        # the "Reco working" sheet only - see views/page_reports.py's
        # _build_workbook()) - same column, same OK value, just written
        # under the client's own preferred header text.
        "Query": {"Okk"},
        "Reconciliation status": {"Reconciled"},
        "Reconciliation Status": {"Reconciled"},
        "Remarks": lambda v: isinstance(v, str) and v.endswith("Ties to bank credit"),
        "Reconciliation Category": {
            "COD - Not Delivered (No Receipt Expected)",
            "COD - Settlement Received & Bank Matched",
            "Prepaid - Settlement Received & Bank Matched",
        },
        # Amazon Settlement Register (engine.amazon_bank.settlement_register) and
        # Settlement Tie-Out (engine.amazon_invoice_check.settlement_tie_out) both
        # use a plain "status" column - "Nil / Negative" is a normal, expected
        # outcome (nothing due), not an exception, so it's included as "ok" too.
        "status": {"Matched", "Nil / Negative", "Tied out"},
    }
    flag_col_name = None
    for candidate in ("query", "Query", "Reconciliation status", "Reconciliation Status", "Remarks",
                      "Reconciliation Category", "status"):
        if candidate in df.columns:
            flag_col_name = candidate
            break

    if flag_col_name and len(df) <= MAX_STYLED_ROWS:
        ok_values = ok_values_by_col[flag_col_name]
        is_ok = ok_values if callable(ok_values) else (lambda v: v in ok_values)
        q_col_idx = list(df.columns).index(flag_col_name) + 1
        for row_idx, val in enumerate(df[flag_col_name], start=2):
            if not is_ok(val):
                ws.cell(row=row_idx, column=q_col_idx).fill = FLAG_FILL


def style_workbook(writer, sheet_dataframes, header_colors=None):
    """
    sheet_dataframes: dict of {sheet_name: dataframe} already written via
    df.to_excel(writer, sheet_name=..., index=False) BEFORE calling this.

    header_colors: optional {sheet_name: hex_color} to pin a specific
    header colour for a sheet (e.g. matching a related on-page colour).
    Any sheet not named there gets the next colour in a fixed rotating
    palette, cycling if there are more sheets than palette colours - each
    sheet gets its own accent instead of every single sheet in the
    workbook being identically navy, while staying inside one coherent,
    professional palette rather than a random assortment of colours.
    """
    palette = [NAVY, CAT_BLUE, CAT_VIOLET, CAT_AQUA, CAT_AMBER, CAT_ORANGE]
    header_colors = header_colors or {}
    for i, (sheet_name, df) in enumerate(sheet_dataframes.items()):
        ws = writer.sheets[sheet_name]
        color = header_colors.get(sheet_name) or palette[i % len(palette)]
        style_sheet(ws, df, header_color=color)


# ---------------------------------------------------------------------------
# Dashboard build
# ---------------------------------------------------------------------------
# A validated, colourblind-safe palette (see dataviz skill reference): plain
# magnitude figures get a categorical hue, "reconciled vs unreconciled"
# figures get the reserved status colours (good/critical) so a status colour
# never gets reused for an unrelated series.

FONT_NAME = "Calibri"
NAVY = "1A1A2E"
INK_MUTED = "6B6B6B"
INK_SECONDARY = "52514E"
CARD_BORDER = "D8D8D8"

STATUS_GOOD = "0CA30C"      # reconciled
STATUS_CRITICAL = "D03B3B"  # unreconciled
CAT_BLUE = "2A78D6"         # total orders
CAT_AQUA = "1BAF7A"         # gross order value
CAT_AMBER = "EDA100"        # total deduction
CAT_VIOLET = "4A3AA7"       # net settlement
CAT_ORANGE = "EB6834"       # open queries


def _lighten(hex_color, amount=0.86):
    """Blends a hex colour toward white by `amount` (0 = no change, 1 = white) -
    used for the soft KPI-card tint behind a full-strength accent colour."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    r = round(r + (255 - r) * amount)
    g = round(g + (255 - g) * amount)
    b = round(b + (255 - b) * amount)
    return f"{r:02X}{g:02X}{b:02X}"


def _darken(hex_color, amount=0.35):
    """Blends a hex colour toward black by `amount` (0 = no change, 1 =
    black) - the inverse of _lighten() above. Used for the "Reco working"
    sheet's column-name header row (client-reported 2026-08-30, item 3 -
    see style_reco_working_sections()'s own docstring): a darker shade of
    each section's own group-header colour, not a flat one-colour-fits-all
    header, matching the client's own reference workbook's per-section
    colour-coding."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    r = round(r * (1 - amount))
    g = round(g * (1 - amount))
    b = round(b * (1 - amount))
    return f"{r:02X}{g:02X}{b:02X}"


def _fill_block(ws, r1, c1, r2, c2, hex_color):
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).fill = fill


def _box_border(ws, r1, c1, r2, c2, color=CARD_BORDER, weight="thin"):
    """Draws a single box border around the outer edge of a cell block
    (rather than every cell in it), by only assigning border sides that sit
    on the block's perimeter."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                top=Side(style=weight, color=color) if r == r1 else Side(style=None),
                bottom=Side(style=weight, color=color) if r == r2 else Side(style=None),
                left=Side(style=weight, color=color) if c == c1 else Side(style=None),
                right=Side(style=weight, color=color) if c == c2 else Side(style=None),
            )


def _kpi_card(ws, r1, c1, r2, c2, label, value, accent_hex, number_format="#,##0"):
    """One stat tile: a thin accent strip, a muted uppercase label, and a
    big bold accent-coloured number, on a soft tint of the accent colour."""
    _fill_block(ws, r1, c1, r2, c2, _lighten(accent_hex, 0.88))
    _fill_block(ws, r1, c1, r1, c2, accent_hex)  # accent strip along the top
    _box_border(ws, r1, c1, r2, c2, color=accent_hex)

    label_row = r1 + 1
    ws.merge_cells(start_row=label_row, start_column=c1, end_row=label_row, end_column=c2)
    lcell = ws.cell(row=label_row, column=c1, value=label.upper())
    lcell.font = Font(name=FONT_NAME, size=9, bold=True, color=INK_MUTED)
    lcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    value_row = r1 + 2
    ws.merge_cells(start_row=value_row, start_column=c1, end_row=r2, end_column=c2)
    vcell = ws.cell(row=value_row, column=c1, value=value)
    vcell.font = Font(name=FONT_NAME, size=20, bold=True, color=accent_hex)
    vcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    vcell.number_format = number_format


def _section_title(ws, row, col, text, span=6):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    return cell


def _colour_points(series, colors):
    series.data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=c))
        for i, c in enumerate(colors)
    ]


def _find_col_letter(ws, header_name):
    """Looks up a column by its header text in row 1 of an already-written
    sheet, so a formula referencing "whichever column has_settlement_row
    ended up in" doesn't have to hardcode a column letter that would
    silently go stale the moment a source dataframe's column order
    changes."""
    for cell in ws[1]:
        if cell.value == header_name:
            return cell.column_letter
    return None


RECO_WORKING_GROUPS = [
    ("SHOPIFY REPORT", [
        "order_id", "month", "created_at", "financial_status", "fulfillment_status",
        "subtotal", "shipping", "taxes", "total", "payment_method",
    ]),
    # Delivery partner group is completed dynamically in reco_working_layout()
    # below - each uploaded partner's own "<label>_raw_status" column varies
    # run to run depending on which delivery-partner reports were uploaded.
    ("DELIVERY PARTNER REPORT", ["delivery_partner", "final_delivery_status"]),
    ("PAYMENT GATEWAY", [
        "Payment Method", "Payment Provider", "receipt_amount", "total_deduction",
        "refund_amount", "diff", "settlement_amount", "receipt_status", "query",
    ]),
    ("BANK MATCHING", [
        "Payment UTR", "Payment Date (Bank Date)", "Setlment Remarks", "Refund UTR", "Refund Date", "Gateway",
    ]),
]

# Columns reco_df carries internally that the client's own reference
# workbook does NOT show as their own Reco working column - excluded
# outright by reco_working_layout() below, rather than falling through to
# its "unrecognised column -> append at the end" leftover handling (that
# fallback is for a genuine future/unanticipated addition; these are
# already-known internal-only fields, not stray unknowns):
#   - settlement_pending_amount: engine/reco.py::attach_settlement_pending -
#     feeds headline totals/Executive Summary only.
#   - delivered_date / rto_date: kept on reco_df for Order Lookup and other
#     internal use, but not part of the client's own 31-column layout.
RECO_WORKING_INTERNAL_ONLY_COLUMNS = {"settlement_pending_amount", "delivered_date", "rto_date"}


def reco_working_layout(reco_df, delivery_partner_labels=None):
    """
    Returns (ordered_columns, groups) for the Reco working sheet, matching
    the client's own reference workbook layout (client-reported 2026-08-27):
    SHOPIFY REPORT | DELIVERY PARTNER REPORT | PAYMENT GATEWAY | BANK
    MATCHING, each its own merged group-header band (see
    style_reco_working_sections() below) - "ordered_columns" is the exact
    column order to reindex reco_df to before writing it to the sheet, and
    "groups" is [(group_title, first_col_idx, last_col_idx), ...] (1-based,
    inclusive) for the band each group occupies.

    Only columns actually present in reco_df are included. An internal-only
    column such as "settlement_pending_amount" (engine/reco.py::
    attach_settlement_pending - feeds the headline totals/Executive
    Summary, never shown as its own Reco working column in the client's own
    reference workbook) is silently excluded here, not appended as a stray
    extra column. Conversely, any column this function doesn't recognise at
    all (a future addition, or a channel-specific field this layout never
    anticipated) is still appended, after the four known groups, rather
    than silently dropped - a schema change here degrades to "extra column
    at the end", never data loss.

    delivery_partner_labels: optional - a config's own delivery_partners
    order (i.e. [cfg["label"] for cfg in config["delivery_partners"]]).
    Orders each partner's own "<label>_raw_status" column to match delivery
    priority order instead of whatever order they happened to be added to
    reco_df in (dict/column insertion order). Falls back to reco_df's own
    column order when omitted.
    """
    cols = set(reco_df.columns)
    shopify_report_spec, delivery_tail_spec, payment_gateway_spec, bank_matching_spec = (
        RECO_WORKING_GROUPS[0][1], RECO_WORKING_GROUPS[1][1], RECO_WORKING_GROUPS[2][1], RECO_WORKING_GROUPS[3][1],
    )

    shopify_report = [c for c in shopify_report_spec if c in cols]

    raw_status_cols = [c for c in reco_df.columns if c.endswith("_raw_status")]
    if delivery_partner_labels:
        order_key = {
            f"{lbl.lower().replace(' ', '_')}_raw_status": i for i, lbl in enumerate(delivery_partner_labels)
        }
        raw_status_cols = sorted(raw_status_cols, key=lambda c: order_key.get(c, len(order_key)))
    delivery_partner_report = raw_status_cols + [c for c in delivery_tail_spec if c in cols]

    payment_gateway = [c for c in payment_gateway_spec if c in cols]
    bank_matching = [c for c in bank_matching_spec if c in cols]

    ordered = shopify_report + delivery_partner_report + payment_gateway + bank_matching
    known = set(ordered)
    leftover = [
        c for c in reco_df.columns
        if c not in known and c not in RECO_WORKING_INTERNAL_ONLY_COLUMNS
    ]
    ordered += leftover

    groups = []
    pos = 1
    for title, block in [
        ("SHOPIFY REPORT", shopify_report), ("DELIVERY PARTNER REPORT", delivery_partner_report),
        ("PAYMENT GATEWAY", payment_gateway), ("BANK MATCHING", bank_matching),
    ]:
        if block:
            groups.append((title, pos, pos + len(block) - 1))
        pos += len(block)

    return ordered, groups


# Client-reported 2026-08-30 (item 3, Report Format): reverse-engineered
# straight off the client's own reference workbook's actual cell fills -
# each of the 4 "Reco working" section bands gets its OWN colour (from
# this engine's existing CAT_* dashboard palette, not a generic cycle -
# these 4 particular section names always appear in this fixed order, see
# reco_working_layout()), rather than one flat NAVY band for every
# section. The client's own workbook was clearly originally produced by
# this same tool (its exact CAT_BLUE/CAT_AQUA/CAT_VIOLET hex values show
# up unchanged on their group-header row) - this per-section colouring
# was lost somewhere before the current single-flat-NAVY implementation;
# restored here to match.
RECO_WORKING_GROUP_COLORS = {
    "SHOPIFY REPORT": CAT_BLUE,
    "DELIVERY PARTNER REPORT": CAT_AQUA,
    "PAYMENT GATEWAY": CAT_VIOLET,
    "BANK MATCHING": CAT_AMBER,
}


def style_reco_working_sections(ws, groups):
    """
    Adds the client's own 4-band group-header row (SHOPIFY REPORT /
    DELIVERY PARTNER REPORT / PAYMENT GATEWAY / BANK MATCHING - client-
    reported 2026-08-27, "professionally formatted... with clear sections")
    above the column-name header row that style_sheet()/style_workbook()
    already wrote and styled at row 1. Call this AFTER style_workbook() has
    run for every sheet (so the existing header styling, column widths, and
    banded rows are already in place) - groups should be reco_working_layout()'s
    own second return value, computed off the SAME column order the sheet
    was actually written in.

    Inserting exactly one row at the very top (before any merged cells
    exist elsewhere in this sheet) is one of the safe, well-behaved cases
    for openpyxl's insert_rows() - every existing cell (values, fills,
    borders, fonts) shifts down by one row intact, so the column-name
    header row that was row 1 becomes row 2, and data starts at row 3.

    Client-reported 2026-08-30 (item 3): each band now gets its own colour
    (RECO_WORKING_GROUP_COLORS above, matching the client's own reference
    workbook) instead of one flat NAVY fill for every section - and the
    column-name row underneath (now row 2) is re-coloured to a DARKER
    shade of that same section colour (see _darken()), rather than left in
    whatever flat header colour style_workbook() applied before this
    function knew which section each column belonged to. An unrecognised
    group title (custom config, future column) falls back to plain NAVY
    for both rows, never errors.
    """
    if not groups:
        return
    ws.insert_rows(1)
    ws.row_dimensions[1].height = 24
    for title, c1, c2 in groups:
        if c2 < c1:
            continue
        base_color = RECO_WORKING_GROUP_COLORS.get(title, NAVY)
        header_color = _darken(base_color)
        band_fill = PatternFill(start_color=base_color, end_color=base_color, fill_type="solid")
        header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        cell = ws.cell(row=1, column=c1, value=title)
        cell.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(c1, c2 + 1):
            ws.cell(row=1, column=c).fill = band_fill
            # Row 2 (the actual column-name header, after insert_rows shifted
            # it down from row 1) - recolour to this section's darker shade.
            header_cell = ws.cell(row=2, column=c)
            header_cell.fill = header_fill
            header_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    ws.freeze_panes = "A3"


def link_waterfall_formulas(ws, waterfall_df):
    """
    Rewrites the four roll-forward TOTAL rows already written to an
    Amazon "Waterfall" sheet (Net Sales, Order-level Payable, Receivable,
    Balance Receivable) as real Excel formulas that add up the rows sitting
    right above them in this same sheet, instead of the static number
    pandas' to_excel() wrote - and adds a third column explaining how
    every OTHER (base/source) line was actually derived.

    Only these four become live formulas, deliberately: each is a plain
    addition of rows already visible in this sheet, so the formula is
    guaranteed to equal the Python-computed figure exactly - it's the
    identical addition, just expressed as a cell reference instead of a
    hardcoded number, so there's zero risk of ever showing an end user a
    formula that recalculates to a different figure than what the rest of
    the app reports. The base/source lines (Sales as per MTR, Refunds,
    Order/Settlement-level Deductions, Received to date) involve real
    per-row filtering logic in engine/amazon_reco.py and
    engine/amazon_consolidator.py (bucket/category rules, reporting
    cut-off scoping) that a cross-sheet SUMIF could only approximate, so
    those stay the verified static number with a plain-language "how this
    is derived" note pointing at the exact source sheet instead.
    """
    particulars = list(waterfall_df["Particular"])
    row_of = {p: i + 2 for i, p in enumerate(particulars)}  # row 1 is the header

    descriptions = {
        "Sales as per MTR (Invoice Value)":
            "Sum of Invoice Amount minus sum of Refund Amount, across every order on "
            "the 'Order-wise Detail' sheet.",
        "Less: Refunds":
            "Sum of the Refund Amount column on the 'Order-wise Detail' sheet.",
        "Less: Order-level Deductions (Flat File, incl. TDS/TCS)":
            "Sum of every order-level deduction/credit line (incl. TDS/TCS) from the "
            "Settlement Flat File - see each order's own total on the "
            "'Expense Ledger (Order-wise)' sheet.",
        "Less: Settlement-level Deductions (Flat File)":
            "Sum of every settlement-level (non-order) deduction/credit line from the "
            "Settlement Flat File - Advertising, Storage, MCF fees, etc. - see the "
            "'Expense Ledger (Settlement)' sheet.",
        "Less: Received to date (Bank Statement)":
            "Sum of bank credits matched to settlements dated on/before the reporting "
            "cut-off - see the 'Settlement Register' sheet.",
        "(Memo) Net Reserve Movement - cash-flow timing, not in the total above":
            "Net Reserve Movement lines from the Settlement Flat File - a cash-flow "
            "timing item, memo only, not part of the Receivable total above.",
        "(Memo) Non-MTR order-level items (MCF/Shopify pass-through) - not in the total above":
            "Order-level lines in the Settlement Flat File for order-ids that never "
            "appeared in the MTR report(s) - memo only, not part of the total above.",
    }
    rollforward = {
        "Net Sales": ("Sales as per MTR (Invoice Value)", "Less: Refunds"),
        "Order-level Payable": ("Net Sales", "Less: Order-level Deductions (Flat File, incl. TDS/TCS)"),
        "Receivable": ("Order-level Payable", "Less: Settlement-level Deductions (Flat File)"),
        "Balance Receivable": ("Receivable", "Less: Received to date (Bank Statement)"),
    }

    hdr = ws.cell(row=1, column=3, value="How This Is Calculated")
    hdr.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    hdr.fill = copy.copy(ws.cell(row=1, column=2).fill)
    hdr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions["C"].width = 62

    for particular, row in row_of.items():
        note_cell = ws.cell(row=row, column=3)
        note_cell.font = Font(name=FONT_NAME, size=9.5, italic=True, color=INK_MUTED)
        note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if particular in rollforward:
            top_label, delta_label = rollforward[particular]
            top_row, delta_row = row_of.get(top_label), row_of.get(delta_label)
            if top_row and delta_row:
                ws.cell(row=row, column=2, value=f"=B{top_row}+B{delta_row}")
                note_cell.value = f"= B{top_row} ({top_label}) + B{delta_row} ({delta_label})"
        elif particular in descriptions:
            note_cell.value = descriptions[particular]


def _fmt_date(d):
    if d is None:
        return None
    if isinstance(d, str):
        d = pd.to_datetime(d, errors="coerce")
        if pd.isna(d):
            return None
    return pd.Timestamp(d).strftime("%d %b %Y")


def _effective_date_range(reco_df, date_from, date_to):
    """The range to display on the report header: the explicit From/To the
    user picked, falling back to the actual span of dates in the data when
    no explicit range was set (e.g. "All months" with no custom range)."""
    if date_from or date_to:
        lo = _fmt_date(date_from) or "…"
        hi = _fmt_date(date_to) or "…"
        return lo, hi
    if reco_df is not None and "created_at" in reco_df.columns and len(reco_df):
        dates = pd.to_datetime(reco_df["created_at"], errors="coerce").dropna()
        if len(dates):
            return _fmt_date(dates.min()), _fmt_date(dates.max())
    return None, None


def _unreconciled_trend(reco_df):
    """Cumulative unreconciled amount (running total of `diff`, sorted by
    order date) across the report's date span - buckets by day for a
    typical single-month report, widening to weekly/monthly automatically
    for a longer combined-months report so the chart doesn't get crowded."""
    if reco_df is None or "created_at" not in reco_df.columns or len(reco_df) == 0:
        return pd.DataFrame(columns=["bucket_label", "cumulative_diff"])

    dates = pd.to_datetime(reco_df["created_at"], errors="coerce")
    valid = reco_df.loc[dates.notna(), ["diff"]].copy()
    valid["_date"] = dates[dates.notna()]
    if valid.empty:
        return pd.DataFrame(columns=["bucket_label", "cumulative_diff"])

    span_days = (valid["_date"].max() - valid["_date"].min()).days
    freq, label_fmt = (("D", "%d-%b") if span_days <= 45 else
                        ("W", "%d-%b") if span_days <= 180 else
                        ("MS", "%b-%Y"))

    grouped = valid.groupby(pd.Grouper(key="_date", freq=freq))["diff"].sum().reset_index()
    grouped = grouped.sort_values("_date")
    grouped["cumulative_diff"] = grouped["diff"].cumsum().round(2)
    grouped["bucket_label"] = grouped["_date"].dt.strftime(label_fmt)
    return grouped[["bucket_label", "cumulative_diff"]]


def add_dashboard_sheet(writer, totals, month_df, reco_df=None, channel_name=None,
                         date_from=None, date_to=None, generated_at=None):
    """
    Inserts a front-page 'Dashboard' sheet, styled as a management-facing
    MIS reconciliation summary rather than a plain data sheet: a title
    band, channel/date-range/generated-on header, KPI stat cards,
    reconciled-vs-unreconciled bar + donut charts, a key metrics table, and
    a cumulative unreconciled-amount trend line. Placed first so it's the
    first thing anyone sees when they open the file.
    """
    wb = writer.book
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False  # look like a report, not a worksheet

    gross = totals.get("Gross order value", 0) or 0
    receipt = totals.get("Receipt before deduction", 0) or 0   # = "reconciled amount"
    diff = totals.get("Total diff (unreconciled)", 0) or 0     # = "unreconciled amount"; diff + receipt = gross
    deduction = totals.get("Total deduction", 0) or 0
    refund = totals.get("Total refund", 0) or 0                # Settlement = Receipt - Deduction - Refund
    settlement_pending = totals.get("Settlement pending", 0) or 0  # not yet bank-credited (engine/summary.py headline_totals)
    settlement = totals.get("Net settlement", 0) or 0           # already nets out settlement_pending - should tie to Bank Credit
    total_orders = totals.get("Total orders", 0) or 0
    open_queries = totals.get("Open queries (orders)", 0) or 0
    reconciled_orders = max(total_orders - open_queries, 0)
    pct_amount_reconciled = (receipt / gross) if gross else 0
    pct_orders_reconciled = (reconciled_orders / total_orders) if total_orders else 0

    date_lo, date_hi = _effective_date_range(reco_df, date_from, date_to)
    generated_at = generated_at or dt.datetime.now()

    # --- Title band ------------------------------------------------------
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:T1")
    title_cell = ws["A1"]
    title_cell.value = "📊  RECONCILIATION DASHBOARD"
    title_cell.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 1, 1, 1, 20, NAVY)

    # --- Sub-header: channel / date range / generated-on ------------------
    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:T2")
    parts = []
    if channel_name:
        parts.append(f"Channel: {channel_name}")
    if date_lo and date_hi:
        parts.append(f"Date Range: {date_lo} to {date_hi}")
    parts.append(f"Report Generated On: {generated_at.strftime('%d %b %Y, %I:%M %p')}")
    sub_cell = ws["A2"]
    sub_cell.value = "      |      ".join(parts)
    sub_cell.font = Font(name=FONT_NAME, size=10.5, bold=True, color=INK_SECONDARY)
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 2, 1, 2, 20, "F2F2F0")

    # --- KPI stat cards ----------------------------------------------------
    # Row 1: the four headline figures. Row 2: the remaining three from the
    # checklist (Total Deduction, Net Settlement, Open Queries), so every
    # figure the report is meant to lead with is visible above the fold,
    # not just the first four.
    card_row_defs = [
        [
            ("Total Orders", total_orders, CAT_BLUE, "#,##0"),
            ("Gross Order Value", gross, CAT_AQUA, '"₹"#,##0.00'),
            ("Reconciled Amount", receipt, STATUS_GOOD, '"₹"#,##0.00'),
            ("Unreconciled Amount", diff, STATUS_CRITICAL, '"₹"#,##0.00'),
        ],
        [
            ("Total Deduction", deduction, CAT_AMBER, '"₹"#,##0.00'),
            ("Net Settlement", settlement, CAT_VIOLET, '"₹"#,##0.00'),
            ("Open Queries (Orders)", open_queries, CAT_ORANGE, "#,##0"),
        ],
    ]
    card_top = 4
    card_h = 4
    card_w = 4
    gap = 1
    for row_i, cards in enumerate(card_row_defs):
        r1 = card_top + row_i * (card_h + 1)
        r2 = r1 + card_h - 1
        c = 1
        for label, value, accent, numfmt in cards:
            _kpi_card(ws, r1, c, r2, c + card_w - 1, label, value, accent, numfmt)
            c += card_w + gap

    charts_top = card_top + len(card_row_defs) * (card_h + 1) + 1
    charts_row = charts_top

    # --- Chart data area (feeds the charts below) ---------------------------
    # Parked well off to the right (column V+), clear of every card/chart
    # anchored in columns A-T, so a long trend table can never run behind or
    # past a chart. Kept plain/muted since it's a data appendix, not a
    # headline figure - but not hidden, so "Select Data" in Excel still
    # shows real, inspectable cells if anyone needs to check the source.
    data_row = charts_top
    _section_title(ws, data_row, 22, "Chart data (feeds the charts to the left)", span=10)
    ws.cell(row=data_row, column=22).font = Font(name=FONT_NAME, size=9, italic=True, color=INK_MUTED)

    summary_hdr_row = data_row + 1
    ws.cell(row=summary_hdr_row, column=22, value="Category")
    ws.cell(row=summary_hdr_row, column=23, value="Amount")
    summary_rows = [
        ("Gross Order Value", gross),
        ("Reconciled Amount", receipt),
        ("Unreconciled Amount", diff),
    ]
    for i, (cat, val) in enumerate(summary_rows, start=1):
        ws.cell(row=summary_hdr_row + i, column=22, value=cat)
        ws.cell(row=summary_hdr_row + i, column=23, value=val)

    status_amt_hdr_row = summary_hdr_row
    ws.cell(row=status_amt_hdr_row, column=25, value="Status")
    ws.cell(row=status_amt_hdr_row, column=26, value="Amount")
    ws.cell(row=status_amt_hdr_row + 1, column=25, value="Reconciled")
    ws.cell(row=status_amt_hdr_row + 1, column=26, value=receipt)
    ws.cell(row=status_amt_hdr_row + 2, column=25, value="Unreconciled")
    ws.cell(row=status_amt_hdr_row + 2, column=26, value=diff)

    status_cnt_hdr_row = summary_hdr_row
    ws.cell(row=status_cnt_hdr_row, column=28, value="Status")
    ws.cell(row=status_cnt_hdr_row, column=29, value="Orders")
    ws.cell(row=status_cnt_hdr_row + 1, column=28, value="Reconciled")
    ws.cell(row=status_cnt_hdr_row + 1, column=29, value=reconciled_orders)
    ws.cell(row=status_cnt_hdr_row + 2, column=28, value="Unreconciled")
    ws.cell(row=status_cnt_hdr_row + 2, column=29, value=open_queries)

    trend_df = _unreconciled_trend(reco_df)
    trend_hdr_row = summary_hdr_row
    ws.cell(row=trend_hdr_row, column=31, value="Period")
    ws.cell(row=trend_hdr_row, column=32, value="Cumulative unreconciled amount")
    # Only label every Nth point on the x-axis (~10 labels total) - writing
    # blank labels for the skipped points, rather than relying on the
    # chart's own tick-skip setting, so this renders legibly in any viewer,
    # not just ones that honour tickLblSkip.
    label_every = max(1, len(trend_df) // 10)
    for i, r in enumerate(trend_df.itertuples(index=False), start=1):
        show_label = ((i - 1) % label_every == 0) or (i == len(trend_df))
        ws.cell(row=trend_hdr_row + i, column=31, value=r.bucket_label if show_label else "")
        ws.cell(row=trend_hdr_row + i, column=32, value=r.cumulative_diff)

    for row in range(data_row, data_row + max(len(trend_df), 3) + 2):
        for col in (22, 23, 25, 26, 28, 29, 31, 32):
            cell = ws.cell(row=row, column=col)
            if row > data_row and cell.value is not None:
                cell.font = Font(name=FONT_NAME, size=9, color=INK_MUTED)
                if col in (23, 26, 32):
                    cell.number_format = '"₹"#,##0.00'
                elif col == 29:
                    cell.number_format = "#,##0"

    # --- Reconciliation Summary bar chart ----------------------------------
    bar = BarChart()
    bar.type = "col"
    bar.title = "Reconciliation Summary"
    bar.y_axis.title = "Amount (₹)"
    bar.gapWidth = 60
    bar.legend = None
    data = Reference(ws, min_col=23, min_row=summary_hdr_row, max_row=summary_hdr_row + len(summary_rows))
    cats = Reference(ws, min_col=22, min_row=summary_hdr_row + 1, max_row=summary_hdr_row + len(summary_rows))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    _colour_points(bar.series[0], [CAT_AQUA, STATUS_GOOD, STATUS_CRITICAL])
    bar.width, bar.height = 11, 8.5
    ws.add_chart(bar, f"A{charts_row}")

    # --- Reconciliation Status (by amount) donut ---------------------------
    donut_amt = DoughnutChart()
    donut_amt.title = "Reconciliation Status (by Amount)"
    donut_amt.legend.position = "b"
    data = Reference(ws, min_col=26, min_row=status_amt_hdr_row, max_row=status_amt_hdr_row + 2)
    cats = Reference(ws, min_col=25, min_row=status_amt_hdr_row + 1, max_row=status_amt_hdr_row + 2)
    donut_amt.add_data(data, titles_from_data=True)
    donut_amt.set_categories(cats)
    _colour_points(donut_amt.series[0], [STATUS_GOOD, STATUS_CRITICAL])
    donut_amt.dataLabels = DataLabelList()
    donut_amt.dataLabels.showPercent = True
    donut_amt.dataLabels.showVal = False
    donut_amt.dataLabels.showCatName = False
    donut_amt.dataLabels.showSerName = False
    donut_amt.dataLabels.showLegendKey = False
    donut_amt.width, donut_amt.height = 11, 9.5
    ws.add_chart(donut_amt, f"J{charts_row}")

    # --- Order Summary (by count) donut -------------------------------------
    donut_cnt = DoughnutChart()
    donut_cnt.title = "Order Summary (Count)"
    donut_cnt.legend.position = "b"
    data = Reference(ws, min_col=29, min_row=status_cnt_hdr_row, max_row=status_cnt_hdr_row + 2)
    cats = Reference(ws, min_col=28, min_row=status_cnt_hdr_row + 1, max_row=status_cnt_hdr_row + 2)
    donut_cnt.add_data(data, titles_from_data=True)
    donut_cnt.set_categories(cats)
    _colour_points(donut_cnt.series[0], [STATUS_GOOD, STATUS_CRITICAL])
    donut_cnt.dataLabels = DataLabelList()
    donut_cnt.dataLabels.showPercent = True
    donut_cnt.dataLabels.showVal = False
    donut_cnt.dataLabels.showCatName = False
    donut_cnt.dataLabels.showSerName = False
    donut_cnt.dataLabels.showLegendKey = False
    donut_cnt.width, donut_cnt.height = 11, 9.5
    ws.add_chart(donut_cnt, f"Q{charts_row}")

    key_metrics_row = charts_row + 18

    # --- Key Metrics table ---------------------------------------------------
    _section_title(ws, key_metrics_row, 1, "Key Metrics", span=6)
    table_hdr_row = key_metrics_row + 1
    for col, text in [(1, "Metric"), (5, "Value")]:
        cell = ws.cell(row=table_hdr_row, column=col, value=text)
        ws.merge_cells(start_row=table_hdr_row, start_column=col,
                        end_row=table_hdr_row, end_column=col + (3 if col == 1 else 1))
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # 4th tuple element (when not None) is this row's key in the `totals`
    # dict passed in above - captured into key_metrics_cell_map below so
    # add_executive_summary_sheet() can reference "Dashboard!E<row>" for a
    # given headline figure via a real Excel formula, without hardcoding a
    # row number that would silently go stale if this table's shape ever
    # changes - same "look it up, don't hardcode it" spirit as
    # _find_col_letter() above, just row-based and resolved once here
    # rather than by searching label text (which would be fragile against
    # this table's own em-dash/parenthetical label wording).
    key_metric_rows = [
        ("Total Orders", total_orders, "#,##0", "Total orders"),
        ("Gross Order Value", gross, '"₹"#,##0.00', "Gross order value"),
        ("Receipt Before Deduction (Reconciled Amount)", receipt, '"₹"#,##0.00', "Receipt before deduction"),
        ("Total Deduction", deduction, '"₹"#,##0.00', "Total deduction"),
        ("Total Refund", refund, '"₹"#,##0.00', "Total refund"),
        ("Settlement Pending (not yet bank-credited)", settlement_pending, '"₹"#,##0.00', "Settlement pending"),
        # Client-reported 2026-08-31 (item 8): label text only - the VALUE
        # was already "nets out settlement_pending" (see engine/summary.py::
        # headline_totals()'s own docstring for the fix), but the old label
        # implied a literal Receipt-Deduction-Refund-SettlementPending
        # subtraction of the FOUR OTHER VISIBLE LINES above, which no
        # longer holds now that the deduction is capped per-order (a
        # pending order that never had anything counted in Receipt to
        # begin with correctly contributes 0, not its full order value) -
        # worded so a reviewer checking this line against the visible
        # Settlement Pending total above isn't misled into expecting an
        # exact match.
        ("Net Settlement (Receipt − Deduction − Refund, excluding orders still Settlement Pending)",
         settlement, '"₹"#,##0.00', "Net settlement"),
        ("Total Diff / Unreconciled Amount", diff, '"₹"#,##0.00', "Total diff (unreconciled)"),
        ("Open Queries / Unreconciled Orders", open_queries, "#,##0", "Open queries (orders)"),
        ("Reconciliation % (by Amount)", pct_amount_reconciled, "0.00%", "Reconciliation % (by amount)"),
        ("Reconciliation % (by Orders)", pct_orders_reconciled, "0.00%", "Reconciliation % (by orders)"),
    ]
    key_metrics_cell_map = {}
    for i, (label, value, numfmt, totals_key) in enumerate(key_metric_rows, start=1):
        r = table_hdr_row + i
        band = "F7F7F5" if i % 2 else "FFFFFF"
        _fill_block(ws, r, 1, r, 6, band)
        lcell = ws.cell(row=r, column=1, value=label)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        lcell.font = Font(name=FONT_NAME, size=10, color=INK_SECONDARY)
        lcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vcell = ws.cell(row=r, column=5, value=value)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        vcell.font = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
        vcell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        vcell.number_format = numfmt
        if totals_key:
            key_metrics_cell_map[totals_key] = r
    _box_border(ws, table_hdr_row, 1, table_hdr_row + len(key_metric_rows), 6)

    # --- Unreconciled Amount Trend line chart -------------------------------
    if len(trend_df) > 0:
        line = LineChart()
        line.title = "Unreconciled Amount Trend (cumulative)"
        line.y_axis.title = "Cumulative Amount (₹)"
        line.x_axis.title = "Period"
        line.legend = None
        data = Reference(ws, min_col=32, min_row=trend_hdr_row, max_row=trend_hdr_row + len(trend_df))
        cats = Reference(ws, min_col=31, min_row=trend_hdr_row + 1, max_row=trend_hdr_row + len(trend_df))
        line.add_data(data, titles_from_data=True)
        line.set_categories(cats)
        s = line.series[0]
        s.graphicalProperties.line.solidFill = STATUS_CRITICAL
        s.graphicalProperties.line.width = 22000  # EMUs (~1.7pt)
        s.marker = Marker(symbol="circle", size=5)
        s.marker.graphicalProperties.solidFill = STATUS_CRITICAL
        s.marker.graphicalProperties.line.solidFill = STATUS_CRITICAL
        s.smooth = False
        # Thin out x-axis labels so a daily-bucketed month (~30 points)
        # doesn't render as an illegible wall of overlapping dates.
        line.x_axis.tickLblSkip = max(1, len(trend_df) // 10)
        line.width, line.height = 15, 8.5
        ws.add_chart(line, f"H{key_metrics_row}")

    # --- Footer -------------------------------------------------------------
    footer_row = key_metrics_row + len(key_metric_rows) + 3
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=12)
    fcell = ws.cell(row=footer_row, column=1,
                     value="This is an auto-generated report. For any queries, please contact the Finance Team.")
    fcell.font = Font(name=FONT_NAME, size=9, italic=True, color=INK_MUTED)
    fcell.alignment = Alignment(horizontal="left", vertical="center")

    # --- Column widths / page setup -----------------------------------------
    for letter in "ABCDEFGHIJKLMNOPQRST":
        ws.column_dimensions[letter].width = 13

    # Print/PDF only the dashboard itself (cols A-T) - not the chart-data
    # appendix parked out at column V+, which would otherwise get squeezed
    # onto the same page and defeat the "clean single view" point of this.
    ws.print_area = f"A1:T{footer_row}"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return key_metrics_cell_map


def _reco_col_letter(reco_working_cols, name):
    """Column letter for `name` within the Reco working sheet's own column
    order (reco_working_layout()'s first return value) - resolved once in
    Python at build time, since that order is already known there (unlike
    _find_col_letter() above, no need to search the worksheet itself).
    Returns None if the column isn't present (e.g. no bank statement this
    run, so the Bank Matching group's columns were never attached)."""
    if name not in reco_working_cols:
        return None
    return get_column_letter(reco_working_cols.index(name) + 1)


def apply_settlement_pending_summary_formulas(writer, reco_working_cols, gateway_configs):
    """
    Client-reported 2026-08-31 (round 4, point 7): "Executive Summary"
    Section 3 (a pure row-by-row mirror of the "Settlement Pending
    Summary" sheet - see that section's own comment in
    add_executive_summary_sheet() below) kept showing small differences
    against Reco working even after engine.settlement_pending.
    settlement_pending_summary_by_gateway() was rebuilt to mirror Reco
    working's own columns exactly (2026-08-31, item 1). Root cause: BOTH
    sheets were still Python-computed STATIC values written once at
    export time - correct for that one instant, but with nothing tying
    them back to Reco working's own cells the way Section 7 above already
    ties itself to other sheets via live formulas. The client asked
    explicitly for the Settlement Pending Summary sheet itself to be
    linked to Reco working via formulas, "so that the figures are
    directly linked and there is better visibility and consistency
    between both sheets" - exactly the same live-formula pattern already
    used elsewhere in this workbook.

    Called AFTER both 'Reco working' and 'Settlement Pending Summary' are
    already written (write_workbook_sheets()/style_workbook() in
    views/page_reports.py::_build_workbook()) - this does NOT change
    which (Group, Payment Gateway) rows exist on the sheet (still exactly
    whatever settlement_pending_summary_by_gateway() decided - discovering
    which gateways currently have pending orders at all isn't something a
    static Excel range can do on its own). It only REPLACES the already-
    written "Orders Pending"/"Amount Pending" VALUE cells (columns C/D)
    of each row with COUNTIFS/SUMIFS formulas that reproduce that
    function's exact business rule directly off Reco working's own
    columns, so a later edit to Reco working (or a re-run for a different
    date range that reuses this same sheet layout) recalculates this
    sheet - and, through Section 3's existing row mirror, Executive
    Summary too - with no separate Python recomputation able to drift out
    of sync:
      - the unresolved-gateway catch-all row ("COD Delivered Amount not
        Received"): SUMIFS/COUNTIFS of Total where Recipt Remark = "Not
        Received" AND Query = "COD Delivered Amount not Received" (that
        Query text is, by construction of
        engine.reco.refine_queries_with_settlement_status(), synonymous
        with "no resolvable Gateway" for a still-pending order).
      - a resolved COD courier row (Delhivery COD / Shiprocket COD /
        Prozo COD, from the client's own gateway config's payment_mode):
        the SUM of two SUMIFS - receipt_amount for that courier's own
        non-catch-all pending rows, PLUS Total for the handful of rows
        attributed to that courier via the Gateway column even though
        their Query is still the generic catch-all (the same "fold-in"
        case settlement_pending_summary_by_gateway()'s own docstring
        describes) - and the matching COUNTIFS pair for Orders Pending.
      - any other resolved gateway (Payu/Easebuzz/Gokwik/Razorpay/...):
        SUMIFS/COUNTIFS of Total where Recipt Remark = "Not Received" AND
        Gateway = that label.
    Skipped entirely (both sheets left exactly as write_workbook_sheets()
    wrote them) if either sheet is missing, or if any of the Reco working
    columns this needs (Gateway/Recipt Remark/Query/receipt_amount/total)
    aren't present in reco_working_cols - the same safe-degrade already
    used throughout this module rather than raising on an older/partial
    saved run.
    """
    ws = writer.sheets.get("Settlement Pending Summary")
    reco_ws = writer.sheets.get("Reco working")
    if ws is None or reco_ws is None:
        return

    gateway_col = _reco_col_letter(reco_working_cols, "Gateway")
    remark_col = _reco_col_letter(reco_working_cols, "receipt_status")
    query_col = _reco_col_letter(reco_working_cols, "query")
    receipt_col = _reco_col_letter(reco_working_cols, "receipt_amount")
    total_col = _reco_col_letter(reco_working_cols, "total")
    if not all([gateway_col, remark_col, query_col, receipt_col, total_col]):
        return

    last_row = max(reco_ws.max_row, 2)
    rng = lambda c: f"'Reco working'!${c}$2:${c}${last_row}"
    gw_rng, remark_rng, query_rng = rng(gateway_col), rng(remark_col), rng(query_col)
    receipt_rng, total_rng = rng(receipt_col), rng(total_col)

    cod_labels = {cfg["label"] for cfg in (gateway_configs or []) if str(cfg.get("payment_mode", "")).strip().lower() == "cod"}
    catch_all = "COD Delivered Amount not Received"
    # 2026-09-04 (round 10): engine.reco.attach_receipt_status() now also
    # returns "Received Bank settlement pending" (not just "Not Received")
    # for a still-settlement-pending order whose receipt_amount is ALREADY
    # populated (a COD order the courier's own report already confirms
    # collected, just not yet bank-credited - see that function's own
    # docstring, point 2, and engine.settlement_pending.settlement_
    # pending_summary_by_gateway()'s matching fix). SUMIFS/COUNTIFS have no
    # native OR across criteria values, so every formula below is now the
    # SUM of one term per Recipt Remark value in _PENDING_REMARK_VALUES,
    # rather than a single term hardcoded to "Not Received" alone - an
    # order with the new remark would otherwise silently vanish from this
    # sheet's live Excel totals even though the Python-side summary
    # (settlement_pending_summary_by_gateway) already counts it.
    remark_values = ["Not Received", "Received Bank settlement pending"]

    def _sumifs_over_remarks(value_rng, extra_criteria=""):
        terms = []
        for rv in remark_values:
            terms.append("SUMIFS(" + value_rng + "," + remark_rng + ',"' + rv + '"' + extra_criteria + ")")
        return "+".join(terms)

    def _countifs_over_remarks(extra_criteria=""):
        terms = []
        for rv in remark_values:
            terms.append("COUNTIFS(" + remark_rng + ',"' + rv + '"' + extra_criteria + ")")
        return "+".join(terms)

    # 2026-09-06 (round 22, client-reported, order #29284): wildcard
    # criteria matching engine.settlement_pending.NOT_REFLECTING_LABEL's
    # own COD_REPORT_ABSENT_QUERY_FRAGMENT ("delivered but amount not
    # reflecting in") - Excel's SUMIFS/COUNTIFS treat "*text*"/"<>*text*"
    # criteria as contains/does-not-contain, so this reproduces
    # _is_cod_report_absent_query()'s own substring rule directly off the
    # Reco working sheet's Query column, off the exact same fragment
    # constant (never a second, hand-typed copy of the phrase). No Gateway
    # constraint is needed - this phrasing is only ever produced by
    # engine.reco.py::apply_cod_report_gap_query() for an order already
    # confirmed to be a recognised-COD-courier delivery, so matching on
    # Query text alone can't accidentally pull in an unresolved or Prepaid
    # order.
    gap_wildcard = "*" + COD_REPORT_ABSENT_QUERY_FRAGMENT + "*"
    gap_query_excl = "," + query_rng + ',"<>' + gap_wildcard + '"'
    # 2026-09-06 (round 22): the BROAD "not reflecting" wildcard (both
    # phrasings) - used, for a COD courier's OWN row, to separate a
    # genuinely-reflecting pending order (receipt_amount is meaningful)
    # from one where NOTHING has reflected anywhere yet (Total is the only
    # meaningful figure) - mirrors pending_amount_by_order()'s own
    # is_catch_all_query rule (engine/settlement_pending.py), which this
    # live-formula rewrite must match or silently disagree with the
    # already-correct Python-computed Dashboard/Executive Summary total.
    # Confirmed via a genuine LibreOffice recalc (not just "the formula
    # parses") while verifying this round's new NOT_REFLECTING_LABEL
    # split - a courier whose only pending order used this generic
    # phrasing (e.g. order #30456-style "Setlment pending not reflecting")
    # was otherwise silently summed via receipt_amount (0), undercounting
    # against the Python side's Total-based figure - a pre-existing gap in
    # this live-formula rewrite, not introduced by this round's own
    # NOT_REFLECTING_LABEL split, but caught by this round's own
    # verification pass and fixed alongside it rather than left disclosed-
    # but-broken now that it's been directly observed.
    broad_wildcard = "*" + BROAD_NOT_REFLECTING_FRAGMENT + "*"

    for r in range(2, ws.max_row + 1):
        label = ws.cell(row=r, column=2).value  # Payment Gateway
        if not label:
            continue
        lit = str(label).replace('"', '""')
        if label == catch_all:
            amount_f = "=" + _sumifs_over_remarks(total_rng, "," + query_rng + ',"' + lit + '"')
            count_f = "=" + _countifs_over_remarks("," + query_rng + ',"' + lit + '"')
        elif label == NOT_REFLECTING_LABEL:
            amount_f = "=" + _sumifs_over_remarks(total_rng, "," + query_rng + ',"' + gap_wildcard + '"')
            count_f = "=" + _countifs_over_remarks("," + query_rng + ',"' + gap_wildcard + '"')
        elif label in cod_labels:
            # 2026-09-06 (round 22): three terms, a strict partition of
            # this courier's own pending orders (excluding the new
            # NOT_REFLECTING_LABEL row's gap-specific orders entirely -
            # see gap_query_excl above, so an order like #29284 is never
            # double-counted between this row and that one):
            #   1. ordinary reflecting (query doesn't match "not
            #      reflecting" at all) - receipt_amount is meaningful.
            #   2. the OTHER, broader "not reflecting" phrasing (e.g.
            #      "Setlment pending not reflecting") - nothing has
            #      reflected anywhere, so Total is the only meaningful
            #      figure, exactly like pending_amount_by_order()'s own
            #      is_catch_all_query rule.
            #   3. the literal generic catch_all text folded in via the
            #      Gateway column (pre-existing "fold-in" case) - Total.
            amount_f = (
                "=" + _sumifs_over_remarks(receipt_rng, "," + gw_rng + ',"' + lit + '",' + query_rng + ',"<>' + catch_all + '"' + gap_query_excl + "," + query_rng + ',"<>' + broad_wildcard + '"')
                + "+" + _sumifs_over_remarks(total_rng, "," + gw_rng + ',"' + lit + '",' + query_rng + ',"<>' + catch_all + '"' + gap_query_excl + "," + query_rng + ',"' + broad_wildcard + '"')
                + "+" + _sumifs_over_remarks(total_rng, "," + gw_rng + ',"' + lit + '",' + query_rng + ',"' + catch_all + '"')
            )
            count_f = (
                "=" + _countifs_over_remarks("," + gw_rng + ',"' + lit + '",' + query_rng + ',"<>' + catch_all + '"' + gap_query_excl + "," + query_rng + ',"<>' + broad_wildcard + '"')
                + "+" + _countifs_over_remarks("," + gw_rng + ',"' + lit + '",' + query_rng + ',"<>' + catch_all + '"' + gap_query_excl + "," + query_rng + ',"' + broad_wildcard + '"')
                + "+" + _countifs_over_remarks("," + gw_rng + ',"' + lit + '",' + query_rng + ',"' + catch_all + '"')
            )
        else:
            amount_f = "=" + _sumifs_over_remarks(total_rng, "," + gw_rng + ',"' + lit + '"')
            count_f = "=" + _countifs_over_remarks("," + gw_rng + ',"' + lit + '"')
        ws.cell(row=r, column=3, value=count_f)
        ws.cell(row=r, column=4, value=amount_f)


def add_executive_summary_sheet(writer, key_metrics_cell_map, reco_working_cols,
                                 open_queries_df, settlement_pending_summary_df, gateway_recon_df,
                                 reco_df=None, utr_bank_reco_df=None,
                                 has_bank_reco=False, channel_name=None, date_from=None, date_to=None,
                                 generated_at=None, period_bank_credit_total=None):
    """
    Inserts an "Executive Summary" sheet (client-reported 2026-08-27,
    matching their own hand-built reference workbook's sections - six
    originally, a seventh "Receivable Collection Period Analysis" added
    2026-08-31, point 7) as a
    management-facing companion to the Dashboard sheet - built almost
    entirely from live Excel formulas referencing other sheets in this same
    workbook (Dashboard's own Key Metrics table, Reco working, Open
    queries, Settlement Pending Summary, Gateway Settlement, Bank Reco
    (UTR-wise)), the same way the client's own Executive Summary is built -
    so editing any of those source sheets (or re-running this tool for a
    different date range) updates every figure here automatically, without
    a separate refresh step.

    Must be called AFTER add_dashboard_sheet() (needs its returned
    key_metrics_cell_map - see that function's own docstring) and after
    every other sheet this one cross-references has already been written
    to `writer` (so the formulas below resolve against real, already-
    populated ranges) - see views/page_reports.py::_build_workbook() for
    the call order.

    key_metrics_cell_map: add_dashboard_sheet()'s return value - {totals
        dict key -> Dashboard row number}, e.g. key_metrics_cell_map["Net
        settlement"] -> the row whose column E holds that figure.
    reco_working_cols: reco_working_layout()'s first return value - the
        exact column order "Reco working" was actually written in, so
        SUMIFS/COUNTIFS criteria/sum ranges below reference the right
        column regardless of how many delivery-partner raw-status columns
        happened to exist this run.
    open_queries_df / settlement_pending_summary_df / gateway_recon_df:
        the SAME dataframes already written to their own sheets this run -
        read here only to know how many rows/which distinct groups exist
        (row COUNTS, never their values - the actual figures shown always
        come from a live formula into the sheet itself, never a Python-
        computed number pasted in), so this section is sized to what's
        actually in the report, not a hardcoded guess.
    has_bank_reco: whether the "Bank Reco (UTR-wise)" sheet was actually
        written this run (utr_bank_reco_df non-empty) - Total Bank Credit/
        the Cashflow Waterfall's own Bank Credit step fall back to a
        clearly-labelled "Not available - no bank statement uploaded" text
        instead of a formula into a sheet that doesn't exist.
    reco_df (client-reported 2026-08-30, item 12): the same reco_df already
        written as "Reco working" this run - used ONLY for two edge cases
        the rest of this sheet's live-formula design can't reach cleanly:
        (a) Section 2's exact list of distinct (receipt_status, query) row
        PAIRS to write (their VALUES are still live SUMIFS/COUNTIFS
        formulas below, same convention as open_queries_df etc. above -
        only which rows exist is decided in Python), and (b) Section 5's
        "#N/A" (unattributed Payment Provider) row, whose 4 non-gateway-
        recon figures need to match against BLANK cells - unsafe to do
        with a live whole-column COUNTIF/SUMIFS (an empty-string/blank
        criteria against a full 'Reco working'!R:R column reference would
        also match every unused cell below the data, wildly overcounting),
        so those 4 cells are Python-computed literals for this one row
        only; every other row and every other cell in the sheet stays a
        live formula. None-safe: if not provided, Section 2 shows an
        empty-state row and Section 5 skips the "#N/A" row.
    utr_bank_reco_df (client-reported 2026-08-31, item 4/point 7): the
        same "Bank Reco (UTR-wise)" dataframe already written this run -
        read here only for its distinct "Payment Gateway" labels (which
        gateway/courier rows Section 7 below needs), never its values;
        the actual figures are always a live SUMIF into that sheet.
        None-safe: if not provided (or has_bank_reco is False), Section 7
        shows the same "Not available - no bank statement uploaded" text
        as the rest of this sheet's bank-dependent figures.
    period_bank_credit_total (added 2026-08-31, client-reported - Section
        7's own explicit reconciliation requirement: "the total of the
        first three columns should match the actual Bank Credit/Receipt
        amount for the selected period... This should match the bank
        statement exactly"): the selected period's own uploaded bank
        statement credit total, filtered to [period_start_date,
        period_end_date] and computed in views/page_reports.py - a plain
        Python number, not a live formula, because the RAW bank statement
        isn't written to its own sheet anywhere in this workbook (there is
        nothing for a formula to point at) - same reasoning Section 5's
        "#N/A" row already uses for its own few Python-computed cells (see
        the reco_df parameter note above). None-safe: if not provided,
        Section 7's reconciliation check row shows a plain "not available"
        note instead of a formula, rather than a wrong or blank number.
    """
    wb = writer.book
    ws = wb.create_sheet("Executive Summary", 1)  # right after Dashboard
    ws.sheet_view.showGridLines = False

    def dash_cell(totals_key):
        """Bare 'Dashboard!E<row>' reference for a headline_totals() key
        (see key_metrics_cell_map's own docstring above), or None if that
        row wasn't found (e.g. an older code path that skipped a metric)."""
        row = key_metrics_cell_map.get(totals_key)
        return f"Dashboard!E{row}" if row else None

    def dash_ref(totals_key, fallback="0"):
        """Full '=Dashboard!E<row>' formula, or a plain fallback value."""
        cell = dash_cell(totals_key)
        return f"={cell}" if cell else fallback

    # --- Title band --------------------------------------------------------
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "📋  EXECUTIVE SUMMARY"
    title_cell.font = Font(name=FONT_NAME, size=17, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 1, 1, 1, 7, NAVY)

    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:G2")
    date_lo, date_hi = _effective_date_range(None, date_from, date_to)
    parts = []
    if channel_name:
        parts.append(f"Channel: {channel_name}")
    if date_lo and date_hi:
        parts.append(f"Date Range: {date_lo} to {date_hi}")
    parts.append(f"Report Generated On: {(generated_at or dt.datetime.now()).strftime('%d %b %Y, %I:%M %p')}")
    sub_cell = ws["A2"]
    sub_cell.value = "      |      ".join(parts)
    sub_cell.font = Font(name=FONT_NAME, size=10.5, bold=True, color=INK_SECONDARY)
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 2, 1, 2, 7, "F2F2F0")

    row = 4

    def section_header(text):
        nonlocal row
        _section_title(ws, row, 1, text, span=7)
        _fill_block(ws, row, 1, row, 7, "EDEEF7")
        row += 1

    def table_header(cols):
        nonlocal row
        for c, text in enumerate(cols, start=1):
            cell = ws.cell(row=row, column=c, value=text)
            cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
            cell.font = Font(name=FONT_NAME, size=9.5, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        row += 1

    def data_row(values, numfmts=None, band=None, left_cols=None):
        # left_cols (2026-09-06, round 21, client-reported): 1-indexed
        # column numbers to left-align in ADDITION to column 1 (already
        # always left) - every other column defaults to right, correct for
        # this sheet's usual money/count value columns, but wrong for a
        # text label column (e.g. "Final Delivery Status", "Query",
        # "Query / Remark") sharing a row with those. Column 1 needs no
        # entry here (already left unconditionally); every other caller
        # that doesn't pass left_cols keeps the exact old behaviour.
        nonlocal row
        left_cols = left_cols or set()
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.font = Font(name=FONT_NAME, size=10, color=INK_SECONDARY if c == 1 else NAVY)
            horiz = "left" if (c == 1 or c in left_cols) else "right"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", indent=1)
            if numfmts and c - 1 < len(numfmts) and numfmts[c - 1]:
                cell.number_format = numfmts[c - 1]
            if band:
                cell.fill = PatternFill(start_color=band, end_color=band, fill_type="solid")
        row += 1

    money_fmt = '"₹"#,##0.00'
    count_fmt = "#,##0"

    # ------------------------------------------------------------------
    # 1. Headline Numbers
    # ------------------------------------------------------------------
    section_header("1. Headline Numbers")
    table_header(["Metric", "", "", "", "Value", "", ""])
    ws.merge_cells(start_row=row - 1, start_column=1, end_row=row - 1, end_column=4)
    ws.merge_cells(start_row=row - 1, start_column=5, end_row=row - 1, end_column=7)
    # Client-reported 2026-08-30 (item 12): re-verified directly against a
    # LIVE FORMULA in the client's own reference workbook (not just a
    # value comparison) - MOD.xlsx's own "Total Bank Credit" cell is
    # literally =SUM('Bank Reco (UTR-wise)'!D:E), i.e. this period's
    # bank-matched settlement (Same Month + Next Month columns - see
    # engine.bank.bank_reconciliation_by_utr), NOT this engine's own
    # "Net settlement" headline figure (a different computation path -
    # Reco working's receipt-deduction-refund, netted by settlement_
    # pending_amount). An earlier fix here (2026-08-30) approximated Total
    # Bank Credit as Net Settlement on the assumption the two would tie
    # out closely enough - true in aggregate, but not the client's actual
    # formula; replaced with the literal replica now that the exact
    # formula text has been confirmed. Excludes 'Bank Reco (UTR-wise)'
    # columns F/H/I (other-period / order-not-found settlement) - same
    # period-scoping already used elsewhere on this sheet (see the old
    # Section 5 comment this replaced).
    bank_credit_formula = (
        "=SUM('Bank Reco (UTR-wise)'!D:E)" if has_bank_reco
        else "Not available - no bank statement uploaded"
    )
    headline_rows = [
        ("Total Order Value", dash_ref("Gross order value"), money_fmt),
        ("Total Expected Collection", dash_ref("Gross order value"), money_fmt),
        ("Total Bank Credit", bank_credit_formula, money_fmt if has_bank_reco else None),
        ("Reconciled Amount", dash_ref("Receipt before deduction"), money_fmt),
        # Client-reported 2026-08-30 (item 12): label matched verbatim to
        # the client's own reference workbook - a plain "Unreconciled
        # Amount" read as ambiguous on its own.
        ('Unreconciled Amount (Unrealised value "Order Value less receipts")',
         dash_ref("Total diff (unreconciled)"), money_fmt),
        ("Settlement Pending Amount", dash_ref("Settlement pending"), money_fmt),
        # Client-reported 2026-08-30 (item 4): "Open queries" is now 7
        # columns wide (Query|orders|order_value|Recipt|Refund|Bank
        # receipt|exposure - matching the client's own reference workbook,
        # see engine/summary.py::open_queries()), so "exposure" moved from
        # column E to column G.
        ("Open Query Amount", "=SUM('Open queries'!G:G)", money_fmt),
        ("Number of Orders / Transactions", dash_ref("Total orders", fallback="0"), count_fmt),
    ]
    # Value lives in column 5 (E), matching the "Value" header written at
    # E5 above - NOT column 2. Merging A:D for the label and E:G for the
    # value happens AFTER data_row() writes both cells; merging a range
    # that included the value's own cell (e.g. merging A:D when the value
    # was written to column 2, inside that same range) would silently wipe
    # it - openpyxl's merge_cells() clears every cell in a merge range
    # except the top-left anchor. Column 5 sits outside the A:D label
    # merge, so this ordering is safe.
    headline_start = row
    for i, (label, value, numfmt) in enumerate(headline_rows):
        band = "F7F7F5" if i % 2 else "FFFFFF"
        data_row([label, None, None, None, value, None, None], [None, None, None, None, numfmt], band=band)
        ws.merge_cells(start_row=row - 1, start_column=1, end_row=row - 1, end_column=4)
        ws.merge_cells(start_row=row - 1, start_column=5, end_row=row - 1, end_column=7)
    _box_border(ws, headline_start - 1, 1, row - 1, 7)
    row += 1

    # ------------------------------------------------------------------
    # 2. Where the Unreconciled Gap Sits - by Receipt Status x Query
    # ------------------------------------------------------------------
    # Client-reported 2026-08-30 (item 12): rebuilt against the client's
    # own reference workbook's ACTUAL Executive Summary (reverse-engineered
    # from its live formulas in MOD.xlsx, not from an earlier untested
    # assumption) - grouped by "Recipt Remark" (receipt_status) CROSSED
    # with "Query", not by delivery status at all. Every row's two figures
    # are live formulas (SUMIFS/COUNTIFS against Reco working's "diff"
    # column, X=Recipt Remark, Y=Query - column letters resolved
    # dynamically via _reco_col_letter, not hardcoded); only the row
    # IDENTITIES (which (receipt_status, query) pairs exist, and in what
    # order) are decided here in Python from reco_df, same convention as
    # open_queries_df/settlement_pending_summary_df above - orders whose
    # receipt_status is "Received" or "Refunded" (i.e. already fully
    # reconciled) are excluded, matching MOD's own row list exactly.
    # Client-reported 2026-09-05 (points 4/5): header renamed to "...by
    # Delivery and receipt Status" (this section was always grouped by
    # Recipt Remark, i.e. receipt status, crossed with Query - the header
    # text just hadn't caught up), and a new "Payment Provider" column
    # added next to "Recipt Remark". Since every row here is itself a
    # GROUPED summary (by Recipt Remark + Query, not one row per order),
    # Payment Provider is added as a THIRD grouping key rather than a
    # lookup against an already-aggregated row - that's the only way each
    # row still resolves to exactly one provider value instead of a
    # blended/ambiguous one, and it keeps this section's existing
    # exact-match SUMPRODUCT convention intact (just one more EXACT()
    # criterion). A blank/unresolved Payment Provider groups under
    # "Unresolved / #N/A", matching Section 5's own convention for the
    # same situation below.
    #
    # Client-reported 2026-09-05 (follow-up, point 2): a second new column,
    # "Final Delivery Status", added next to "Recipt Remark" too (i.e.
    # BEFORE Payment Provider) - same treatment, a fourth grouping key.
    # final_delivery_status is always a real string on every order (see
    # engine.reco.attach_delivery_status - it defaults to "Status
    # Undefined", never blank/NaN), so unlike Payment Provider there's no
    # "Unresolved / #N/A" relabelling needed here - just fillna as a safety
    # net for an older saved reco_df that predates this column.
    #
    # Both new columns are entirely optional/independent (a reco_df
    # missing one still gets the other), so the column layout - and the
    # column letter that holds "Unreconciled Amount" - shifts depending on
    # which are present. gap_amount_col below is computed to match
    # whatever was ACTUALLY written, and Section 6's Cashflow Waterfall
    # (the only other place that reads this section's own layout) uses
    # that same computed letter instead of a hardcoded one - see that
    # section's own comment.
    section_header("2. Where the Unreconciled Gap Sits (by Delivery and receipt Status)")
    recipt_col = _reco_col_letter(reco_working_cols, "receipt_status")
    query_col = _reco_col_letter(reco_working_cols, "query")
    diff_col = _reco_col_letter(reco_working_cols, "diff")
    provider_col_s2 = _reco_col_letter(reco_working_cols, "Payment Provider")
    status_col_s2 = _reco_col_letter(reco_working_cols, "final_delivery_status")
    gap_pairs_df = None
    gap_has_provider = False
    gap_has_status = False
    gap_amount_col = None
    if (reco_df is not None and not reco_df.empty and recipt_col and query_col and diff_col
            and "receipt_status" in reco_df.columns and "query" in reco_df.columns):
        # 2026-08-31 (round 6) fix - two client-reported defects, both
        # traced to the SAME root cause: engine/reco.py::attach_receipt_
        # status() can legitimately produce two receipt_status values that
        # differ ONLY in case for two genuinely different situations - the
        # generic "still Settlement Pending" bucket reads "Not Received"
        # (capital R), while an order whose delivery status itself could
        # never be classified reads the client's own literal wording, "Not
        # received" (lowercase r, client-specified 2026-08-31 round 4).
        # Both are correct and intentionally distinct in Python. The bug:
        # SUMIFS/COUNTIFS (used below to fill in each row's live Amount/
        # Orders figures) are CASE-INSENSITIVE in both Excel and
        # LibreOffice - so a "Not Received" row's formula and a "Not
        # received" row's formula each silently matched BOTH sets of rows
        # combined, showing the identical (inflated) total on both lines
        # and double-counting that amount into every total this section
        # feeds (Point 2's own total, and the Cashflow Waterfall's
        # "Unreconciled Gap" step below, which SUMs this section's own
        # column - client-reported: table showed Rs 6,45,919.96 against an
        # actual Rs 6,07,221.59, a gap of exactly Rs 38,698.37 - the size
        # of the duplicated "COD Delivery status Undifined  Amount not
        # Received" group). Fixed by switching these two formulas from
        # SUMIFS/COUNTIFS to SUMPRODUCT(EXACT(...)), which compares
        # case-SENSITIVELY - each row now sums only its own exact-cased
        # group, and the two legitimately-distinct rows finally get their
        # own correct, non-overlapping totals instead of an identical
        # inflated one.
        #
        # Second, separate fix bundled here (client-reported 2026-08-31,
        # round 6): "If the unreconciled amount is zero, that delivery
        # status should not appear in this table" - a (receipt_status,
        # query) group can legitimately sum to zero (e.g. a COD order
        # already fully accounted for at the Order-Value-less-Receipts
        # level, purely awaiting BANK settlement rather than awaiting
        # collection at all - a different, unrelated metric from this
        # section's own "Unrealised value" framing). Filtered out here in
        # Python (case-EXACT sum, matching the formula above) before any
        # row is written, rather than hidden after the fact - a
        # zero-amount group simply never gets a row.
        gap_source_df = reco_df
        group_keys = ["receipt_status"]
        gap_has_status = status_col_s2 is not None and "final_delivery_status" in reco_df.columns
        if gap_has_status:
            fds = gap_source_df["final_delivery_status"]
            gap_source_df = gap_source_df.assign(
                **{"final_delivery_status": fds.fillna("Status Undefined").astype(str).str.strip()}
            )
            group_keys.append("final_delivery_status")
        gap_has_provider = provider_col_s2 is not None and "Payment Provider" in reco_df.columns
        if gap_has_provider:
            pv = gap_source_df["Payment Provider"]
            pv_str = pv.astype(str).str.strip()
            blank_mask = pv.isna() | pv_str.isin(["", "nan", "None", "#N/A", "NA"])
            gap_source_df = gap_source_df.assign(
                **{"Payment Provider": pv_str.where(~blank_mask, "Unresolved / #N/A")}
            )
            group_keys.append("Payment Provider")
        group_keys.append("query")
        gap_pairs_df = (
            gap_source_df[~gap_source_df["receipt_status"].isin(["Received", "Refunded"])]
            .groupby(group_keys)
            .agg(orders=("order_id", "size"), _amount=("diff", "sum"))
            .reset_index()
        )
        gap_pairs_df = gap_pairs_df[gap_pairs_df["_amount"].round(2) != 0]
        gap_pairs_df = gap_pairs_df.sort_values(group_keys).reset_index(drop=True)
    if gap_pairs_df is not None and not gap_pairs_df.empty:
        header_cols = ["Recipt Remark"]
        if gap_has_status:
            header_cols.append("Final Delivery Status")
        if gap_has_provider:
            header_cols.append("Payment Provider")
        header_cols += ["Query", "Unreconciled Amount", "Orders"]
        # gap_amount_col: the column letter "Unreconciled Amount" actually
        # lands in THIS run, given whichever of the two optional columns
        # above are present - computed here (not hardcoded) so Section 6's
        # Cashflow Waterfall below, which sums this exact column, always
        # points at the right one regardless of layout.
        gap_amount_col = get_column_letter(header_cols.index("Unreconciled Amount") + 1)
        table_header(header_cols + [""] * (7 - len(header_cols)))
        gap_start_row = row
        # Data-start row on 'Reco working' isn't always row 2: style_reco_
        # working_sections() inserts an extra group-header band row above
        # the column-name row whenever the sheet has column groups (this
        # config's normal case - "SHOPIFY REPORT"/"DELIVERY PARTNER
        # REPORT"/... - pushing data to row 3), but does nothing (data
        # stays at row 2) when groups is empty. Reading it straight off
        # the actual 'Reco working' worksheet - already fully written and
        # styled by this point in the pipeline - rather than assuming
        # either layout avoids a #VALUE! error from a range that
        # accidentally includes a text header cell (round 6 fix - this
        # exact off-by-one broke the very first version of the fix below).
        recon_ws = writer.sheets.get("Reco working")
        recon_header_rows = (recon_ws.max_row - len(reco_df)) if recon_ws is not None else 1
        recon_data_start = max(recon_header_rows, 1) + 1
        recon_last_row = max(recon_data_start, recon_data_start + len(reco_df) - 1)
        # Blank/unresolved Payment Provider cells on 'Reco working' can't
        # be matched with a single EXACT() literal the way a real label
        # can - a group relabelled "Unresolved / #N/A" in Python (see
        # blank_mask above) is matched live against every literal spelling
        # that maps to it, same convention as Section 5's own "#N/A" row.
        _BLANK_PROVIDER_LITERALS = ["", "nan", "None", "#N/A", "NA"]

        def _provider_mask(col, lit):
            if lit == "Unresolved / #N/A":
                terms = "+".join(
                    f"EXACT('Reco working'!{col}{recon_data_start}:{col}{recon_last_row},\"{b}\")"
                    for b in _BLANK_PROVIDER_LITERALS
                )
                return f"(({terms})>0)"
            return f"EXACT('Reco working'!{col}{recon_data_start}:{col}{recon_last_row},\"{lit}\")"

        for i, r in gap_pairs_df.iterrows():
            status_lit = str(r["receipt_status"]).replace('"', '""')
            query_lit = str(r["query"]).replace('"', '""')
            exact_mask = (
                f"EXACT('Reco working'!{recipt_col}{recon_data_start}:{recipt_col}{recon_last_row},\"{status_lit}\")*"
                f"EXACT('Reco working'!{query_col}{recon_data_start}:{query_col}{recon_last_row},\"{query_lit}\")"
            )
            if gap_has_status:
                fds_lit = str(r["final_delivery_status"]).replace('"', '""')
                exact_mask += (
                    f"*EXACT('Reco working'!{status_col_s2}{recon_data_start}:{status_col_s2}{recon_last_row},\"{fds_lit}\")"
                )
            if gap_has_provider:
                provider_lit = str(r["Payment Provider"]).replace('"', '""')
                exact_mask += f"*{_provider_mask(provider_col_s2, provider_lit)}"
            amt_formula = (
                f"=SUMPRODUCT(({exact_mask})*'Reco working'!{diff_col}{recon_data_start}:{diff_col}{recon_last_row})"
            )
            cnt_formula = f"=SUMPRODUCT(({exact_mask})*1)"
            band = "F7F7F5" if i % 2 else "FFFFFF"
            row_values = [r["receipt_status"]]
            row_numfmts = [None]
            if gap_has_status:
                row_values.append(r["final_delivery_status"])
                row_numfmts.append(None)
            if gap_has_provider:
                row_values.append(r["Payment Provider"])
                row_numfmts.append(None)
            row_values += [r["query"], amt_formula, cnt_formula]
            row_numfmts += [None, money_fmt, count_fmt]
            row_values += [None] * (7 - len(row_values))
            # 2026-09-06 (round 21, client-reported): every text/label
            # column here - Recipt Remark, Final Delivery Status, Payment
            # Provider (whichever are present), Query - reads left-aligned;
            # only the trailing "Unreconciled Amount"/"Orders" value
            # columns stay right. header_cols always ends with exactly
            # those two, so everything before them is a label column.
            data_row(row_values, row_numfmts, band=band, left_cols=set(range(1, len(header_cols) - 1)))
        gap_end_row = row - 1
        _box_border(ws, gap_start_row - 1, 1, gap_end_row, len(header_cols))
    else:
        gap_start_row = gap_end_row = None
        data_row(["No unreconciled-gap data available for this selection.", None, None, None, None, None, None])
    row += 1

    # ------------------------------------------------------------------
    # 3. Settlement Pending & Exceptions - by Payment Gateway
    # ------------------------------------------------------------------
    section_header("3. Settlement Pending & Exceptions (by Payment Gateway)")
    # Client-reported 2026-09-05: "'COD Delivered Amount not Received'
    # showing zero value thats correct if value is zero no need to show" -
    # a Payment Gateway row whose Amount Pending is genuinely 0 (every
    # order that ONCE sat in this gateway's pending bucket has since been
    # bank-matched/resolved) still passed straight through before this
    # fix, exactly like Section 2's own analogous zero-row noise the
    # client already asked to drop there. Filtered the same way: on the
    # underlying Python-computed "Amount Pending" figure (rounded, since
    # apply_settlement_pending_summary_formulas() below only ever REPLACES
    # the already-written value cells with a live SUMIFS/COUNTIFS formula
    # reproducing this exact same business rule - see its own docstring -
    # never changes which rows exist, so a row zero here will still
    # recalculate to zero live and is safe to drop up front).
    _pending_rows_df = (
        settlement_pending_summary_df[settlement_pending_summary_df["Amount Pending"].round(2) != 0]
        if (settlement_pending_summary_df is not None and not settlement_pending_summary_df.empty
            and "Amount Pending" in settlement_pending_summary_df.columns)
        else settlement_pending_summary_df
    )
    if _pending_rows_df is not None and not _pending_rows_df.empty:
        # Client-reported 2026-08-30 (item 8/12): dropped the "Exceptions"/
        # "Exception Amount" columns (E/F) to match the client's own
        # reference workbook's Executive Summary exactly - verified
        # against MOD.xlsx directly, only B/C/D (Payment Gateway/Orders
        # Pending/Amount Pending) are referenced there.
        table_header(["Payment Gateway", "Orders Pending", "Amount Pending", "", "", "", ""])
        for i, src_idx in enumerate(_pending_rows_df.index):
            # src_row must stay keyed off the ORIGINAL (unfiltered)
            # settlement_pending_summary_df position, since that is the
            # exact row order 'Settlement Pending Summary' was written in
            # (see _build_workbook()) - a dropped zero-value row still
            # occupies its own row there, so this cannot simply use i+2.
            src_row = src_idx + 2  # Settlement Pending Summary sheet: header row 1, data from row 2
            band = "F7F7F5" if i % 2 else "FFFFFF"
            data_row([
                f"='Settlement Pending Summary'!B{src_row}",
                f"='Settlement Pending Summary'!C{src_row}",
                f"='Settlement Pending Summary'!D{src_row}",
                None, None, None, None,
            ], [None, count_fmt, money_fmt], band=band)
        _box_border(ws, row - len(_pending_rows_df) - 1, 1, row - 1, 3)
    else:
        data_row(["No settlement pending / exceptions for this selection.", None, None, None, None, None, None])
    row += 1

    # ------------------------------------------------------------------
    # 4. Open Queries - Top Exposure
    # ------------------------------------------------------------------
    # Client-reported 2026-08-30 (item 12): re-verified directly against
    # MOD.xlsx's own live formulas - the "Query / Remark" text there turns
    # out to be a LITERAL value on every single row (not the LARGE/INDEX/
    # MATCH ranking formula an earlier pass here assumed and labelled
    # "client confirmed" - that assumption didn't survive an actual
    # formula-by-formula check), with only Orders/Exposure Amount as live
    # formulas referencing that literal text - COUNTIF('Reco working'!
    # Query column, text) and SUMIFS('Open queries'!G:G,'Open queries'!
    # A:A, text) respectively. Replicated exactly: Query text is written
    # as a plain value (already known in Python from open_queries_df,
    # itself already sorted by exposure descending), Orders/Exposure stay
    # live. Also: MOD lists EVERY distinct query (12 of 12 in the client's
    # own July run), not a fixed top-N - the previous top-6 cap here is
    # removed to match.
    section_header("4. Open Queries - Top Exposure")
    if open_queries_df is not None and not open_queries_df.empty and query_col:
        table_header(["Rank", "Query / Remark", "Orders", "Exposure Amount", "", "", ""])
        gap_top_start = row
        for k, (_, r) in enumerate(open_queries_df.iterrows(), start=1):
            band = "F7F7F5" if k % 2 else "FFFFFF"
            query_text = r["Query"] if "Query" in open_queries_df.columns else r["query"]
            query_lit = str(query_text).replace('"', '""')
            # 2026-09-06 (round 21, client-reported): "Query / Remark"
            # (column 2) is a text label sharing a row with three numeric
            # columns (Rank, Orders, Exposure Amount) - left-align just
            # this one rather than the numeric columns either side of it.
            data_row([
                k,
                query_text,
                f"=COUNTIF('Reco working'!{query_col}:{query_col},\"{query_lit}\")",
                f"=SUMIFS('Open queries'!G:G,'Open queries'!A:A,\"{query_lit}\")",
                None, None, None,
            ], [count_fmt, None, count_fmt, money_fmt], band=band, left_cols={2})
        _box_border(ws, gap_top_start - 1, 1, row - 1, 4)
    else:
        data_row(["No open queries for this selection.", None, None, None, None, None, None])
    row += 1

    # ------------------------------------------------------------------
    # 5. Payment Gateway / Collection Channel
    # ------------------------------------------------------------------
    section_header("5. Payment Gateway / Collection Channel")
    # Client-reported 2026-08-30 (item 12): fully rebuilt against MOD.xlsx's
    # OWN actual formulas (an earlier pass here had guessed at this
    # section's structure from stale memory rather than checking the real
    # file - genuinely different, and wrong, on every column). Real
    # columns/sources, verified formula-by-formula:
    #   Provider                              - Payment Provider label
    #   Order count                           - COUNTIF('Reco working'!R:R, provider)
    #   Order Value                           - SUMIFS('Reco working'!I:I, R:R, provider)   [total]
    #   Amount Received before PG Deduction   - SUMIFS('Reco working'!S:S, R:R, provider)   [receipt_amount, raw]
    #   PG Deduction                          - SUMIFS('Reco working'!T:T, R:R, provider)   [total_deduction]
    #   Bank credit                           - SUMIFS('Reco working'!W:W, R:R, provider)   [exported/zeroed Bank credit]
    #   setlemtn Pending Amount               - SUMIFS('Gateway Recon Recoperiod'!F:F, A:A, provider) [Pending Settlement]
    # (column letters resolved dynamically via _reco_col_letter, not
    # hardcoded - shown above using MOD's own letters for readability).
    # Provider rows come from Payment Provider's own distinct values in
    # reco_df (excluding blank/unresolved, handled as its own "#N/A" row
    # below) - NOT from "Gateway Settlement overall" (a UTR/settlement-side
    # attribution that can legitimately carry a different label set) and
    # NOT replicating MOD's own stray "easebuzz, Delhivery COD" row (a
    # leftover from the comma-joined-gateway bug the item-7 fix already
    # eliminates at the source - see engine/settlement.py's own comment on
    # this same stale artifact in "Gateway Recon Recoperiod").
    provider_col = _reco_col_letter(reco_working_cols, "Payment Provider")
    total_col = _reco_col_letter(reco_working_cols, "total")
    receipt_col = _reco_col_letter(reco_working_cols, "receipt_amount")
    deduction_col = _reco_col_letter(reco_working_cols, "total_deduction")
    bank_credit_col = _reco_col_letter(reco_working_cols, "settlement_amount")
    provider_labels = []
    has_unresolved = False
    if reco_df is not None and not reco_df.empty and "Payment Provider" in reco_df.columns:
        pv = reco_df["Payment Provider"]
        pv_str = pv.astype(str).str.strip()
        blank_mask = pv.isna() | pv_str.isin(["", "nan", "None", "#N/A", "NA"])
        provider_labels = sorted(pv_str[~blank_mask].unique())
        has_unresolved = bool(blank_mask.any())
    if provider_labels and provider_col and total_col and receipt_col and deduction_col and bank_credit_col:
        table_header([
            "Provider", "Order count", "Order Value", "Amount Received before PG Deduction",
            "PG Deduction", "Bank credit", "setlemtn Pending Amount",
        ])
        gw_start_row = row
        for i, label in enumerate(provider_labels):
            band = "F7F7F5" if i % 2 else "FFFFFF"
            label_lit = label.replace('"', '""')
            data_row([
                label,
                f"=COUNTIF('Reco working'!{provider_col}:{provider_col},\"{label_lit}\")",
                f"=SUMIFS('Reco working'!{total_col}:{total_col},'Reco working'!{provider_col}:{provider_col},\"{label_lit}\")",
                f"=SUMIFS('Reco working'!{receipt_col}:{receipt_col},'Reco working'!{provider_col}:{provider_col},\"{label_lit}\")",
                f"=SUMIFS('Reco working'!{deduction_col}:{deduction_col},'Reco working'!{provider_col}:{provider_col},\"{label_lit}\")",
                f"=SUMIFS('Reco working'!{bank_credit_col}:{bank_credit_col},'Reco working'!{provider_col}:{provider_col},\"{label_lit}\")",
                f"=SUMIFS('Gateway Recon Recoperiod'!F:F,'Gateway Recon Recoperiod'!A:A,\"{label_lit}\")",
            ], [None, count_fmt, money_fmt, money_fmt, money_fmt, money_fmt, money_fmt], band=band)
        if has_unresolved:
            # "#N/A" (unattributed Payment Provider) - see this function's
            # own docstring on reco_df: a live whole-column blank-match
            # formula would overcount every unused cell below the data, so
            # these 4 cells are Python-computed literals for this one row
            # only; Bank credit/setlemtn Pending stay live formulas against
            # Gateway Recon Recoperiod's own already-isolated "Unresolved /
            # #N/A" row (safe - matches on literal text, not blank cells).
            unresolved = reco_df[
                reco_df["Payment Provider"].isna()
                | reco_df["Payment Provider"].astype(str).str.strip().isin(["", "nan", "None", "#N/A", "NA"])
            ]
            band = "F7F7F5" if len(provider_labels) % 2 else "FFFFFF"
            # A bare "#N/A" string, entered as a plain cell value, is
            # auto-detected as the Excel/LibreOffice #N/A ERROR literal
            # (not text) on file open - confirmed via LibreOffice recalc
            # (a real formula-error cell, not a display quirk). Using the
            # same "Unresolved / #N/A" label Gateway Recon Recoperiod's
            # own row already carries instead - unambiguous text, no
            # collision, and it's the exact criteria the two SUMIFS below
            # already match against.
            data_row([
                "Unresolved / #N/A",
                int(len(unresolved)),
                round(float(unresolved["total"].sum()), 2) if "total" in unresolved.columns else 0.0,
                round(float(unresolved["receipt_amount"].sum()), 2) if "receipt_amount" in unresolved.columns else 0.0,
                round(float(unresolved["total_deduction"].sum()), 2) if "total_deduction" in unresolved.columns else 0.0,
                f"=SUMIFS('Gateway Recon Recoperiod'!I:I,'Gateway Recon Recoperiod'!A:A,\"Unresolved / #N/A\")",
                f"=SUMIFS('Gateway Recon Recoperiod'!F:F,'Gateway Recon Recoperiod'!A:A,\"Unresolved / #N/A\")",
            ], [None, count_fmt, money_fmt, money_fmt, money_fmt, money_fmt, money_fmt], band=band)
        _box_border(ws, gw_start_row - 1, 1, row - 1, 7)
    else:
        data_row(["No gateway settlement data available for this selection.", None, None, None, None, None, None])
    row += 1

    # ------------------------------------------------------------------
    # 6. Cashflow Waterfall - Order Value to Bank Credit
    # ------------------------------------------------------------------
    # Client-reported 2026-08-30 (item 12): rebuilt to match MOD.xlsx's own
    # 8-row waterfall exactly (verified formula-by-formula) - an
    # "Unreconciled Gap" step (summing Section 2's own Unreconciled Amount
    # column, so it's always internally consistent with that section) sits
    # between Order Value and Expected Collection, and Expected Collection/
    # Gateway Deduction/Refund/Settlement Amount are now sourced straight
    # from Reco working's own totals rather than the Dashboard's headline
    # figures - the same numbers by construction, but matching the
    # client's own formula text/source sheet exactly rather than only its
    # result.
    section_header("6. Cashflow Waterfall (Order Value to Bank Credit)")
    table_header(["Step", "Amount", "", "", "", "", ""])
    waterfall_start = row
    refund_col = _reco_col_letter(reco_working_cols, "refund_amount")
    # Client-reported 2026-09-05 (point 3): this used to hardcode column
    # "C" for Section 2's own "Unreconciled Amount" column - correct back
    # when that section was only ever 4 columns wide (Recipt Remark/Query/
    # Unreconciled Amount/Orders), but silently wrong the moment Section 2
    # grew a "Payment Provider" column (and now a "Final Delivery Status"
    # column too - see that section's own comment) shifted "Unreconciled
    # Amount" further right. Uses gap_amount_col - computed in Section 2
    # from the ACTUAL header layout written this run - instead, so this
    # formula can never drift out of sync with that section again no
    # matter how many more columns it grows in future.
    gap_sum_formula = (
        f"=SUM({gap_amount_col}{gap_start_row}:{gap_amount_col}{gap_end_row})"
        if gap_start_row and gap_end_row and gap_amount_col else "0"
    )
    waterfall_steps = [
        ("Order Value", dash_ref("Gross order value"), True),
        ("Unreconciled Gap (RTO, Cancelled, Lost, Setlment Pending)", gap_sum_formula, True),
        ("Expected Collection", f"=SUM('Reco working'!{receipt_col}:{receipt_col})" if receipt_col else "0", True),
        ("Gateway  Deduction", f"=SUM('Reco working'!{deduction_col}:{deduction_col})" if deduction_col else "0", True),
        ("Refund", f"=SUM('Reco working'!{refund_col}:{refund_col})" if refund_col else "0", True),
        ("Settlement Amount", f"=SUM('Reco working'!{bank_credit_col}:{bank_credit_col})" if bank_credit_col else "0", True),
        ("Bank Credit", bank_credit_formula, has_bank_reco),
    ]
    for i, (label, value, has_fmt) in enumerate(waterfall_steps):
        band = "F7F7F5" if i % 2 else "FFFFFF"
        data_row([label, value, None, None, None, None, None],
                  [None, money_fmt if has_fmt else None], band=band)
    settlement_row = waterfall_start + 5
    bank_credit_row = waterfall_start + 6
    gap_formula = (
        f"=B{settlement_row}-B{bank_credit_row}" if has_bank_reco else
        "Not available - no bank statement uploaded"
    )
    data_row(["Gap / Difference (timing, per engine/bank.py reconciliation rules)", gap_formula,
              None, None, None, None, None],
              [None, money_fmt if has_bank_reco else None], band="FFF8E1")
    _box_border(ws, waterfall_start - 1, 1, row - 1, 2)
    row += 1

    # ------------------------------------------------------------------
    # 7. Receivable Collection Period Analysis - For Book Closure &
    #    Receivable Reconciliation
    # ------------------------------------------------------------------
    # Section added 2026-08-27 (client-reported, point 7); REBUILT
    # 2026-08-31 after a further client-reported bug in its very first
    # column: "when I filter July, the tool is showing the entire June
    # receipt value... instead of identifying only the June orders whose
    # payment was actually received/credited in July." Root cause (see
    # engine.bank.bank_reconciliation_by_utr's own "THE BUG THIS FIXES"
    # docstring note): that function's "Settled (previous period
    # transaction settled this period)" column (col G) used to show a
    # previous-period order's FULL claimed settlement regardless of
    # whether its bank credit actually posted within the selected window,
    # after it, or even many months before it - fixed at the source now,
    # so this section only had to add the missing third bucket and a
    # genuine reconciliation check on top of that fix.
    #
    # Four categories, exactly as specified by the client, each sourced
    # from a column engine.bank.bank_reconciliation_by_utr() now computes
    # correctly for the SELECTED period specifically (its period_
    # start_date/period_end_date parameters - see views/page_reports.py's
    # _render_dtc for how those are derived from this report's own
    # From/To date filters):
    #   1. "Previous Period Amount Received This Period" - col G,
    #      "Settled (previous period transaction settled this period)":
    #      a PRIOR period's sale, its bank credit landed WITHIN the
    #      selected window.
    #   2. "This Period Amount Received in Same Period" - col D, "Amount
    #      (this period settled this period)": this period's own sale,
    #      received in the same period.
    #   3. "Order ID Not Found - Credited This Period" - col H, "Order ID
    #      Not Found - Settled During Reco Period" (NEW to this section
    #      2026-08-31 - previously computed by engine/bank.py but never
    #      surfaced here): a bank credit landed within the selected
    #      window, but the settlement ledger row it belongs to couldn't be
    #      traced to any known order at all.
    #   4. "This Period Amount Received in Subsequent Period" - col E,
    #      "Amount (this period transaction Settled Subsequent Period)":
    #      this period's own sale, not yet received by the time the bank
    #      statement closed - the roll-forward candidate that should
    #      reappear under THIS column's bucket 1 once the following
    #      period's own report is generated (client-reported, confirmed
    #      by a dedicated three-month June->July->August test - see
    #      /tmp/round7_point7/test_bank_window_fix.py).
    # A fifth, informational-only column ("Current Period Receivable
    # Pending", sourced from Settlement Pending Summary, unchanged from
    # before this rebuild) is kept alongside the four - it answers a
    # related but different question ("how much of THIS period's own
    # sales hasn't even been claimed by the gateway/courier yet", as
    # opposed to bucket 4's "claimed but not yet bank-credited"), and the
    # client's spec doesn't ask for it to be removed.
    #
    # Reconciliation check (the client's own explicit requirement): "the
    # total of the first three columns should match the actual Bank
    # Credit/Receipt amount for the selected period... This should match
    # the bank statement exactly, subject only to clearly identified
    # reconciliation exceptions." Added below the Total row as two new
    # rows - the Total row's own bucket 1+2+3 sum, and the selected
    # period's ACTUAL bank statement credit total (period_bank_credit_
    # total - a Python-computed figure, not a formula into this same
    # sheet's own Total row, which would just restate it circularly and
    # prove nothing; see this function's own docstring for why a plain
    # number is used here specifically).
    #
    # Orders/UTRs classify_order_periods() can't date at all ("Order not
    # found" - no prior saved period covers them) still fall outside the
    # first three buckets except via bucket 3 itself (col H) - the "Order
    # ID Not Found" concept always meant exactly this, not a data gap this
    # section invents.
    #
    # Disclosed gap, not guessed around: "Bank Reco (UTR-wise)"'s own
    # Payment Gateway column (engine.attribution.build_payment_gateway_
    # lookups()'s UTR-keyed half) and "Settlement Pending Summary"'s
    # Payment Gateway column (this Reco working row's own Gateway column -
    # see engine/settlement_pending.py) are resolved by two different
    # lookups built at different points in this pipeline's history: they
    # agree for every gateway/courier confirmed against the client's own
    # July data, but a channel resolvable to one sheet and not the other
    # (e.g. a COD courier whose bank UTR can't be matched at all) would
    # show its own figure on only one side of this table rather than
    # blending into a wrong row - SUMIF's own case-insensitive text match
    # handles a mere casing difference (e.g. "payu" vs "Payu") already.
    section_header("7. Receivable Collection Period Analysis (For Book Closure & Receivable Reconciliation)")
    table_header([
        "Payment Gateway / Channel", "Previous Period Amount Received This Period",
        "This Period Amount Received in Same Period", "Order ID Not Found - Credited This Period",
        "This Period Amount Received in Subsequent Period", "Current Period Receivable Pending", "",
    ])
    if not has_bank_reco:
        data_row(["Not available - no bank statement uploaded", None, None, None, None, None, None])
    else:
        # Client-reported 2026-08-31 (round 4, point 5): "Easebuzz and
        # PayU are appearing twice" in this section. Root cause: the two
        # source sheets feeding this label list disagree on casing for
        # the same gateway - "Bank Reco (UTR-wise)"'s own Payment Gateway
        # column carries the RAW lowercase processor string
        # (engine.attribution's own "payu"/"easebuzz", by design - see
        # that module's docstring), while "Settlement Pending Summary"'s
        # Payment Gateway column already applies engine.settlement_
        # pending._display_gateway_label()'s capitalisation. The dedup
        # below used to be a plain case-SENSITIVE set, so "easebuzz" and
        # "Easebuzz" (or "payu"/"Payu") survived as two distinct rows
        # instead of collapsing into one - each computed correctly on its
        # own (SUMIF's text match IS case-insensitive, as the comment
        # above already notes), but the same gateway's pending amount got
        # split across two lines instead of shown as one. Fixed by
        # canonicalising every label's casing (same all-lowercase-only
        # capitalisation rule as _display_gateway_label(), duplicated
        # here in miniature rather than importing across engine modules,
        # since formatting.py otherwise takes only pre-computed
        # dataframes and never reaches back into engine logic) BEFORE the
        # set-based dedup runs, so both sources agree on one canonical
        # spelling per gateway.
        def _canonical_gateway_label(label):
            label = str(label or "").strip()
            return label.capitalize() if label and label.islower() else label

        gateway_labels = []
        if utr_bank_reco_df is not None and not utr_bank_reco_df.empty and "Payment Gateway" in utr_bank_reco_df.columns:
            gateway_labels += list(utr_bank_reco_df["Payment Gateway"].dropna().astype(str))
        if (settlement_pending_summary_df is not None and not settlement_pending_summary_df.empty
                and "Payment Gateway" in settlement_pending_summary_df.columns):
            gateway_labels += list(settlement_pending_summary_df["Payment Gateway"].dropna().astype(str))
        gateway_labels = sorted({
            _canonical_gateway_label(lbl) for lbl in gateway_labels
            if lbl and lbl.strip() and lbl.strip().lower() != "unknown"
        })
        if gateway_labels:
            section7_start = row
            for i, label in enumerate(gateway_labels):
                lit = label.replace('"', '""')
                band = "F7F7F5" if i % 2 else "FFFFFF"
                data_row([
                    label,
                    f"=SUMIF('Bank Reco (UTR-wise)'!C:C,\"{lit}\",'Bank Reco (UTR-wise)'!G:G)",
                    f"=SUMIF('Bank Reco (UTR-wise)'!C:C,\"{lit}\",'Bank Reco (UTR-wise)'!D:D)",
                    f"=SUMIF('Bank Reco (UTR-wise)'!C:C,\"{lit}\",'Bank Reco (UTR-wise)'!H:H)",
                    f"=SUMIF('Bank Reco (UTR-wise)'!C:C,\"{lit}\",'Bank Reco (UTR-wise)'!E:E)",
                    f"=SUMIF('Settlement Pending Summary'!B:B,\"{lit}\",'Settlement Pending Summary'!D:D)",
                    None,
                ], [None, money_fmt, money_fmt, money_fmt, money_fmt, money_fmt], band=band)
            _box_border(ws, section7_start - 1, 1, row - 1, 6)
        # Total row sums the whole underlying column directly (not just
        # the gateway rows listed above) - stays correct even if a
        # channel appears on one source sheet but not the other.
        total_row_num = row
        data_row([
            "Total",
            "=SUM('Bank Reco (UTR-wise)'!G:G)",
            "=SUM('Bank Reco (UTR-wise)'!D:D)",
            "=SUM('Bank Reco (UTR-wise)'!H:H)",
            "=SUM('Bank Reco (UTR-wise)'!E:E)",
            "=SUM('Settlement Pending Summary'!D:D)",
            None,
        ], [None, money_fmt, money_fmt, money_fmt, money_fmt, money_fmt], band="E8F0FE")

        # Reconciliation check (client's own explicit requirement - see
        # this section's docstring above): bucket 1 + bucket 2 + bucket 3
        # (columns B, C, D of the Total row just written) should equal
        # the selected period's ACTUAL bank statement credit total.
        data_row([
            "Sum of Categories 1+2+3 (Previous Period + Same Period + Order ID Not Found)",
            f"=B{total_row_num}+C{total_row_num}+D{total_row_num}", None, None, None, None, None,
        ], [None, money_fmt], band="FFF8E1")
        if period_bank_credit_total is not None:
            bank_total_row_num = row
            data_row([
                "Actual Bank Credit for Selected Period (per uploaded Bank Statement)",
                period_bank_credit_total, None, None, None, None, None,
            ], [None, money_fmt], band="FFF8E1")
            check_row_num = bank_total_row_num - 1
            data_row([
                "Reconciliation Check (Categories 1+2+3 less Actual Bank Credit - should be zero, "
                "subject to disclosed reconciliation exceptions)",
                f"=B{check_row_num}-B{bank_total_row_num}", None, None, None, None, None,
            ], [None, money_fmt], band="FFE0B2")
        else:
            data_row([
                "Reconciliation Check: not available - no bank statement date range could be determined "
                "for the selected period",
                None, None, None, None, None, None,
            ], band="FFE0B2")
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    fcell = ws.cell(row=row, column=1,
                     value="This is an auto-generated report. Figures are live formulas linked to the other "
                           "sheets in this workbook - editing a source sheet or re-running the tool for a "
                           "different date range updates every figure here.")
    fcell.font = Font(name=FONT_NAME, size=9, italic=True, color=INK_MUTED)
    fcell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for letter, width in [("A", 34), ("B", 20), ("C", 18), ("D", 20), ("E", 18), ("F", 18), ("G", 18)]:
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A4"


def add_amazon_dashboard_sheet(writer, waterfall_df, order_reco_df=None, settlement_register_df=None,
                                subsequent_settlements_df=None, channel_name=None, date_from=None,
                                date_to=None, generated_at=None, cutoff_date=None):
    """
    Amazon/marketplace-channel equivalent of add_dashboard_sheet() above -
    same front-page MIS idea (title band, KPI stat cards, charts, a key
    metrics table), but sourced from the waterfall / order-wise detail /
    settlement register tables engine/amazon_reco.py, engine/amazon_bank.py
    produce, instead of the DTC reco_df/lookup_df shape. Placed first
    (sheet index 0) so it's the first thing anyone sees when they open the
    downloaded workbook.
    """
    wb = writer.book
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    wf = dict(zip(waterfall_df["Particular"], waterfall_df["Amount"])) if waterfall_df is not None and not waterfall_df.empty else {}

    def _wf(label, default=0.0):
        return wf.get(label, default)

    # Row each waterfall Particular ends up on in the "Waterfall" sheet
    # (header is row 1, data starts row 2 - same order build_waterfall()
    # produces it in) - lets every KPI card below link straight to that
    # cell with a real formula instead of duplicating the number, so
    # clicking a KPI on the Dashboard shows the end user exactly which
    # cell (and, via link_waterfall_formulas()'s "How This Is Calculated"
    # column on that sheet, exactly how) it was derived.
    wf_row = {}
    if waterfall_df is not None and not waterfall_df.empty:
        wf_row = {p: i + 2 for i, p in enumerate(waterfall_df["Particular"])}

    def _wf_formula(label, use_abs=False):
        """A ='Waterfall'!B<row> (optionally ABS()-wrapped) formula string
        for this label if it's on the Waterfall sheet, else None - callers
        fall back to the plain computed number when there's no sheet to
        link to (e.g. no waterfall_df at all)."""
        row = wf_row.get(label)
        if row is None:
            return None
        ref = f"'Waterfall'!B{row}"
        return f"=ABS({ref})" if use_abs else f"={ref}"

    total_orders = len(order_reco_df) if order_reco_df is not None else 0
    if order_reco_df is not None and not order_reco_df.empty and "has_settlement_row" in order_reco_df.columns:
        pending = int((~order_reco_df["has_settlement_row"]).sum())
    else:
        pending = 0
    settled = max(total_orders - pending, 0)

    net_sales = _wf("Net Sales")
    order_ded = abs(_wf("Less: Order-level Deductions (Flat File, incl. TDS/TCS)"))
    settlement_ded = abs(_wf("Less: Settlement-level Deductions (Flat File)"))
    receivable = _wf("Receivable")
    balance_receivable = _wf("Balance Receivable")
    received_to_date = abs(_wf("Less: Received to date (Bank Statement)"))

    # Order-wise Detail sheet is already written by the time this Dashboard
    # sheet gets built (see _build_amazon_workbook's write loop) - link
    # "Total MTR Orders"/"Settlement Pending (Orders)" to it live rather
    # than duplicating counts, finding the has_settlement_row column by
    # its header text so this never goes stale if that dataframe's column
    # order changes.
    total_orders_formula = None
    pending_formula = None
    if "Order-wise Detail" in wb.sheetnames:
        od_ws = wb["Order-wise Detail"]
        order_id_col = _find_col_letter(od_ws, "order_id") or "A"
        last_row = od_ws.max_row
        if last_row >= 2:
            total_orders_formula = f"=COUNTA('Order-wise Detail'!{order_id_col}2:{order_id_col}{last_row})"
            settlement_col = _find_col_letter(od_ws, "has_settlement_row")
            if settlement_col:
                pending_formula = (
                    f"=COUNTIF('Order-wise Detail'!{settlement_col}2:{settlement_col}{last_row}, FALSE)"
                )

    generated_at = generated_at or dt.datetime.now()
    date_lo, date_hi = _fmt_date(date_from), _fmt_date(date_to)

    # --- Title band ------------------------------------------------------
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:T1")
    title_cell = ws["A1"]
    title_cell.value = "AMAZON RECONCILIATION DASHBOARD"
    title_cell.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 1, 1, 1, 20, NAVY)

    # --- Sub-header: channel / date range / cut-off / generated-on --------
    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:T2")
    parts = []
    if channel_name:
        parts.append(f"Channel: {channel_name}")
    if date_lo and date_hi:
        parts.append(f"Date Range: {date_lo} to {date_hi}")
    if cutoff_date is not None:
        parts.append(f"Reporting Cut-off: {_fmt_date(cutoff_date)}")
    parts.append(f"Report Generated On: {generated_at.strftime('%d %b %Y, %I:%M %p')}")
    sub_cell = ws["A2"]
    sub_cell.value = "      |      ".join(parts)
    sub_cell.font = Font(name=FONT_NAME, size=10.5, bold=True, color=INK_SECONDARY)
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    _fill_block(ws, 2, 1, 2, 20, "F2F2F0")

    # --- KPI stat cards ----------------------------------------------------
    # Each card uses a live formula linking back to the Waterfall / Order-
    # wise Detail sheet where one exists, falling back to the plain
    # computed number only when there's nothing to link to (e.g. this
    # workbook has no waterfall_df at all).
    card_row_defs = [
        [
            ("Total MTR Orders", total_orders_formula or total_orders, CAT_BLUE, "#,##0"),
            ("Net Sales", _wf_formula("Net Sales") or net_sales, CAT_AQUA, '"₹"#,##0.00'),
            ("Receivable", _wf_formula("Receivable") or receivable, CAT_VIOLET, '"₹"#,##0.00'),
            ("Balance Receivable", _wf_formula("Balance Receivable") or balance_receivable,
             STATUS_CRITICAL, '"₹"#,##0.00'),
        ],
        [
            ("Order-level Deductions",
             _wf_formula("Less: Order-level Deductions (Flat File, incl. TDS/TCS)", use_abs=True) or order_ded,
             CAT_AMBER, '"₹"#,##0.00'),
            ("Settlement-level Deductions",
             _wf_formula("Less: Settlement-level Deductions (Flat File)", use_abs=True) or settlement_ded,
             CAT_ORANGE, '"₹"#,##0.00'),
            ("Received to Date",
             _wf_formula("Less: Received to date (Bank Statement)", use_abs=True) or received_to_date,
             STATUS_GOOD, '"₹"#,##0.00'),
            ("Settlement Pending (Orders)", pending_formula or pending, STATUS_CRITICAL, "#,##0"),
        ],
    ]
    card_top = 4
    card_h = 4
    card_w = 4
    gap = 1
    for row_i, cards in enumerate(card_row_defs):
        r1 = card_top + row_i * (card_h + 1)
        r2 = r1 + card_h - 1
        c = 1
        for label, value, accent, numfmt in cards:
            _kpi_card(ws, r1, c, r2, c + card_w - 1, label, value, accent, numfmt)
            c += card_w + gap

    charts_top = card_top + len(card_row_defs) * (card_h + 1) + 1
    charts_row = charts_top

    # --- Chart data area (feeds the charts below), parked off to the right -
    data_row = charts_top
    _section_title(ws, data_row, 22, "Chart data (feeds the charts to the left)", span=10)
    ws.cell(row=data_row, column=22).font = Font(name=FONT_NAME, size=9, italic=True, color=INK_MUTED)

    wf_hdr_row = data_row + 1
    ws.cell(row=wf_hdr_row, column=22, value="Stage")
    ws.cell(row=wf_hdr_row, column=23, value="Amount")
    wf_stage_rows = [
        ("Net Sales", net_sales),
        ("Order Deductions", -order_ded),
        ("Settlement Deductions", -settlement_ded),
        ("Receivable", receivable),
        ("Balance Receivable", balance_receivable),
    ]
    for i, (label, val) in enumerate(wf_stage_rows, start=1):
        ws.cell(row=wf_hdr_row + i, column=22, value=label)
        ws.cell(row=wf_hdr_row + i, column=23, value=val)

    status_cnt_hdr_row = wf_hdr_row
    ws.cell(row=status_cnt_hdr_row, column=25, value="Status")
    ws.cell(row=status_cnt_hdr_row, column=26, value="Orders")
    ws.cell(row=status_cnt_hdr_row + 1, column=25, value="Settled")
    ws.cell(row=status_cnt_hdr_row + 1, column=26, value=settled)
    ws.cell(row=status_cnt_hdr_row + 2, column=25, value="Pending")
    ws.cell(row=status_cnt_hdr_row + 2, column=26, value=pending)

    settlement_status_rows = []
    if (settlement_register_df is not None and not settlement_register_df.empty
            and "status" in settlement_register_df.columns):
        settlement_status_rows = list(settlement_register_df["status"].value_counts().items())

    settle_hdr_row = wf_hdr_row
    ws.cell(row=settle_hdr_row, column=28, value="Settlement Status")
    ws.cell(row=settle_hdr_row, column=29, value="Count")
    for i, (label, val) in enumerate(settlement_status_rows, start=1):
        ws.cell(row=settle_hdr_row + i, column=28, value=str(label))
        ws.cell(row=settle_hdr_row + i, column=29, value=int(val))

    for row in range(data_row, data_row + max(len(wf_stage_rows), len(settlement_status_rows), 3) + 2):
        for col in (22, 23, 25, 26, 28, 29):
            cell = ws.cell(row=row, column=col)
            if row > data_row and cell.value is not None:
                cell.font = Font(name=FONT_NAME, size=9, color=INK_MUTED)
                if col == 23:
                    cell.number_format = '"₹"#,##0.00'
                elif col in (26, 29):
                    cell.number_format = "#,##0"

    # --- Reconciliation Waterfall bar chart --------------------------------
    bar = BarChart()
    bar.type = "col"
    bar.title = "Reconciliation Waterfall"
    bar.y_axis.title = "Amount (Rs)"
    bar.gapWidth = 60
    bar.legend = None
    data = Reference(ws, min_col=23, min_row=wf_hdr_row, max_row=wf_hdr_row + len(wf_stage_rows))
    cats = Reference(ws, min_col=22, min_row=wf_hdr_row + 1, max_row=wf_hdr_row + len(wf_stage_rows))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    _colour_points(bar.series[0], [CAT_AQUA, CAT_AMBER, CAT_ORANGE, CAT_VIOLET, STATUS_CRITICAL])
    bar.width, bar.height = 15, 8.5
    ws.add_chart(bar, f"A{charts_row}")

    # --- Order Settlement Status donut --------------------------------------
    donut_orders = DoughnutChart()
    donut_orders.title = "Order Settlement Status"
    donut_orders.legend.position = "b"
    data = Reference(ws, min_col=26, min_row=status_cnt_hdr_row, max_row=status_cnt_hdr_row + 2)
    cats = Reference(ws, min_col=25, min_row=status_cnt_hdr_row + 1, max_row=status_cnt_hdr_row + 2)
    donut_orders.add_data(data, titles_from_data=True)
    donut_orders.set_categories(cats)
    _colour_points(donut_orders.series[0], [STATUS_GOOD, STATUS_CRITICAL])
    donut_orders.dataLabels = DataLabelList()
    donut_orders.dataLabels.showPercent = True
    donut_orders.dataLabels.showVal = False
    donut_orders.dataLabels.showCatName = False
    donut_orders.dataLabels.showSerName = False
    donut_orders.dataLabels.showLegendKey = False
    donut_orders.width, donut_orders.height = 11, 9.5
    ws.add_chart(donut_orders, f"J{charts_row}")

    # --- Settlement-to-Bank Status donut (only if bank matching ran) -------
    if settlement_status_rows:
        donut_settle = DoughnutChart()
        donut_settle.title = "Settlement-to-Bank Status"
        donut_settle.legend.position = "b"
        data = Reference(ws, min_col=29, min_row=settle_hdr_row, max_row=settle_hdr_row + len(settlement_status_rows))
        cats = Reference(ws, min_col=28, min_row=settle_hdr_row + 1, max_row=settle_hdr_row + len(settlement_status_rows))
        donut_settle.add_data(data, titles_from_data=True)
        donut_settle.set_categories(cats)
        cycle = [STATUS_GOOD, STATUS_CRITICAL, CAT_AMBER, CAT_BLUE]
        _colour_points(donut_settle.series[0], cycle[:len(settlement_status_rows)])
        donut_settle.dataLabels = DataLabelList()
        donut_settle.dataLabels.showPercent = True
        donut_settle.dataLabels.showVal = False
        donut_settle.dataLabels.showCatName = False
        donut_settle.dataLabels.showSerName = False
        donut_settle.dataLabels.showLegendKey = False
        donut_settle.width, donut_settle.height = 11, 9.5
        ws.add_chart(donut_settle, f"Q{charts_row}")

    key_metrics_row = charts_row + 18

    # --- Key Metrics table (the full waterfall, every line visible) --------
    _section_title(ws, key_metrics_row, 1, "Key Metrics (Full Waterfall)", span=6)
    table_hdr_row = key_metrics_row + 1
    for col, text in [(1, "Particular"), (5, "Amount")]:
        cell = ws.cell(row=table_hdr_row, column=col, value=text)
        ws.merge_cells(start_row=table_hdr_row, start_column=col,
                        end_row=table_hdr_row, end_column=col + (3 if col == 1 else 1))
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    wf_rows_list = (
        list(zip(waterfall_df["Particular"], waterfall_df["Amount"]))
        if waterfall_df is not None and not waterfall_df.empty else []
    )
    for i, (label, value) in enumerate(wf_rows_list, start=1):
        r = table_hdr_row + i
        band = "F7F7F5" if i % 2 else "FFFFFF"
        _fill_block(ws, r, 1, r, 6, band)
        lcell = ws.cell(row=r, column=1, value=label)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        lcell.font = Font(name=FONT_NAME, size=10, color=INK_SECONDARY)
        lcell.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        # Row i here lines up exactly with row i+1 on the "Waterfall" sheet
        # (same waterfall_df, same order) - link to it with a real formula
        # rather than repeating the number, so this table is traceable
        # straight back to the sheet that has the full "how calculated"
        # detail (see link_waterfall_formulas()), not just a duplicate.
        waterfall_row = wf_row.get(label, i + 1)
        vcell = ws.cell(row=r, column=5, value=f"='Waterfall'!B{waterfall_row}")
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        vcell.font = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
        vcell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        vcell.number_format = '"₹"#,##0.00'
    if wf_rows_list:
        _box_border(ws, table_hdr_row, 1, table_hdr_row + len(wf_rows_list), 6)

    # --- Footer -------------------------------------------------------------
    footer_row = table_hdr_row + len(wf_rows_list) + 3
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=12)
    fcell = ws.cell(row=footer_row, column=1,
                     value="This is an auto-generated report. For any queries, please contact the Finance Team.")
    fcell.font = Font(name=FONT_NAME, size=9, italic=True, color=INK_MUTED)
    fcell.alignment = Alignment(horizontal="left", vertical="center")

    # --- Column widths / page setup -----------------------------------------
    for letter in "ABCDEFGHIJKLMNOPQRST":
        ws.column_dimensions[letter].width = 13
    ws.print_area = f"A1:T{footer_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
