"""
shiprocket_cod.py
------------------
Auto-converts Shiprocket's own RAW COD Remittance export into the same
shape as the finance team's hand-built "mapped" reference file, so users
can upload the raw file Shiprocket actually gives them instead of being
forced to manually rebuild the mapped one first.

Background (2026-08-20 - client-reported "Invalid File" error on
Shiprocket COD upload): the tool's "Shiprocket COD" gateway config reads
its payable amount from a "COD Available - Line Allocation" column (plus
a "Freight Charges - Line Allocation" deduction column). Those two columns
- along with "Remarks" and "Final Payable Amount" - only exist in the
finance team's own reference/"mapped" file. Shiprocket's raw export
doesn't have them at all, because Shiprocket only reports COD Available
and Freight Charges at the CRF (remittance batch) level, not per
individual shipment (AWB) - the mapped file's extra columns were being
built by hand, allocating each CRF batch's totals down across its AWB
lines.

Comparing a real client-provided raw file against its already-mapped
counterpart (2,842 AWB rows across 42 CRF batches) confirmed the exact,
consistent allocation rule used (verified to ~1e-10 floating-point
tolerance on every single row):

    COD Available - Line Allocation   = the AWB's own "Order Value"
    Freight Charges - Line Allocation = (Order Value / CRF's "COD
                                         Available") * CRF's "Freight
                                         Charges from COD"
                                         -> each AWB's share of its
                                         batch's total freight, weighted
                                         by how much of the batch's COD
                                         Available that AWB's own Order
                                         Value represents.
    Final Payable Amount              = COD Available - Line Allocation
                                         minus Freight Charges - Line
                                         Allocation
    Remarks                           = the matching CRF batch's own
                                         "remarks" column (looked up by
                                         CRF ID)

This module reproduces that same computation automatically, so the raw
file becomes indistinguishable from the mapped one by the time it reaches
the rest of the engine (engine.consolidator, engine.bank, etc. keep
reading "COD Available - Line Allocation" / "Freight Charges - Line
Allocation" exactly as before - see configs/*.json's "Shiprocket COD"
gateway entry, unchanged).

If a user still uploads an already-mapped file (e.g. an old habit, or a
finance team member who prefers to keep building it by hand), this is a
no-op passthrough - see is_already_mapped() below.

Update (2026-08-21 - client-reported false "Invalid File" when Shiprocket
renamed its own "CRF level report" sheet to "shiprocket CRF level
report"): the two source sheets are no longer located by NAME at all.
Both are found by scanning every sheet in the workbook and checking which
one's own COLUMN HEADERS match what an AWB-level line-item sheet vs a
CRF-level batch-totals sheet actually needs (see AWB_SIGNATURE /
CRF_SIGNATURE and engine.loaders.find_sheet_by_columns) - so a renamed
sheet, a re-cased sheet, or the two sheets simply being in the opposite
order all keep working with no code change, and only a genuine absence of
the required data (not just an unexpected name) ever produces an error.
"""

import pandas as pd

from .loaders import resolve_col, resolve_col_or_raise, find_sheet_by_columns

# Historical sheet names Shiprocket happens to have used - kept only as
# human-readable labels in error messages / the "already mapped" shortcut
# below, never as the basis for finding a sheet. Sheet identification
# itself is 100% signature-driven (see AWB_SIGNATURE / CRF_SIGNATURE).
AWB_SHEET_LABEL = "AWB-level report"
CRF_SHEET_LABEL = "CRF-level report"

# The column headers that identify an AWB-level (one row per shipment)
# sheet, regardless of what the sheet itself is named or where it sits in
# the workbook. CRF ID links each AWB line back to its remittance batch;
# AWB/Order Id/Order Value are the shipment's own identifying fields - a
# CRF-level (batch-totals) sheet has none of the latter three.
AWB_SIGNATURE = ["CRF ID", "AWB", "Order Id", "Order Value"]
AWB_MIN_MATCH = 3  # tolerate one of the 4 being renamed without failing detection

# The column headers that identify a CRF-level (one row per remittance
# batch) sheet. "Freight Charges from COD" is the primary label seen in
# practice; "Freight Charges" is kept as a fallback alias in case a future
# export drops the "from COD" suffix.
CRF_SIGNATURE = ["CRF ID", "COD Available", ["Freight Charges from COD", "Freight Charges"]]
CRF_MIN_MATCH = 2  # CRF ID + COD Available alone is already a strong, distinctive signal

# The 4 columns that only exist in the MAPPED file - their presence is
# what distinguishes "already mapped" from "still raw" (see
# is_already_mapped()).
MAPPED_ONLY_COLS = [
    "COD Available - Line Allocation",
    "Freight Charges - Line Allocation",
    "Final Payable Amount",
    "Remarks",
]


def is_already_mapped(awb_df):
    """True once every column the mapping step would otherwise add is
    already present - lets a user who still uploads an already-mapped
    file keep working exactly as before, with nothing recomputed."""
    return all(c in awb_df.columns for c in MAPPED_ONLY_COLS)


def map_shiprocket_cod_raw_to_mapped(sheets, context=None):
    """
    sheets: dict of {sheet_name: DataFrame}, e.g. from
    pd.read_excel(file, sheet_name=None) - every sheet in the uploaded
    workbook, unfiltered.

    context: shared calling-convention parameter (see engine.raw_transforms
    and engine.razorpay_settlement, which actually uses it to receive the
    already-uploaded Shopify order report) - this transform needs nothing
    from it, so it's accepted and ignored. Keeping the same (sheets,
    context) signature on every RAW_TRANSFORMS entry means
    views/page_upload.py never needs a per-source special case for what
    a given transform happens to need.

    Returns the AWB-level DataFrame, augmented with the 4 mapped-only
    columns computed per this module's docstring - ready to be validated
    and used exactly like the finance team's own mapped file.

    Which sheet is which is determined ENTIRELY by column headers (see
    AWB_SIGNATURE / CRF_SIGNATURE) - never by sheet name or position, so
    Shiprocket renaming either sheet, changing capitalisation/spacing, or
    swapping their order in the workbook all keep working unchanged.

    Raises KeyError with a specific, human-readable message (never a
    generic "Invalid File") when something genuinely required is missing
    - which data/columns, not which sheet NAME - so the Upload Data page
    can show the user exactly what to fix, per the client's own explicit
    ask ("the error should clearly mention which field/column is
    missing", "errors should be generated only when the actual required
    data/columns are missing").
    """
    all_sheets = {name: _clean_columns(df) for name, df in sheets.items()}

    awb_name, awb_df = find_sheet_by_columns(all_sheets, AWB_SIGNATURE, min_required=AWB_MIN_MATCH)
    if awb_df is None:
        raise KeyError(
            f"Could not find a sheet that looks like a Shiprocket {AWB_SHEET_LABEL} in this file - "
            f"none of the {len(all_sheets)} sheet(s) found have the shipment-level columns expected "
            f"(needs at least {AWB_MIN_MATCH} of: {AWB_SIGNATURE}). "
            f"Sheets found: {list(all_sheets.keys())}."
        )
    awb_df = awb_df.copy()

    if is_already_mapped(awb_df):
        return awb_df

    crf_id_awb_col = resolve_col_or_raise(awb_df, "CRF ID", f'the "{awb_name}" sheet')
    order_value_col = resolve_col_or_raise(awb_df, "Order Value", f'the "{awb_name}" sheet')

    remaining_sheets = {name: df for name, df in all_sheets.items() if name != awb_name}
    crf_name, crf_df = find_sheet_by_columns(remaining_sheets, CRF_SIGNATURE, min_required=CRF_MIN_MATCH)
    if crf_df is None:
        raise KeyError(
            f'Found the shipment-level "{awb_name}" sheet, but none of the other sheet(s) in this '
            f"file have the CRF/remittance-batch-level columns needed to compute the COD and Freight "
            f"split per shipment (needs at least {CRF_MIN_MATCH} of: {CRF_SIGNATURE}) - Shiprocket only "
            "reports COD Available and Freight Charges at the remittance-batch level, so without that "
            f"data the per-shipment payable amount can't be computed. Sheets found: {list(all_sheets.keys())}."
        )
    crf_df = crf_df.copy()

    crf_id_crf_col = resolve_col_or_raise(crf_df, "CRF ID", f'the "{crf_name}" sheet')
    cod_available_col = resolve_col_or_raise(crf_df, "COD Available", f'the "{crf_name}" sheet')
    freight_col = resolve_col_or_raise(crf_df, ["Freight Charges from COD", "Freight Charges"], f'the "{crf_name}" sheet')
    # Optional - degrades to a blank Remarks column rather than failing
    # the whole upload over a purely informational field.
    remarks_col = resolve_col(crf_df, ["remarks", "Remarks"])
    # Optional (2026-08-31, client-reported: order 33160 showing as a
    # confirmed Bank Credit while the client had directly confirmed with
    # Shiprocket that its remittance was still genuinely pending) - see
    # module docstring below for the full story on why this needed adding.
    status_col = resolve_col(crf_df, ["Status", "Remittance Status"])
    crf_utr_col = resolve_col(crf_df, ["UTR", "UTR No"])

    crf_lookup = crf_df.set_index(crf_df[crf_id_crf_col].astype(str))
    crf_id_key = awb_df[crf_id_awb_col].astype(str)
    line_order_value = pd.to_numeric(awb_df[order_value_col], errors="coerce").fillna(0.0)

    matched_cod_available = crf_id_key.map(crf_lookup[cod_available_col])
    matched_freight = crf_id_key.map(crf_lookup[freight_col])

    # A CRF batch with zero total COD Available can't be proportionally
    # split (would divide by zero) - degrade that batch's freight
    # allocation to zero rather than crashing/NaN-ing the whole upload
    # over what would be an unusual, near-empty batch.
    safe_cod_available = matched_cod_available.replace(0, pd.NA)
    freight_share = (line_order_value / safe_cod_available * matched_freight).fillna(0.0)

    awb_df["COD Available - Line Allocation"] = line_order_value
    awb_df["Freight Charges - Line Allocation"] = freight_share
    awb_df["Final Payable Amount"] = awb_df["COD Available - Line Allocation"] - awb_df["Freight Charges - Line Allocation"]
    awb_df["Remarks"] = crf_id_key.map(crf_lookup[remarks_col]) if remarks_col else ""

    # 2026-08-31 fix (client-reported, order 33160): this module's own
    # docstring used to describe Shiprocket's raw export as only ever
    # listing amounts it has "actually remitted" - true of the AWB-level
    # sheet's OWN per-shipment fields, but NOT true of the file as a
    # whole. The CRF-level sheet carries a "Status" column with values
    # "Remittance success" / "Error" / "Remittance Initiated" - a real
    # client file had CRF 13360159 (order 33160's batch, 139 shipments)
    # sitting at "Remittance Initiated", not "success", meaning Shiprocket
    # itself had NOT confirmed this money as paid yet - exactly matching
    # the client's own statement that this remittance was still pending.
    # Every one of that CRF's AWB rows nonetheless had a real, non-null
    # "Order Value" and "Remittance Date" - fields this transform (and
    # everything downstream) previously had no reason to distrust.
    #
    # Propagated onto every AWB row here as "Remittance Status" so
    # configs/*.json's Shiprocket COD gateway entry can opt into the SAME
    # settled_status_col/settled_values filter engine/consolidator.py's
    # normalize_gateway_df() already applies for Prozo COD - rows whose
    # CRF hasn't actually succeeded are dropped before any amount/
    # deduction/date processing, exactly as if Shiprocket had never listed
    # them at all, rather than being counted as money already collected.
    # Degrades to no column at all (unchanged prior behaviour) if the CRF
    # sheet's Status column can't be found - never blocks the upload over
    # a field this fix is layering on top of, not requiring.
    if status_col:
        awb_df["Remittance Status"] = crf_id_key.map(crf_lookup[status_col])

    # Bonus precision, same fix: the AWB-level sheet's own "UTR" column
    # turns out to NOT be reliably populated even for CRFs that DID
    # succeed (confirmed against the same real file - 9 of 43 successful
    # CRFs had every one of their AWB rows blank on UTR, while the
    # CRF-level sheet's own UTR was populated) - Shiprocket evidently only
    # guarantees recording the bank reference once, at the batch level.
    # Backfilling it here means those orders get a real, precise
    # order-level UTR match (engine.bank.match_consolidated_to_bank)
    # instead of falling back to the less-precise settlement-batch
    # amount/date heuristic (engine.bank.match_batches_to_bank) purely
    # because of where Shiprocket happened to record the reference - never
    # overwrites a genuine per-shipment UTR the AWB sheet already has.
    if crf_utr_col:
        awb_utr_col = resolve_col(awb_df, ["UTR", "UTR No"])
        crf_utr_by_id = crf_id_key.map(crf_lookup[crf_utr_col])
        if awb_utr_col:
            awb_df[awb_utr_col] = awb_df[awb_utr_col].where(awb_df[awb_utr_col].notna(), crf_utr_by_id)
        else:
            awb_df["UTR"] = crf_utr_by_id

    return awb_df


def _clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


# The combined registry of every raw_transform this app knows about (this
# source plus e.g. Razorpay's) lives in engine.raw_transforms, not here -
# see that module for why it's kept separate from any one source's own file.
