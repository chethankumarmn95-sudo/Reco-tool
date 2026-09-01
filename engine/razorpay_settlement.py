"""
razorpay_settlement.py
-----------------------
Auto-maps Razorpay's own RAW settlement export against the Shopify order
report already uploaded in the same session, so the user no longer has to
hand-build a "mapped" reference file every time (2026-08-21 client request).

Background: Razorpay's raw settlement export has all the columns the
Razorpay gateway config needs to compute amounts/fees/UTR/date directly
(amount, "fee (exclusive tax)", tax, transaction_entity, settled_at,
settlement_utr - all already correctly configured in configs/*.json's
"Razorpay" gateway entry) - EXCEPT the order-linking column. Razorpay's
own "shopify_order_id" (and identical "order_receipt") column is NOT the
real Shopify order number - it's an opaque internal payment/checkout
token (e.g. "rQoM98S8eds7LuVaZ9EXs6QcO"), so joining on it directly
against the Shopify order report's own order-id column (Name/Order
Number) never matches anything.

Comparing a real client-provided raw settlement export against its
hand-built "mapped" counterpart (2,515 rows) and the Shopify order export
it was mapped against showed the real linking logic: that same opaque
token also appears, verbatim, in the Shopify order report's own "Payment
Reference" (equivalently "Payment ID"/"Payment References") column - so
an exact-match lookup of Razorpay's shopify_order_id/order_receipt
against the Shopify order report's Payment Reference column recovers the
real Shopify order number (e.g. "#4593") for ~98.5% of rows out of the
box; the client's own hand-built mapped file matches this exact rule
(same lookup column, same exact-match semantics, not a fuzzy/exploded
match) to within a small handful of rows plausibly explained by the two
reference files being different point-in-time snapshots (e.g. a Shopify
order added after the Razorpay export was pulled), not a logic
difference.

This module reproduces that lookup automatically: it adds a "Shopify
Order No" column (the client's own naming, from their mapped file) to
the raw Razorpay export, which configs/esca_shopify.json's "Razorpay"
gateway entry lists FIRST in its order_id_col aliases - so it's what the
rest of the engine (engine.reco, engine.consolidator, etc.) actually
joins orders against, exactly as if the file had been hand-mapped.
Wherever a Razorpay row's token doesn't resolve to any Shopify order
(genuinely no match, or the token itself is missing), the new column is
simply left blank - the same "no match found" outcome every other
gateway/delivery-partner source already produces for an unmatched row,
handled by the existing reconciliation join with no extra sentinel value
needed.

Sheet/column identification is 100% signature-driven (see
RAZORPAY_SIGNATURE and engine.loaders.find_sheet_by_columns /
resolve_col) - never by file name, sheet name, or sheet position, per
the same app-wide framework used for Shiprocket COD (see
engine.shiprocket_cod).
"""

import pandas as pd

from .loaders import resolve_col, resolve_col_or_raise, find_sheet_by_columns
from .transform_errors import TransformPrerequisiteError

# The column headers that identify a Razorpay raw settlement export,
# regardless of what the sheet/file is named. All 6 are Razorpay's own,
# fairly distinctive field names - a couple of these being renamed or
# absent in a future export still leaves this recognisable (see
# RAZORPAY_MIN_MATCH).
RAZORPAY_SIGNATURE = [
    "transaction_entity", "entity_id", "amount", "settlement_id", "settled_at",
    ["shopify_order_id", "order_receipt"],
]
RAZORPAY_MIN_MATCH = 4

# The column in the raw Razorpay file holding the opaque
# checkout/payment token that also shows up in the Shopify order
# report - this IS the join key, it just isn't a usable order number on
# its own.
TOKEN_COL_ALIASES = ["shopify_order_id", "order_receipt", "Shopify Order Id", "Order Receipt"]

# Where that same token re-appears in the Shopify order report, keyed to
# the real order number. Checked in priority order - "Payment Reference"
# is what the client's own verified mapping logic actually used (a
# single, always-singular value); "Payment ID"/"Payment References" are
# kept as fallbacks for exports where the primary field is named or
# populated differently, but are NOT exploded/split on inspection of a
# combined "tokenA + tokenB" value, since doing so was shown to produce
# matches the client's own hand-built reference file deliberately did not
# make (see this module's docstring).
SHOPIFY_LOOKUP_COL_ALIASES = ["Payment Reference", "Payment ID", "Payment References"]

SHOPIFY_ORDER_NUMBER_ALIASES = ["Name", "Order ID", "Shopify Order ID", "Order Number", "Shopify Order Number"]

# The column this module adds - named after the client's own mapped
# file's column, so a user who still hand-maps occasionally (old habit)
# is recognised as "already mapped" and passed through untouched (see
# is_already_mapped).
MAPPED_ORDER_COL = "Shopify Order No"
ALREADY_MAPPED_ALIASES = ["Shopify Order No", "Shopify order No"]


def is_already_mapped(df):
    return resolve_col(df, ALREADY_MAPPED_ALIASES) is not None


def map_razorpay_settlement_to_shopify(sheets, context=None):
    """
    sheets: dict of {sheet_name: DataFrame} - every sheet in the uploaded
    Razorpay workbook (a plain CSV upload arrives here as a single-entry
    {"Sheet1": df} dict - see views/page_upload.py's _read_all_sheets).

    context: {"orders_df": <the Shopify order report DataFrame already
    uploaded and confirmed earlier on this same Upload Data page>} - see
    views/page_upload.py's _load_gateway_file, which passes
    st.session_state["orders_df"] through here. Optional/ignored keys may
    be added to context in future without breaking this function - every
    raw_transform in engine.raw_transforms shares this same (sheets,
    context) calling convention so page_upload.py never needs a
    per-transform special case.

    Returns the Razorpay DataFrame augmented with MAPPED_ORDER_COL.

    Raises a specific, human-readable error (never a generic "Invalid
    File") when something genuinely required is missing - which
    sheet/column/prerequisite-upload, not a vague failure - per the
    client's explicit ask for named, actionable errors. Two different
    exception types, deliberately (2026-08-21, client-reported - see
    engine.transform_errors.TransformPrerequisiteError's own docstring for
    the full story of why this distinction exists):
      - KeyError: this file's OWN shape doesn't look like a Razorpay
        export at all, or is missing a column within it (no matching
        sheet found; no token column). These are genuinely candidates for
        "maybe this is actually some OTHER report" - views/page_upload.py
        follows up a KeyError here with that check.
      - TransformPrerequisiteError: the file IS a Razorpay export, but
        something ELSE this mapping depends on isn't ready yet - the
        Shopify order report hasn't been uploaded, or is missing the
        Payment Reference column this lookup needs. Has nothing to do
        with whether THIS file is the right report, so page_upload.py
        shows this message directly rather than running that check.
    """
    context = context or {}
    orders_df = context.get("orders_df")

    all_sheets = {name: _clean_columns(df) for name, df in sheets.items()}

    rp_name, rp_df = find_sheet_by_columns(all_sheets, RAZORPAY_SIGNATURE, min_required=RAZORPAY_MIN_MATCH)
    if rp_df is None:
        raise KeyError(
            f"Could not find a sheet that looks like a Razorpay settlement report in this file - "
            f"none of the {len(all_sheets)} sheet(s) found have the columns expected (needs at least "
            f"{RAZORPAY_MIN_MATCH} of: {RAZORPAY_SIGNATURE}). Sheets found: {list(all_sheets.keys())}."
        )
    rp_df = rp_df.copy()

    if is_already_mapped(rp_df):
        return rp_df

    token_col = resolve_col_or_raise(rp_df, TOKEN_COL_ALIASES, f'the "{rp_name}" sheet')

    if orders_df is None or orders_df.empty:
        raise TransformPrerequisiteError(
            "the Shopify order report needs to be uploaded (and confirmed above) first - Razorpay's "
            "settlement export only carries an internal payment/checkout token, not the real Shopify "
            "order number, so it can't be matched to an order until the order report's own Payment "
            "Reference column is available to look it up against."
        )

    orders_clean = _clean_columns(orders_df)
    lookup_col = resolve_col(orders_clean, SHOPIFY_LOOKUP_COL_ALIASES)
    if lookup_col is None:
        raise TransformPrerequisiteError(
            "could not find a Payment Reference / Payment ID column in the uploaded Shopify order "
            "report - this is needed to translate Razorpay's internal payment token into the real "
            f"Shopify order number. Actual columns in the order report: {list(orders_clean.columns)}."
        )
    order_number_col = resolve_col_or_raise(orders_clean, SHOPIFY_ORDER_NUMBER_ALIASES, "the Shopify order report")

    # Built only from rows that actually HAVE a lookup value, and only
    # applied to rows that actually HAVE a token - otherwise a blank
    # token (NaN) on one side and a blank lookup value (NaN) on the other
    # would both stringify to the literal text "nan" and collide as a
    # false match, silently attaching one row's real order number to a
    # completely unrelated row that never had any token at all.
    valid_orders = orders_clean[orders_clean[lookup_col].notna()]
    token_to_order_number = dict(zip(valid_orders[lookup_col].astype(str), valid_orders[order_number_col]))

    has_token = rp_df[token_col].notna()
    rp_df[MAPPED_ORDER_COL] = pd.NA
    rp_df.loc[has_token, MAPPED_ORDER_COL] = rp_df.loc[has_token, token_col].astype(str).map(token_to_order_number)

    return rp_df


def remap_unmapped_rows(df, orders_df):
    """
    Client-reported (2026-08-23): a Razorpay settlement row whose Shopify
    order genuinely exists and correctly maps end-to-end (verified against
    the client's own cross-checked reference: same token, same order
    number, same amount, same settlement UTR) was still showing up as an
    entirely unmatched row in the live tool - receipt_amount = 0 for that
    order in Reco working, the order wrongly still shown "Settlement
    Pending", and that portion of the settlement UTR's bank credit missing
    from Bank Reco (UTR-wise)'s "this period" total. Confirmed directly
    against the client's own live raw Razorpay data: the row was present,
    correctly amount/UTR/date, with MAPPED_ORDER_COL (the "Shopify Order
    No" column map_razorpay_settlement_to_shopify adds above) genuinely
    blank for it.

    Root cause: map_razorpay_settlement_to_shopify() above only ever runs
    ONCE, at the exact moment a Razorpay file is uploaded (see
    views/page_upload.py's _load_gateway_file), using whatever orders_df
    happened to be in session at that instant. A Razorpay settlement
    report very often gets uploaded before every order it eventually
    settles has itself been uploaded yet (payout lag routinely exceeds the
    order-upload cadence) - that row's token is then permanently stuck
    unresolved, because nothing ever asks the question again once the
    matching order finally does show up.

    Call this from run_dtc_reconciliation (not just upload time) so the
    SAME lookup gets one more real chance on every "Run reconciliation"
    click, against whatever orders_df is CURRENTLY known - not only the
    snapshot that existed the moment this file was first uploaded. Only
    rows still blank in MAPPED_ORDER_COL are looked up again; anything
    already mapped (correctly or not) is left exactly as it is - this
    never second-guesses a match that already succeeded, so a genuinely
    unmapped row (no matching order exists anywhere, ever) simply keeps
    getting retried harmlessly on every run rather than being touched
    incorrectly.

    df: this gateway's current accumulated raw DataFrame, in its own raw
    column shape - i.e. whatever's sitting in st.session_state[
    "gateway_frames"][label] right now. Must already carry MAPPED_ORDER_COL
    (already been through map_razorpay_settlement_to_shopify at least
    once) - a df that was never mapped at all is returned untouched, since
    this function only fills in gaps left by that first pass, it doesn't
    perform the original mapping itself.

    orders_df: the CURRENT Shopify order report (st.session_state[
    "orders_df"]) - cumulative across every upload confirmed so far.

    Returns df with MAPPED_ORDER_COL filled in wherever a previously-blank
    row's token now resolves; every other row (including one that still
    can't resolve - a genuinely orphaned transaction with no matching
    order at all) is returned exactly as it was.
    """
    if df is None or df.empty or MAPPED_ORDER_COL not in df.columns:
        return df
    if orders_df is None or orders_df.empty:
        return df

    mapped_text = df[MAPPED_ORDER_COL].astype(str).str.strip()
    still_blank = df[MAPPED_ORDER_COL].isna() | mapped_text.str.lower().isin(["nan", "none", ""])
    if not still_blank.any():
        return df

    token_col = resolve_col(df, TOKEN_COL_ALIASES)
    if token_col is None:
        return df

    orders_clean = _clean_columns(orders_df)
    lookup_col = resolve_col(orders_clean, SHOPIFY_LOOKUP_COL_ALIASES)
    order_number_col = resolve_col(orders_clean, SHOPIFY_ORDER_NUMBER_ALIASES)
    if lookup_col is None or order_number_col is None:
        return df

    # Same "don't let two blanks collide" guard as the original mapping
    # pass above.
    valid_orders = orders_clean[orders_clean[lookup_col].notna()]
    token_to_order_number = dict(zip(valid_orders[lookup_col].astype(str), valid_orders[order_number_col]))

    df = df.copy()
    retry_mask = still_blank & df[token_col].notna()
    if retry_mask.any():
        df.loc[retry_mask, MAPPED_ORDER_COL] = df.loc[retry_mask, token_col].astype(str).map(token_to_order_number)
    return df


def _clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df
