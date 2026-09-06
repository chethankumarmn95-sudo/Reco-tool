"""
consolidator.py
----------------
This is Layer 2 of your reconciliation pipeline - the equivalent of your
"Consolidated receipt" sheet.

Each payment/COD gateway (Gokwik, Razorpay, Delhivery COD, Shiprocket COD...)
reports money differently: different column names, different fee structures.
This file's job is to normalize all of them into ONE common table with the
same shape, so that later steps don't need to know or care which gateway a
receipt came from.

Output shape (one row per gateway transaction):
    order_id | source | txn_type | amount | deduction | is_refund
"""

import re

import pandas as pd
from .loaders import normalize_order_id, resolve_col_or_raise, resolve_col

_YEAR_FIRST_RE = re.compile(r"^\d{4}[-/]")


def _column_looks_year_first(series):
    """
    2026-09-06 (round 20, client-reported - Refund Date wrong/blank on the
    Reco working sheet): True when this date column's own values are
    already unambiguous ISO-style text starting with a 4-digit year (e.g.
    "2026-07-05") - dayfirst is NOT safe to force True for these. Verified
    directly: pandas' dayfirst=True can silently SWAP an already-
    unambiguous year-first date's month/day too (e.g. "2026-07-05" comes
    back as "2026-05-07" instead of being left alone) - dayfirst applies to
    whichever two components are left ambiguous, wherever the year sits,
    it isn't a no-op just because a 4-digit year is present. False for a
    day-first "DD-MM-YYYY"-style column (e.g. "01-07-2026", this client's
    own gateway/COD-partner settlement files - see _parse_date_col()'s own
    docstring, below) or for a column already holding real datetime
    objects (a native Excel date cell) - dayfirst is irrelevant to either
    of those (a datetime object has no ambiguity left to resolve, and a
    day-first string needs dayfirst=True to read correctly).

    Decided ONCE per column, off a sample of its own non-null values - a
    real export's own date column is consistently one format, never a mix
    of day-first and year-first rows.
    """
    sample = series.dropna()
    if sample.empty:
        return False
    sample = sample.astype(str).head(20)
    return bool(sample.str.match(_YEAR_FIRST_RE).any())


def _parse_dates_safely(series):
    """
    pd.to_datetime() with the correct dayfirst setting for THIS column,
    decided via _column_looks_year_first() above - see that function's own
    docstring for why a single blanket dayfirst=True (or dayfirst=False)
    is not safe across every gateway's date column. Shared by every date
    column this module parses (settlement dates and refund dates alike),
    so both use the exact same, once-decided rule.
    """
    dayfirst = not _column_looks_year_first(series)
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def normalize_gateway_df(df, gateway_cfg):
    """
    Convert one gateway's raw DataFrame into the common receipt shape,
    using the column mapping supplied in the client's config file.
    """
    # Every field here is resolved by header name via resolve_col (case/
    # whitespace/punctuation-insensitive, same as the upload-time
    # validation check) rather than assumed to match the config's literal
    # string exactly - see engine/reco.py's build_order_master docstring
    # for why that mismatch used to let a file pass validation and then
    # still crash right here. Only order_id and amount are load-bearing
    # enough to require (a gateway receipt with no amount is meaningless);
    # everything else degrades gracefully when the column can't be found.
    label = gateway_cfg["label"]

    # Some COD partner exports (Prozo's own COD report, 2026-08-25 client
    # request) include EVERY shipment's COD status, not just the ones
    # actually remitted to the bank yet - e.g. a "COD Status" column with
    # values like "REMITTED" vs "PENDING"/"IN TRANSIT". Unlike Delhivery's
    # and Shiprocket's existing COD exports (which only ever list amounts
    # they've actually remitted), treating every row here as "settled"
    # would count money that hasn't actually reached the bank yet as if
    # it had. When a gateway config names a settled_status_col + the
    # settled_values it should be filtered to, rows outside that set are
    # dropped BEFORE any of the amount/deduction/date processing below -
    # exactly as if they were never in the file at all, so a not-yet-
    # remitted row can't leak into Settlement Done for this gateway. Left
    # out of the config entirely (the default for every other gateway),
    # every row is treated as settled - unchanged prior behaviour.
    settled_status_col_spec = gateway_cfg.get("settled_status_col")
    if settled_status_col_spec:
        resolved_status_col = resolve_col(df, settled_status_col_spec)
        settled_values = {str(v).strip().lower() for v in gateway_cfg.get("settled_values", [])}
        if resolved_status_col and settled_values:
            status_text = df[resolved_status_col].astype(str).str.strip().str.lower()
            df = df[status_text.isin(settled_values)]

    order_col = resolve_col_or_raise(df, gateway_cfg["order_id_col"], label)
    amount_col = resolve_col_or_raise(df, gateway_cfg["amount_col"], label)
    deduction_col_specs = gateway_cfg.get("deduction_cols", [])
    deduction_cols = [c for c in (resolve_col(df, spec) for spec in deduction_col_specs) if c]
    type_col = resolve_col(df, gateway_cfg.get("type_col")) if gateway_cfg.get("type_col") else None
    refund_values = set(gateway_cfg.get("refund_values", []))
    date_col = resolve_col(df, gateway_cfg.get("date_col")) if gateway_cfg.get("date_col") else None

    out = pd.DataFrame()
    out["order_id"] = normalize_order_id(df[order_col])
    out["source"] = label
    out["amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    if deduction_cols:
        out["deduction"] = sum(
            pd.to_numeric(df[c], errors="coerce").fillna(0) for c in deduction_cols
        )
    else:
        out["deduction"] = 0.0

    if type_col:
        out["is_refund"] = df[type_col].astype(str).isin(refund_values)
    else:
        out["is_refund"] = False

    # Client-reported 2026-08-31 (round 4, point 4): some Refund UTRs come
    # through with no usable date at all. Investigated directly against
    # the client's own July data - every repeated Refund UTR is
    # consistently either always-dated or always-blank (see
    # engine/bank.py::build_refund_utr_detail()'s own docstring for the
    # full finding), which points to the underlying gateway file itself
    # having no completion date recorded yet for those specific refund
    # transactions, using the SAME "date_col" as payment rows. A refund
    # aggregator export commonly logs a refund's own completion date under
    # a DIFFERENT column than the one used for payment settlement dates -
    # this optional "refund_date_col" config key lets a gateway declare
    # that separately without guessing at any specific client's actual
    # column name here. Left unset (the default, and the current setting
    # for every gateway in this client's config), behaviour is completely
    # unchanged - every row, payment or refund, still reads "date_col".
    # Only refund-type rows (is_refund True) look at "refund_date_col";
    # payment rows always use "date_col" regardless.
    refund_date_col_spec = gateway_cfg.get("refund_date_col")
    resolved_refund_date_col = (
        resolve_col(df, refund_date_col_spec) if refund_date_col_spec else None
    )

    def _parse_date_col(col_name):
        # 2026-09-06, round 20, client-reported - Refund Date wrong/blank
        # on the Reco working sheet, order #26216/#26225: this client's
        # COD-partner settlement files use DD-MM-YYYY text dates (every
        # date example given across this engagement has been day-first).
        # Plain pd.to_datetime() (dayfirst=False, the old behaviour here)
        # silently SWAPS day and month whenever both are <=12 (e.g.
        # "01-07-2026", meant as 1 July, read back as 7 January - order
        # #26216's exact wrong Refund Date), and returns an outright
        # unparseable NaT (blank) whenever the day exceeds 12 (e.g.
        # "25-07-2026" - a valid day, but 25 isn't a valid MONTH - order
        # #26225's exact blank Refund Date). Both client examples are the
        # same single root cause, just surfacing differently depending on
        # whether that date's day component happens to exceed 12.
        #
        # _parse_dates_safely() (above) fixes this WITHOUT blindly forcing
        # dayfirst=True on every gateway, unlike the equivalent bank-
        # statement fix in engine/bank.py::to_naive_datetime_series() -
        # verified directly that an unconditional dayfirst=True is not
        # safe here: it can just as easily swap an ALREADY-unambiguous
        # ISO/year-first date (e.g. a Razorpay/Gokwik settlement date that
        # happens to come through as "2026-07-05") the other way, turning
        # it into "2026-05-07". Decided per-column instead, off that
        # column's own values - see _column_looks_year_first()'s docstring.
        parsed = _parse_dates_safely(df[col_name])
        # Normalized to tz-naive right here, at the one place gateway
        # settlement dates get parsed - some gateway exports (e.g.
        # Razorpay's settled_at) include a timezone offset, and once this
        # column is tz-aware every later date-arithmetic consumer
        # downstream (engine/bank.py's classify_order_bank_status ageing
        # logic, the settlement-batch date-window match, Excel export of
        # this date) would either crash comparing it against a naive date,
        # or fail to write to Excel at all (openpyxl doesn't support
        # timezone-aware datetimes). Only the calendar date/time matters
        # here, not the source's UTC offset.
        try:
            if parsed.dt.tz is not None:
                parsed = parsed.dt.tz_localize(None)
        except (AttributeError, TypeError):
            pass
        return parsed

    if date_col:
        out["receipt_date"] = _parse_date_col(date_col)
    else:
        out["receipt_date"] = pd.NaT

    if resolved_refund_date_col is not None:
        refund_date = _parse_date_col(resolved_refund_date_col)
        out.loc[out["is_refund"], "receipt_date"] = refund_date[out["is_refund"]]

    out["payment_mode"] = gateway_cfg.get("payment_mode", "Unknown")

    utr_col_spec = gateway_cfg.get("utr_col")
    resolved_utr_col = resolve_col(df, utr_col_spec) if utr_col_spec else None
    if resolved_utr_col is not None:
        raw_utr = df[resolved_utr_col]
        # Some exports (Prozo's own COD report, 2026-08-25) format this as
        # a comma-separated LIST field even when there's only ever one
        # value in it - e.g. "AXISP00816694369, " - leaving a trailing
        # comma (and sometimes trailing whitespace after it) that a plain
        # .str.strip() alone doesn't remove, so the UTR silently never
        # matched the bank statement's clean value. Splitting on the first
        # comma and keeping just that token handles both this single-
        # value-with-stray-comma case and a genuinely multi-value cell
        # (takes the first UTR rather than a compound string nothing
        # would ever match) - then stripped again for any inner whitespace.
        utr_text = raw_utr.astype(str).str.split(",").str[0].str.strip()
        # A genuinely blank/NaN cell stringifies to the literal 3-letter
        # text "nan" via astype(str) above - left as-is, that fake value
        # used to get treated as a REAL utr by every downstream consumer
        # (receipt_detail_by_order's join_unique, bank.py's matching), so
        # a perfectly good single UTR on one gateway row plus a genuinely
        # blank utr on another row for the SAME order silently became
        # "AXISCN1299268951, nan" once joined - a string that visually
        # still looks like "one real UTR" but never matches anything in
        # the bank statement (client-reported: a single, correct-looking
        # UTR still showing "Bank statement not found"). Kept as real
        # None here instead, so it's filtered out everywhere blanks
        # already are, rather than smuggled through as text.
        blank_mask = raw_utr.isna() | utr_text.str.lower().isin(["nan", "none", "nat", ""])
        out["utr"] = utr_text.where(~blank_mask, None)
    else:
        out["utr"] = None

    # A row with a genuinely blank order id AND no real money behind it is
    # just formatting noise some exports include (a fully blank trailing
    # row, a subtotal row, etc.) - safe to drop outright, since it would
    # contribute nothing to any total either way.
    #
    # A row with a blank order id but a REAL non-zero amount is a
    # different thing entirely: this is the gateway (Razorpay, in the
    # 2026-08-22 client report that surfaced this) telling us it actually
    # settled/refunded real money for a transaction this tool's own
    # order-linking step (engine.razorpay_settlement's token lookup, or
    # any other raw_transform) couldn't trace back to one specific
    # Shopify order - e.g. the Shopify order report snapshot on hand at
    # upload time simply didn't include that order yet. That money is
    # real and the gateway itself confirms it settled. Previously this
    # function dropped it from EVERY downstream total, not just the
    # per-order ones - so a gateway's own "Settlement Done" figure (see
    # engine.settlement.gateway_settlement_summary, which sums straight
    # off this DataFrame) could fall dramatically short of what that
    # gateway's own raw settlement export adds up to (client-reported: a
    # Razorpay file that actually netted ~Rs 60,215 in settlement came out
    # as only ~Rs 25,866 once run through the tool, because roughly half
    # its rows had no resolvable Shopify order and were silently dropped
    # here - the missing ~Rs 34,349 was sitting entirely in rows just like
    # this one, not in any calculation error on the rows that DID match).
    #
    # Kept here with order_id left blank, so gateway-level totals now
    # reflect it correctly. It simply can't join to any real order later
    # (attach_receipts_and_diff's order_master.merge(..., how="left")),
    # since no real order has a blank order_id - which is exactly right,
    # there IS no specific order to credit it to. The per-order Reco
    # working table and the Dashboard's order-driven "Net settlement"
    # headline are therefore correctly unaffected by this change:
    # summarize_receipts_by_order/receipt_detail_by_order (below) filter
    # this same blank order_id back OUT before their own per-order
    # grouping, so for those two views this function's old
    # drop-everything-blank behaviour is preserved exactly. Only the
    # gateway-level aggregate (which never should have been order-scoped
    # in the first place - a gateway's settlement total isn't defined by
    # which of its rows happen to match a Shopify order) sees the fix.
    blank_order_id = _blank_order_id_mask(out["order_id"])
    noise_row = blank_order_id & (out["amount"].abs() < 0.005) & (out["deduction"].abs() < 0.005)
    out = out[~noise_row]

    return out


def _blank_order_id_mask(order_id_series):
    """
    True wherever order_id_series has no usable value - covers every shape
    a "blank" can actually take here, not just one of them:
      - real NaN (pandas may keep this as an actual NaN through
        normalize_order_id()'s own .astype(str) rather than the literal
        text "nan", depending on the Series' dtype - a plain `== "nan"`
        or `.str.len() == 0` check alone silently misses this case and
        was verified, while building this fix, to let real-NaN blanks
        slip past undetected)
      - the literal text "nan" (classic object-dtype astype(str) result)
      - an empty string (already-blank text that stayed blank)
    Centralised here so normalize_gateway_df's noise-row check and the two
    per-order-only views below (summarize_receipts_by_order,
    receipt_detail_by_order) all agree on exactly what counts as "no
    order id" - if they disagreed, a row could count as attributable in
    one and not the other, silently double-counting or dropping money.
    """
    as_text = order_id_series.astype(str)
    return order_id_series.isna() | (as_text.str.len() == 0) | (as_text.str.lower() == "nan")


def build_consolidated_receipt(gateway_frames, gateway_configs):
    """
    gateway_frames: dict like {"Gokwik": df_gokwik, "Razorpay": df_razorpay, ...}
                     Only include the gateways the user actually uploaded -
                     it's fine if some are missing.
    gateway_configs: the "gateways" list from the client config json.

    Returns one combined DataFrame - your "Consolidated receipt" equivalent.
    """
    normalized = []
    for cfg in gateway_configs:
        label = cfg["label"]
        if label not in gateway_frames:
            continue  # user hasn't uploaded this source yet - skip, don't fail
        normalized.append(normalize_gateway_df(gateway_frames[label], cfg))

    if not normalized:
        return pd.DataFrame(columns=["order_id", "source", "amount", "deduction", "is_refund"])

    return pd.concat(normalized, ignore_index=True)


def build_pending_cod_receipts(gateway_frames, gateway_configs):
    """
    Client-reported 2026-09-04 (round 10, point 2): "Shiprocket COD Amount
    Not Reflecting in receipt_amount" - a COD order the Shiprocket COD
    Report plainly lists an amount for was still showing receipt_amount
    == 0 on the Reco working sheet whenever that specific row's own
    "Remittance Status" isn't yet "Remittance success" - i.e. exactly the
    rows normalize_gateway_df()'s settled_status_col filter (see its own
    docstring, 2026-08-25) drops BEFORE they ever reach consolidated_df /
    summarize_receipts_by_order() - as if they were never in the file at
    all. That filter exists for a good reason (money not yet remitted to
    the bank must not be counted as a genuine, bank-matchable settlement
    row - see engine.bank.classify_order_bank_status's has_settlement_row/
    grace-period logic, which depends on it staying exactly as-is) and is
    NOT changed here.

    This is instead an ADDITIVE, separate signal: the exact same not-yet-
    settled rows normalize_gateway_df() drops, kept here purely so
    engine.reco.attach_pending_cod_receipts() can recognise "the courier's
    own report already confirms this money was collected from the
    customer, only the remittance-to-bank hasn't posted yet" and reflect
    THAT in receipt_amount/diff (Reco working's own numbers), completely
    independent of - and without altering - has_settlement_row/
    Reconciliation Category, which keep being computed from the original,
    unchanged, settled-only consolidated_df exactly as before.

    Only gateways that declare "settled_status_col" in gateway_configs are
    considered (today: Shiprocket COD, Prozo COD) - a COD gateway with no
    such column (Delhivery COD) has nothing to add here, since its own
    export format never lists a not-yet-remitted row in the first place
    (every row in a Delhivery COD file already represents money actually
    remitted).

    Returns order_id | source | amount | deduction - one row per order per
    COD gateway with at least one not-yet-settled row (summed, in the rare
    case a single order spans more than one such row). Empty (never None)
    when no configured COD gateway has both a settled_status_col AND an
    uploaded file this run.
    """
    cols = ["order_id", "source", "amount", "deduction"]
    if not gateway_frames or not gateway_configs:
        return pd.DataFrame(columns=cols)

    frames = []
    for cfg in gateway_configs:
        settled_status_col_spec = cfg.get("settled_status_col")
        if not settled_status_col_spec:
            continue
        label = cfg["label"]
        df = gateway_frames.get(label)
        if df is None or df.empty:
            continue
        resolved_status_col = resolve_col(df, settled_status_col_spec)
        if not resolved_status_col:
            continue
        settled_values = {str(v).strip().lower() for v in cfg.get("settled_values", [])}
        status_text = df[resolved_status_col].astype(str).str.strip().str.lower()
        pending_df = df[~status_text.isin(settled_values)]
        if pending_df.empty:
            continue

        order_col = resolve_col(pending_df, cfg["order_id_col"])
        amount_col = resolve_col(pending_df, cfg["amount_col"])
        if not order_col or not amount_col:
            continue
        deduction_col_specs = cfg.get("deduction_cols", [])
        deduction_cols = [c for c in (resolve_col(pending_df, spec) for spec in deduction_col_specs) if c]

        out = pd.DataFrame()
        out["order_id"] = normalize_order_id(pending_df[order_col])
        out["source"] = label
        out["amount"] = pd.to_numeric(pending_df[amount_col], errors="coerce").fillna(0)
        if deduction_cols:
            out["deduction"] = sum(
                pd.to_numeric(pending_df[c], errors="coerce").fillna(0) for c in deduction_cols
            )
        else:
            out["deduction"] = 0.0

        out = out[~_blank_order_id_mask(out["order_id"])]
        out = out[out["amount"].abs() > 0.004]
        if len(out):
            frames.append(out[cols])

    if not frames:
        return pd.DataFrame(columns=cols)
    combined = pd.concat(frames, ignore_index=True)
    grouped = combined.groupby(["order_id", "source"], as_index=False)[["amount", "deduction"]].sum()
    return grouped


def summarize_receipts_by_order(consolidated_df):
    """
    Collapse the consolidated receipt ledger down to one row per order:
        order_id | receipt_amount | total_deduction | refund_amount

    This is the equivalent of the SUMIFS formulas in your 'Reco working'
    sheet (columns V, W, Y) that pull receipt/deduction/refund per order.
    """
    if consolidated_df.empty:
        return pd.DataFrame(columns=["order_id", "receipt_amount", "total_deduction", "refund_amount"])

    # Per-order figures can only ever be attributed to a real order id -
    # normalize_gateway_df() now deliberately KEEPS gateway rows with a
    # blank order id (see its own docstring) so gateway-level totals stay
    # accurate, but those same rows have nothing to join to here and must
    # be excluded before this per-order grouping, or they'd collapse into
    # one fake "order" keyed by the empty string. This restores the exact
    # per-order behaviour this function always had before that change.
    attributable = consolidated_df[~_blank_order_id_mask(consolidated_df["order_id"])]

    payments = attributable[~attributable["is_refund"]]
    refunds = attributable[attributable["is_refund"]]

    receipt = payments.groupby("order_id")["amount"].sum().rename("receipt_amount")
    deduction = payments.groupby("order_id")["deduction"].sum().rename("total_deduction")
    refund = refunds.groupby("order_id")["amount"].sum().rename("refund_amount")

    result = pd.concat([receipt, deduction, refund], axis=1).fillna(0).reset_index()
    return result


def receipt_detail_by_order(consolidated_df):
    """
    For the Order Lookup dashboard: which gateway(s) paid this order, what
    mode (COD/Prepaid), and the earliest receipt date. Kept separate from
    summarize_receipts_by_order() because the financial reco only needs
    totals, but lookup needs to show "how" and "when" too.

    "utr" here is ONE row per order - the single most useful UTR to show
    in a quick-glance, one-row-per-order column (Settlement Pending
    Report, Order Lookup), never a comma-joined "UTR1, UTR2, UTR3" blob
    (client-reported: "Each row should contain only one UTR"). An order
    settled across more than one genuinely distinct reference (e.g. a COD
    order remitted in two separate courier batches) shows its MOST
    RECENT one here; the full picture - every one of that order's
    distinct UTRs checked independently against the bank statement - is
    what engine.bank.build_settlement_ledger/bank_reconciliation_by_utr
    actually reconcile against, and they work from the raw consolidated_df
    directly (one row per gateway transaction) rather than from this
    already-collapsed view, so they were never affected by the old
    joined-UTR string either way.
    """
    cols = ["order_id", "payment_gateway", "payment_mode", "receipt_date", "utr"]
    if consolidated_df.empty:
        return pd.DataFrame(columns=cols)

    # Same reasoning as summarize_receipts_by_order() above: rows without a
    # resolvable order id are real gateway money (kept for gateway-level
    # totals) but have no order to be looked up under here.
    attributable = consolidated_df[~_blank_order_id_mask(consolidated_df["order_id"])]

    payments = attributable[~attributable["is_refund"]].copy()

    def join_unique(series):
        vals = sorted({
            str(v).strip() for v in series
            if pd.notna(v) and str(v).strip() and str(v).strip().lower() not in ("nan", "none", "nat")
        })
        return ", ".join(vals)

    grouped = payments.groupby("order_id").agg(
        payment_gateway=("source", join_unique),
        payment_mode=("payment_mode", join_unique),
        receipt_date=("receipt_date", "min"),
    ).reset_index()

    has_utr = payments["utr"].notna() & (payments["utr"].astype(str).str.strip() != "")
    latest_utr = (
        payments[has_utr]
        .sort_values("receipt_date", ascending=False, na_position="last")
        .drop_duplicates(subset="order_id", keep="first")[["order_id", "utr"]]
    )
    grouped = grouped.merge(latest_utr, on="order_id", how="left")
    grouped["utr"] = grouped["utr"].fillna("")

    return grouped[cols]
