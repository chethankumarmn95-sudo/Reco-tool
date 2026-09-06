"""
settlement_pending.py
----------------------
Settlement Pending Report: every order whose money hasn't fully completed
the chain "Order -> Payment/COD collected -> Gateway settlement -> Bank
credit" yet, as one order-wise detail table plus a gateway-wise / COD
partner-wise summary - the two views asked for alongside the existing
Payment Gateway Settlement Report (engine/settlement.py) and the
reconciliation categories (engine/bank.py).

Built directly on top of engine.bank.classify_order_bank_status() output,
so this report, the Order Lookup "Reconciliation Category" column, and the
Bank Statement / UTR Linking page never disagree about which bucket an
order is in - one classification, three views.

2026-09-06 (round 19, client-reported): "Expected Settlement Date" has been
REMOVED from this report (item 2) - the client never used that estimate
and asked for it to be dropped, so the note above about it being a
best-effort planning figure (not any gateway's actual SLA) no longer
applies to anything in this module.
"""

import pandas as pd
from openpyxl.utils import get_column_letter

from .bank import (
    COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW,
    resolve_split_payment_leg_status,
)
from .settlement import expected_gateway_for_order
from .reco import _COD_SETTLEMENT_PENDING_PHRASES
from .attribution import cod_component_of_gateway_label, prepaid_component_of_gateway_label

PENDING_CATEGORIES = {COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW}

# 2026-09-06 (round 19, client-reported item 2): "Expected Settlement Date"
# removed - client never used that estimate and asked for it to be dropped.
# 2026-09-06 (round 19, client-reported item 5): "Report Period End Date"
# added right after "Order Date" - it is the same value on every row (the
# reconciliation period's own end date, e.g. the Reports page's "To date"
# selector), written out so the "Days Pending" column's own live formula
# (see build_settlement_pending_report()'s Days Pending construction, below)
# has a visible, on-sheet cell to point at - "for better visibility and
# user understanding", per the client's own words, rather than an
# invisible fixed value baked into a formula referencing some cell outside
# the printed columns.
DETAIL_COLUMNS = [
    "Order ID", "Payment Gateway", "Payment/Transaction Reference (UTR)",
    "Order Date", "Report Period End Date", "Payment/Receipt Date", "Order Amount", "Gateway Amount",
    "Settlement Amount", "Actual Settlement Date",
    "Bank Credit Date", "Settlement Status", "Days Pending",
]

# Client-reported 2026-08-30 (item 8): matches the client's own reference
# workbook's "Settlement Pending Summary" sheet exactly - 4 columns, no
# separate "Exceptions"/"Exception Amount" breakout (Orders Pending/Amount
# Pending already includes every still-outstanding order, exceptions
# included - see settlement_pending_summary_by_gateway()'s own docstring).
SUMMARY_COLUMNS = ["Group", "Payment Gateway", "Orders Pending", "Amount Pending"]

# 2026-09-06 (round 22, client-reported, order #29284): the Settlement
# Pending Summary sheet's new distinct category for a COD order correctly
# attributed to its delivering courier but genuinely absent from that
# courier's own COD/settlement report - see settlement_pending_summary_
# by_gateway()'s own note for the full story. Hoisted to module scope (not
# a local inside that function) so engine/formatting.py::apply_settlement_
# pending_summary_formulas() - which rewrites this sheet's Orders/Amount
# Pending cells into live SUMIFS/COUNTIFS formulas after export - can
# recognise this exact label without a second, hand-copied literal that
# could drift out of sync with this one.
NOT_REFLECTING_LABEL = "COD amount not reflecting in COD report and pending to bank credit"

# The substring unique to engine.reco.py::apply_cod_report_gap_query()'s
# own query phrasing ("{courier} COD Delivered but amount not reflecting
# in {courier} COD Report") - see _is_cod_report_absent_query()'s own
# docstring for why this must be narrower than the "not reflecting"
# substring used elsewhere. Hoisted here (not just a literal inside that
# function) so engine/formatting.py::apply_settlement_pending_summary_
# formulas() can build a matching SUMIFS/COUNTIFS wildcard criteria
# ("*" + this + "*") off the exact same text, never a second copy.
COD_REPORT_ABSENT_QUERY_FRAGMENT = "delivered but amount not reflecting in"

# The BROAD substring shared by BOTH "not reflecting" query phrasings
# (apply_cod_report_gap_query()'s own text above, AND refine_queries_
# with_settlement_status()'s more generic "{courier} Setlment pending not
# reflecting") - see _is_report_gap_query()'s own docstring. Hoisted here
# (round 22) so engine/formatting.py::apply_settlement_pending_summary_
# formulas() can reproduce, for a COD courier's OWN row, the exact same
# receipt_amount-vs-Total split pending_amount_by_order() already applies
# in Python - a courier's pending order whose query merely says "Setlment
# pending not reflecting" (nothing has settled anywhere yet) has
# receipt_amount == 0 too, same as the newly-split-out gap-specific case,
# so that courier's own live-formula total must also use Total for it, not
# receipt_amount, or the live sheet would silently undercount against the
# already-correct Python-computed figure.
BROAD_NOT_REFLECTING_FRAGMENT = "not reflecting"


def build_settlement_pending_report(reco_df, recon_status_df, receipt_detail_df, gateway_configs,
                                     receipt_summary_df=None, period_end_date=None):
    """
    One row per order currently sitting in a "not yet fully reconciled"
    state - COD or Prepaid, settlement pending OR flagged as an Exception.
    Delivered/settled/bank-matched orders and COD orders that were never
    delivered are excluded on purpose - they don't need chasing, see
    engine/bank.py's classify_order_bank_status() for why.

    receipt_summary_df (engine.consolidator.summarize_receipts_by_order()
    output - order_id | receipt_amount | total_deduction | refund_amount)
    is optional but recommended: when supplied, "Settlement Amount" is the
    net figure (Gateway Amount minus that gateway's own deduction, minus
    any refund) - the same "final payment" concept used everywhere else in
    this engine (engine/reco.py's settlement_amount, engine/bank.py's
    build_settlement_ledger). Without it, Settlement Amount falls back to
    equalling Gateway Amount (gross) - still correct for orders with no
    settlement row yet (nothing has been deducted from nothing), just less
    precise for orders that do have one.

    period_end_date (2026-09-06, round 19, client-reported item 5) - the
    reconciliation/report period's own end date (e.g. the Reports page's
    "To date" selector). Used for two things: the new "Report Period End
    Date" column (same value every row, added purely so the "Days Pending"
    formula below has something on-sheet to point at), and as the anchor
    for that formula itself. Optional and defaults to today when omitted,
    for backward compatibility with any older caller that doesn't have a
    period end date to pass - but every real call site in this codebase
    now passes one. See the "Days Pending" construction below for why this
    replaces classify_order_bank_status()'s own "days_pending" field
    entirely for this sheet, rather than trying to thread period_end_date
    through that function's own (today-based) pending_clock logic.
    """
    if reco_df is None or reco_df.empty or recon_status_df is None or recon_status_df.empty:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    df = recon_status_df[recon_status_df["Reconciliation Category"].isin(PENDING_CATEGORIES)].copy()
    if df.empty:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    reco_cols = ["order_id", "created_at", "total", "payment_method"]
    if "delivery_partner" in reco_df.columns:
        reco_cols.append("delivery_partner")
    # 2026-09-06 (round 18, client-reported): the Payment Gateway column
    # below used to be sourced from receipt_detail_df's raw consolidated-
    # ledger "source" (e.g. "Gokwik" - the uploaded gateway FILE's own
    # label, before downstream-processor refinement) or, failing that, a
    # fresh best-effort guess (expected_gateway_for_order() - Payment
    # Method text + delivery_partner only) that knows nothing about every
    # other signal engine.attribution.attach_payment_columns() already
    # resolved (Gokwik Transaction Report downstream processor, COD-
    # courier override, ...) - producing "Gokwik" or "Unattributed
    # Prepaid"/"Unattributed COD" for orders whose Reco working sheet
    # already shows a specific, correct Payment Provider ("PayU",
    # "Easebuzz", "Delhivery COD", "Shiprocket COD", ...). Pull that
    # already-resolved column in directly so this sheet can never disagree
    # with Reco working's own - see _attributed_gateway() below for the
    # exact priority order.
    for col in ("Payment Provider", "Gateway"):
        if col in reco_df.columns and col not in reco_cols:
            reco_cols.append(col)
    # 2026-09-06 (round 19, client-reported item 1): "Gateway Amount" below
    # used to be recon_status_df's OWN receipt_amount - a figure engine.bank.
    # classify_order_bank_status() computes strictly from consolidated_df
    # (payments.groupby("order_id")["amount"].sum()), independent of every
    # correction reco_df's own receipt_amount already carries (in
    # particular, engine.reco.attach_pending_cod_receipts()'s round-10
    # Shiprocket-COD rescue - see that function's own docstring). That
    # divergence is exactly the client's own reported example: the Detail
    # sheet's Shiprocket COD Gateway Amount total didn't match the
    # Settlement Pending Summary sheet's (correct) figure, which is built
    # from reco_df's own receipt_amount via pending_amount_by_order() below.
    # Pull reco_df's own receipt_amount in here too (renamed to avoid a
    # merge-suffix collision with recon_status_df's own column of the same
    # name) so this sheet's Gateway Amount can never disagree with the
    # Summary sheet/Reco working sheet again - see the Gateway Amount
    # column construction, below, for the exact fallback order.
    has_reco_receipt_amount = "receipt_amount" in reco_df.columns
    if has_reco_receipt_amount:
        reco_cols.append("receipt_amount")
    reco = reco_df[reco_cols].copy()
    reco["order_id"] = reco["order_id"].astype(str)
    if has_reco_receipt_amount:
        reco = reco.rename(columns={"receipt_amount": "_reco_receipt_amount"})
    df = df.merge(reco, on="order_id", how="left")

    if receipt_detail_df is not None and not receipt_detail_df.empty:
        rd = receipt_detail_df[["order_id", "payment_gateway", "utr"]].copy()
        rd["order_id"] = rd["order_id"].astype(str)
        df = df.merge(rd, on="order_id", how="left")
    else:
        df["payment_gateway"] = None
        df["utr"] = None

    if receipt_summary_df is not None and not receipt_summary_df.empty:
        rs = receipt_summary_df[["order_id", "total_deduction", "refund_amount"]].copy()
        rs["order_id"] = rs["order_id"].astype(str)
        df = df.merge(rs, on="order_id", how="left")
    else:
        df["total_deduction"] = 0.0
        df["refund_amount"] = 0.0
    df["total_deduction"] = df["total_deduction"].fillna(0.0)
    df["refund_amount"] = df["refund_amount"].fillna(0.0)
    # 2026-09-06 (round 19, client-reported item 3): "Settlement Amount"
    # used to show the order's own Total for any order with no settlement
    # row yet (has_settlement_row == False) - i.e. it showed a number even
    # for orders where nothing has actually settled anywhere. Client's own
    # words: "Settlement Amount should be zero since none of the order
    # setled to bank - Settlement Amount should include only bank
    # credited." Rebuilt so this column is 0 unless the order has actually
    # been matched to a bank credit (bank_matched - see engine.bank.
    # classify_order_bank_status()'s own docstring/columns) - only then is
    # there a genuine net-of-deductions/refund figure to show at all. This
    # narrows the column's meaning; it does not change what counts as
    # "still outstanding" for chasing purposes - that's Gateway Amount
    # (item 1, immediately below) and the Settlement Pending Summary sheet
    # (engine.settlement_pending.pending_amount_by_order()), both
    # untouched by this change.
    _net_settlement = (df["receipt_amount"] - df["total_deduction"] - df["refund_amount"]).round(2)
    # .get()-style fallback (same safe-zero-fill convention used elsewhere
    # in this engine, e.g. engine/reco.py::headline_totals()'s own
    # settlement_pending_amount.get()) for a recon_status_df that predates
    # this column or omits it - degrades to "nothing confirmed bank-
    # matched," never a crash.
    _bank_matched = df["bank_matched"] if "bank_matched" in df.columns else pd.Series(False, index=df.index)
    df["_settlement_amount"] = _net_settlement.where(_bank_matched, 0.0)

    # 2026-09-06 (round 18, client-reported): priority order, highest to
    # lowest specificity -
    #   1. reco_df's own resolved "Payment Provider" (falling back to
    #      "Gateway" if that specific column isn't present) - the SAME
    #      value already shown on the Reco working sheet, carrying every
    #      refinement engine.attribution.attach_payment_columns() already
    #      applied (Gokwik Transaction Report downstream processor,
    #      COD-courier override, ...). A combined split-payment label
    #      ("Delhivery COD, PayU") is reduced to its COD leg first via
    #      cod_component_of_gateway_label() - same convention _gateway_
    #      group()/_display_gateway_label() below already use for this
    #      module's other Payment-Gateway-labelled output, so this sheet's
    #      single order-level row picks the same leg those do.
    #   2. receipt_detail_df's raw consolidated-ledger source (the
    #      uploaded gateway FILE's own label, e.g. "Gokwik" before
    #      downstream refinement) - only reached when Payment Provider/
    #      Gateway are both genuinely blank for this order (a run where
    #      attach_payment_columns() wasn't wired in at all, or an older
    #      saved period predating it - disclosed, not guessed around).
    #   3. expected_gateway_for_order()'s own best-effort guess (Payment
    #      Method text + delivery_partner only) - last resort, only for an
    #      order with no settlement row and no resolved Payment Provider/
    #      Gateway/receipt-ledger source at all.
    def _resolved_provider(row):
        for col in ("Payment Provider", "Gateway"):
            val = row.get(col)
            if pd.notna(val):
                text = str(val).strip()
                if text and text.lower() not in ("nan", "none", "#n/a", "na"):
                    return cod_component_of_gateway_label(text)
        return None

    def _attributed_gateway(row):
        resolved = _resolved_provider(row)
        if resolved:
            return resolved
        if pd.notna(row.get("payment_gateway")) and str(row["payment_gateway"]).strip():
            return row["payment_gateway"]
        return expected_gateway_for_order(row.get("payment_method"), row.get("delivery_partner"), gateway_configs or [])

    df["Payment Gateway"] = df.apply(_attributed_gateway, axis=1)

    # 2026-09-06 (round 19, client-reported item 1): Gateway Amount now
    # prefers reco_df's own (canonical, round-10-rescue-inclusive)
    # receipt_amount over recon_status_df's independently-computed one -
    # see the merge above for why they can diverge. An order genuinely
    # absent from its own courier's COD report (e.g. #29284 - "delivered
    # but not reflecting") has receipt_amount == 0 in BOTH sources, so it
    # correctly stays out of this column either way - only orders where
    # reco_df's own figure is more complete (or simply different) than
    # recon_status_df's change.
    gateway_amount = df["_reco_receipt_amount"] if has_reco_receipt_amount else df["receipt_amount"]
    gateway_amount = gateway_amount.fillna(0.0)

    # 2026-09-06 (round 19, client-reported item 4): "Order Date" used to
    # be reco_df's own raw "created_at" text, passed straight through - for
    # a source file whose date column carries an explicit UTC/IST offset
    # (Shopify-style "...+05:30"), this stayed as literal object/string
    # data all the way to the export step, so engine.formatting.strip_tz()
    # (which only strips tz off an actual tz-aware DATETIME column - see
    # its own docstring) had nothing to act on, and the raw "+0530" text
    # showed up in the exported cell. Parsing it into a real datetime here
    # lets strip_tz() do its job at export time, same as every other date
    # column in this workbook. Scoped to just this sheet's own "Order
    # Date" column, not reco_df's "created_at" everywhere else, since nothing
    # else was reported showing this problem.
    order_date = pd.to_datetime(df["created_at"], errors="coerce")
    if isinstance(order_date.dtype, pd.DatetimeTZDtype):
        order_date = order_date.dt.tz_localize(None)

    # 2026-09-06 (round 22, client-reported, order #33225): the "Order
    # Date" cell exported below (see the "out" DataFrame just below) used
    # to keep its own HH:MM:SS time component (e.g. 31-Jul-2026 12:18:19)
    # even though it DISPLAYS as a date-only value (SHEET_SCOPED_DATE_
    # COLUMNS in engine/formatting.py formats this sheet's "Order Date" as
    # DD-MM-YYYY) - a display format never changes a cell's underlying
    # value, so the live "Days Pending" formula below (Report Period End
    # Date cell minus Order Date cell, both real cells on this sheet)
    # still subtracted that hidden time fraction: an order placed at
    # 12:18:19 on the very same calendar day the reconciliation period
    # ends produced -0.51 "days" pending instead of 0. Client's own words:
    # "Both dates should be considered as date-only values." Normalizing
    # the value itself here (not just its export display format) means
    # the live formula now always subtracts two midnight timestamps - this
    # already matched what _days_pending_numeric (used only for sorting,
    # a few lines below) computed all along; only the exported CELL value
    # (what the formula actually reads) didn't.
    order_date = order_date.dt.normalize()

    # 2026-09-06 (round 19, client-reported item 5): "Days Pending" is
    # rebuilt entirely, replacing recon_status_df's own "days_pending"
    # (engine.bank.classify_order_bank_status()'s pending_clock/days_since_
    # ref, both anchored to pd.Timestamp.now() by default - "today"). The
    # client's own words: "the Days Pending calculation should be: Report
    # Period End Date - Order Placed Date... based on the selected
    # reconciliation/report period, not the current date." period_end_date
    # is that period's own end date (see this function's own docstring);
    # falls back to today only for a caller that has none to pass at all.
    #
    # This is written as a LIVE EXCEL FORMULA on each row - "for better
    # visibility and user understanding," per the client's own request -
    # referencing that SAME row's own "Report Period End Date" and "Order
    # Date" cells (both printed columns on this sheet, so the calculation
    # is fully auditable from the sheet itself, not a value computed
    # invisibly in Python). Column letters are derived from DETAIL_COLUMNS
    # itself (not hardcoded) so a future reordering of these columns can't
    # silently point the formula at the wrong cell.
    #
    # Note, disclosed: this replaces classify_order_bank_status()'s own
    # days_pending for THIS sheet only. That field is also used on other
    # pages (e.g. Order Lookup's ageing, the interactive Dashboard) which
    # weren't part of this request and aren't touched here - if any of
    # those show an implausible day count too, that would need its own,
    # separately-verified fix (see the delivery note for the specific,
    # unresolved 241-vs-~37-day example raised this round).
    period_end = pd.to_datetime(period_end_date, errors="coerce") if period_end_date is not None else pd.Timestamp.now()
    if isinstance(period_end, pd.Series):
        period_end = period_end.iloc[0] if len(period_end) else pd.NaT
    if pd.isna(period_end):
        period_end = pd.Timestamp.now()
    if getattr(period_end, "tzinfo", None) is not None:
        period_end = period_end.tz_localize(None)
    period_end = period_end.normalize()

    order_date_col_letter = get_column_letter(DETAIL_COLUMNS.index("Order Date") + 1)
    period_end_col_letter = get_column_letter(DETAIL_COLUMNS.index("Report Period End Date") + 1)

    # Numeric value kept around purely to sort by (see below) - the actual
    # exported "Days Pending" cell is the formula string, assigned AFTER
    # sorting so each formula's row reference matches its final position.
    _days_pending_numeric = (period_end.normalize() - order_date.dt.normalize()).dt.days

    out = pd.DataFrame({
        "Order ID": df["order_id"],
        "Payment Gateway": df["Payment Gateway"],
        "Payment/Transaction Reference (UTR)": df["utr"],
        "Order Date": order_date,
        "Report Period End Date": period_end,
        "Payment/Receipt Date": df["receipt_date"],
        "Order Amount": df["total"],
        "Gateway Amount": gateway_amount,
        "Settlement Amount": df["_settlement_amount"],
        "Actual Settlement Date": df["receipt_date"],
        "Bank Credit Date": df["bank_credit_date"],
        "Settlement Status": df["Reconciliation Category"],
        "Days Pending": _days_pending_numeric,
    })
    out = out.sort_values("Days Pending", ascending=False, na_position="last").reset_index(drop=True)
    out["Days Pending"] = [
        f"={period_end_col_letter}{row_num}-{order_date_col_letter}{row_num}"
        for row_num in range(2, len(out) + 2)
    ]
    return out


def _gateway_group(label, gateway_configs):
    # 2026-09-06 (round 16, order #30456): engine.attribution.combine_
    # prepaid_and_cod_label() can now hand this a COMBINED "<courier>
    # COD, <prepaid provider>" label for a genuine part-prepaid/part-COD
    # order (e.g. "Delhivery COD, PayU"). Group/display purposes both
    # care about the COD leg specifically (the prepaid leg was already
    # collected at checkout, so the money genuinely still pending is
    # always the COD leg) - see cod_component_of_gateway_label()'s own
    # docstring. A bare (non-combined) label passes through unchanged.
    label = cod_component_of_gateway_label(label)
    text = str(label or "").strip().lower()
    if text == "cod":
        return "COD"
    if text in ("prepaid", "unknown"):
        return "Prepaid"
    for cfg in gateway_configs or []:
        if cfg["label"] == label:
            return "COD" if str(cfg.get("payment_mode", "")).lower() == "cod" else "Prepaid"
    if label in _COD_SETTLEMENT_PENDING_PHRASES:
        return "COD"
    return "Prepaid"


def _display_gateway_label(label):
    """
    Client-reported 2026-08-31 (item 1): engine/attribution.py deliberately
    keeps "Gateway"/"Payment Provider" as the RAW downstream-processor
    string it finds in the Gokwik Transaction Report (lowercase - "payu",
    "easebuzz", ...; see that module's own docstring), but the client's
    own Settlement Pending Summary sheet always shows the capitalised form
    ("Payu", never "payu"/"Gokwik"). Only capitalise a label that's
    ENTIRELY lowercase - a COD courier label ("Delhivery COD") or
    "Razorpay" is already correctly cased at the source, and blindly
    title-casing it would mangle "COD" into "Cod".

    2026-09-06 (round 16, order #30456): also collapses a COMBINED
    "<courier> COD, <prepaid provider>" label (see _gateway_group()'s own
    note just above) down to its COD leg first, for the same reason - the
    Settlement Pending Summary sheet's Payment Gateway column names WHICH
    courier the outstanding money is still with, which for a combined
    label is always the COD leg.
    """
    label = cod_component_of_gateway_label(label)
    label = str(label or "").strip()
    return label.capitalize() if label and label.islower() else label


def _resolved_gateway_raw(df):
    """
    Per-row "Payment Provider", falling back to "Gateway" only where
    Payment Provider is itself blank for THAT row - the same priority
    order build_settlement_pending_report()'s own _resolved_provider()
    already uses, above.

    2026-09-06 (round 22, client-reported, order #29284): pending_amount_
    by_order() and settlement_pending_summary_by_gateway() below used to
    pick whichever of the two columns EXISTS at all ("Gateway" if the
    column is present, else "Payment Provider"), not whichever actually
    HAS a value for a given row. Once engine.attribution's order_id_
    gateway_lookup ("Gateway" - engine.attribution.build_payment_gateway_
    lookups()'s first return value) started being merged onto reco_df,
    "Gateway" became present on every row - but it is only ever populated
    from the courier/gateway's OWN settlement file, so it is blank
    precisely for an order genuinely absent from that file (order #29284:
    correctly delivered by Shiprocket, but not yet reflecting in
    Shiprocket's own COD report). That order's "Payment Provider" column
    (engine.attribution.attach_payment_columns()'s COD-courier-delivered-
    by fallback - see that function's own 2026-08-31/round-8 note) already
    correctly says "Shiprocket COD" regardless, exactly like the Reco
    working and Settlement Pending Detail sheets - but the old column-
    exists check ignored it and treated the order as fully unresolved.
    Coalescing per row (not per column) fixes both call sites at once and
    can never disagree with build_settlement_pending_report()'s own
    resolution again.
    """
    def _clean(series):
        cleaned = series.astype(object).where(series.notna(), None)
        return cleaned.map(
            lambda v: None if v is None or not str(v).strip()
            or str(v).strip().lower() in ("nan", "none", "#n/a", "na")
            else str(v).strip()
        )

    has_pp = "Payment Provider" in df.columns
    has_gw = "Gateway" in df.columns
    if not has_pp and not has_gw:
        return pd.Series(None, index=df.index, dtype=object)

    pp_clean = _clean(df["Payment Provider"]) if has_pp else pd.Series(None, index=df.index, dtype=object)
    if not has_gw:
        return pp_clean
    gw_clean = _clean(df["Gateway"])
    return pp_clean.where(pp_clean.notna(), gw_clean)


def _is_report_gap_query(query_col, index):
    """
    True for a query that already specifically flags "genuinely absent
    from this order's own COD/gateway report" - engine.reco.py::
    apply_cod_report_gap_query()'s courier-specific "{courier} COD
    Delivered but amount not reflecting in {courier} COD Report" text, and
    refine_queries_with_settlement_status()'s own "{courier} Setlment
    pending not reflecting" for the same underlying situation - as opposed
    to an ordinary "money already reflects somewhere, still awaiting bank
    credit" pending query. Both phrasings share the substring "not
    reflecting"; the bare generic "COD Delivered Amount not Received" is
    matched too, for parity with pending_amount_by_order()'s pre-existing
    check (in practice apply_cod_report_gap_query() already rewrites that
    generic text away for every COD-courier-delivered order by the time
    this runs, so it only ever matches a genuinely-unresolved order here -
    already routed to its own "unresolved" bucket before this check is
    reached at either call site).

    Extracted out of pending_amount_by_order() (round 19) purely so its
    AMOUNT-selection rule (receipt_amount vs Total) has one named home -
    this is deliberately the BROAD match (covers both phrasings). It is
    NOT used for settlement_pending_summary_by_gateway()'s own "not
    reflecting" CATEGORY (round 22) - see _is_cod_report_absent_query()
    below for why that needs the NARROWER, single-phrasing match instead.
    """
    query_col = query_col if query_col is not None else pd.Series("", index=index)
    text = query_col.fillna("")
    return (
        (text == "COD Delivered Amount not Received")
        | text.str.contains(BROAD_NOT_REFLECTING_FRAGMENT, case=False, regex=False)
    )


def _is_cod_report_absent_query(query_col, index):
    """
    True ONLY for engine.reco.py::apply_cod_report_gap_query()'s own,
    specific phrasing - "{courier} COD Delivered but amount not reflecting
    in {courier} COD Report" (order #29284's exact text: delivered by a
    recognised COD courier, but genuinely absent from THAT COURIER'S OWN
    report). Deliberately narrower than _is_report_gap_query() above -
    round 16's own regression test (order #30456) confirmed the client
    expects the OTHER "not reflecting" phrasing, refine_queries_with_
    settlement_status()'s more generic "{courier} Setlment pending not
    reflecting" (has_settlement_row False for a reason other than "this
    specific courier's report has no row" - e.g. a split-payment leg still
    awaiting its own settlement row), to keep grouping under that
    courier's own ordinary row in the Settlement Pending Summary, not the
    new distinct category below. Matching on the more specific substring
    "delivered but amount not reflecting in" (present only in apply_cod_
    report_gap_query()'s own f-string) keeps the two phrasings apart.
    """
    query_col = query_col if query_col is not None else pd.Series("", index=index)
    text = query_col.fillna("")
    return text.str.contains(COD_REPORT_ABSENT_QUERY_FRAGMENT, case=False, regex=False)


def split_payment_pending_topups(reco_df, consolidated_df, bank_ledger_df, gateway_configs):
    """
    2026-09-06 (round 17) - client-reported direct follow-up to order
    #30456: "Settlement Pending & Exceptions (by Payment Gateway) should
    be update[d]" once the Query text correctly shows a split-payment
    order's still-outstanding leg (see engine.reco.refine_split_payment_
    queries()). Root cause this needs its own, ADDITIVE pass rather than a
    change to settlement_pending_summary_by_gateway()'s own filter: that
    function's "is this order pending" test reads reco_df's own
    receipt_status (Recipt Remark), which for a genuine split-payment
    order with ONE leg already bank-matched (Delhivery COD here) reads
    "Received" at the ORDER level - engine.bank.classify_order_bank_
    status()'s order-level bank_matched aggregate masks the OTHER leg
    (PayU) being genuinely still bank-pending (see that function's own
    docstring, and engine.bank.resolve_split_payment_leg_status()'s, for
    the full root-cause story). Such an order is invisible to that
    function's pending filter entirely - it never even reaches the
    by-gateway groupby - so this is a supplemental pass that tops up the
    summary with exactly the still-pending legs that filter can't see,
    rather than trying to rederive that filter's own (already-correct
    for every non-split order) logic.

    Reuses engine.bank.resolve_split_payment_leg_status() for the
    per-leg bank-matched signal, and sums each still-pending leg's own
    reported amount straight off consolidated_df (net of that leg's own
    "deduction" column if present - the same "final payment" concept used
    everywhere else in this engine, e.g. build_cod_settlement_batches()
    above) - NOT reco_df's own receipt_amount, which is an order-level
    total across both legs and would misstate a single leg's own
    exposure.

    Deliberately narrow, disclosed scope: only tops up an order that is
    NOT already counted by settlement_pending_summary_by_gateway()'s own
    pending filter (receipt_status still "Not Received"/"Received Bank
    settlement pending") - an order that's already in that filter is left
    to its existing (COD-leg-only) attribution rather than risking a
    double-count here; a split-payment order where BOTH legs are still
    unresolved (the rarer case where the order-level classification isn't
    masked at all) is not specifically re-attributed by this pass -
    disclosed, not silently guessed around.

    Returns: order_id | Group ("COD"/"Prepaid") | Payment Gateway
    (display label, same convention as settlement_pending_summary_by_
    gateway()) | Amount Pending - one row per still-pending leg.
    """
    cols = ["order_id", "Group", "Payment Gateway", "Amount Pending"]
    if reco_df is None or reco_df.empty or consolidated_df is None or consolidated_df.empty:
        return pd.DataFrame(columns=cols)

    gateway_col = "Gateway" if "Gateway" in reco_df.columns else (
        "Payment Provider" if "Payment Provider" in reco_df.columns else None
    )
    if gateway_col is None:
        return pd.DataFrame(columns=cols)

    gw_text = reco_df[gateway_col].astype(str).str.strip()
    is_combined = gw_text.str.contains(", ", regex=False)
    if not is_combined.any():
        return pd.DataFrame(columns=cols)

    already_pending = pd.Series(False, index=reco_df.index)
    if "receipt_status" in reco_df.columns:
        already_pending = reco_df["receipt_status"].fillna("").isin(
            ["Not Received", "Received Bank settlement pending"]
        )

    candidates = reco_df.loc[is_combined & (~already_pending)]
    if candidates.empty:
        return pd.DataFrame(columns=cols)

    leg_status_by_order = resolve_split_payment_leg_status(
        candidates, consolidated_df, bank_ledger_df, gateway_configs,
    )
    if not leg_status_by_order:
        return pd.DataFrame(columns=cols)

    gw_by_order = dict(zip(candidates["order_id"].astype(str), gw_text.loc[candidates.index]))
    total_by_order = dict(zip(
        candidates["order_id"].astype(str),
        candidates["total"].fillna(0.0) if "total" in candidates.columns else pd.Series(0.0, index=candidates.index),
    ))
    receipt_by_order = dict(zip(
        candidates["order_id"].astype(str),
        candidates["receipt_amount"].fillna(0.0) if "receipt_amount" in candidates.columns else pd.Series(0.0, index=candidates.index),
    ))

    payments = consolidated_df[~consolidated_df["is_refund"]].copy()
    payments["order_id"] = payments["order_id"].astype(str)
    payments["_net"] = payments["amount"] - payments.get("deduction", 0.0)
    net_by_order_source = payments.groupby(["order_id", "source"])["_net"].sum().to_dict()

    def _leg_amount(oid, label):
        # A leg that already has a real consolidated_df row (settled or
        # not) uses that row's own net reported amount - the most
        # accurate figure available. A leg with NO row at all yet (e.g.
        # order #30456's PayU leg - has_confirmed_prepaid_transaction()
        # can confirm a real checkout-time payment from the Gokwik
        # Transaction Report before that gateway's own settlement file
        # has posted at all, see engine.attribution's own docstring) has
        # nothing in consolidated_df to sum - fall back to the order's
        # own genuinely-missing amount (Total minus what Reco working's
        # receipt_amount already reflects, i.e. the OTHER leg's
        # contribution), the same "nothing reported collected yet, use
        # the order's own at-stake figure" convention used everywhere
        # else in this module (see pending_amount_by_order()'s own
        # docstring).
        key = (oid, label)
        if key in net_by_order_source:
            return net_by_order_source[key]
        return max(total_by_order.get(oid, 0.0) - receipt_by_order.get(oid, 0.0), 0.0)

    rows = []
    for oid, legs in leg_status_by_order.items():
        gw = gw_by_order.get(oid)
        cod_label = cod_component_of_gateway_label(gw)
        prepaid_label = prepaid_component_of_gateway_label(gw)
        if not prepaid_label:
            continue  # not actually a combined label for this order

        if legs.get("cod_matched") is not True:
            amt = _leg_amount(oid, cod_label)
            if amt:
                rows.append({
                    "order_id": oid, "Group": "COD",
                    "Payment Gateway": _display_gateway_label(cod_label),
                    "Amount Pending": round(float(amt), 2),
                })
        if legs.get("prepaid_matched") is not True:
            amt = _leg_amount(oid, prepaid_label)
            if amt:
                rows.append({
                    "order_id": oid, "Group": "Prepaid",
                    "Payment Gateway": _display_gateway_label(prepaid_label),
                    "Amount Pending": round(float(amt), 2),
                })

    return pd.DataFrame(rows, columns=cols)


def pending_amount_by_order(reco_df, gateway_configs):
    """
    Client-reported 2026-09-05: the Dashboard/Executive Summary "Settlement
    Pending Amount" headline figure only showed Delhivery COD's pending
    money - Shiprocket COD's ₹1,17,670.47 and Payu's pending amount were
    both missing. Root cause: that headline number (engine/summary.py::
    headline_totals()'s "Settlement pending") was sourced from engine/reco.
    py::attach_settlement_pending(), which itself summed build_settlement_
    pending_report()'s "Gateway Amount" - a figure read off recon_status_df
    (engine.bank.classify_order_bank_status()'s output), filtered to
    Gateway Amount > 0.01. Two different orders fall through that filter
    for two different reasons:
      - Shiprocket COD orders engine.reco.attach_pending_cod_receipts()
        rescues (round 10) - recon_status_df's OWN receipt_amount never
        learns about that rescue (see that function's own docstring), so
        Gateway Amount stays 0 for them there even though reco_df's own
        receipt_amount is correctly populated.
      - Payu (or any prepaid gateway) orders with NOTHING collected yet at
        all (has_settlement_row False, genuinely no settlement file
        resolved) - excluded from that sum BY DESIGN, because the old
        "Settlement pending" figure was originally built only to net Net
        Settlement (subtracting money already counted as received but not
        yet bank-credited - see attach_settlement_pending()'s own
        docstring), not to show total outstanding exposure.

    Extracted out of settlement_pending_summary_by_gateway()'s own row loop
    (client-reported 2026-08-31, item 1) so BOTH the by-gateway Settlement
    Pending Summary sheet and the single Dashboard/Executive Summary
    headline figure are now built from the exact same per-order amount -
    they can no longer drift apart the way this report and the client's
    own reference workbook once did (see that function's docstring for the
    full history). Returns a Series aligned to reco_df's own index: the
    order's outstanding exposure for every row whose receipt_status is
    still-pending (bare "Not Received", or the round-10 "Received Bank
    settlement pending"), 0.0 for every other row (nothing here is a
    still-pending, at-risk order, so nothing here is missing from the
    Reconciliation Category classification's own COD_BANK_MATCHED/
    PREPAID_BANK_MATCHED/COD_NOT_DELIVERED buckets).

    Same amount rule as the summary sheet: a COD courier's own row (not
    the generic catch-all query) uses receipt_amount (what the courier's
    own report already shows collected, whether or not it's reached the
    bank yet); every other pending row - a resolved prepaid gateway, an
    unresolved/catch-all COD row, a genuinely-uncollected Payu order - uses
    the order's own Total, since nothing has been reported collected for
    those yet and Total is the only meaningful "at stake" figure.
    """
    if reco_df is None or reco_df.empty or "receipt_status" not in reco_df.columns:
        return pd.Series(dtype=float)

    amount = pd.Series(0.0, index=reco_df.index)
    pending = reco_df["receipt_status"].fillna("").isin(["Not Received", "Received Bank settlement pending"])
    if not pending.any():
        return amount

    df = reco_df.loc[pending]
    # 2026-09-06 (round 22): resolved per row (Payment Provider first,
    # Gateway as fallback), not per column - see _resolved_gateway_raw()'s
    # own docstring for the order #29284 root-cause story this replaces.
    gateway_raw = _resolved_gateway_raw(df)
    unresolved = gateway_raw.isna()

    query_col = df["query"] if "query" in df.columns else pd.Series("", index=df.index)
    # 2026-09-06 (round 19, client-reported): "there is small difference
    # 3136 coming because [this] sheet captured only 'amount reflecting in
    # COD report not setled to bank' - it also include COD Delivered but
    # amount not reflecting in COD Report and bank setlment not done" -
    # order #29284's own category. Root cause: this used to match ONLY the
    # literal generic text "COD Delivered Amount not Received" as "nothing
    # genuinely collected yet, use Total" - it didn't recognise either of
    # the two courier-specific "not reflecting" phrasings engine.reco.py
    # can also produce for that exact same "genuinely absent" situation -
    # see _is_report_gap_query()'s own docstring (round 22: extracted out
    # to here to a shared helper) for the exact rule.
    is_catch_all_query = _is_report_gap_query(query_col, df.index)

    total_col = df["total"] if "total" in df.columns else pd.Series(0.0, index=df.index)
    receipt_col = df["receipt_amount"] if "receipt_amount" in df.columns else pd.Series(0.0, index=df.index)

    for idx in df.index:
        if unresolved.loc[idx]:
            amount.loc[idx] = total_col.loc[idx]
            continue
        grp = _gateway_group(gateway_raw.loc[idx], gateway_configs)
        if grp == "COD" and not is_catch_all_query.loc[idx]:
            amount.loc[idx] = receipt_col.loc[idx]
        else:
            amount.loc[idx] = total_col.loc[idx]

    return amount.fillna(0.0)


def settlement_pending_summary_by_gateway(reco_df, gateway_configs, consolidated_df=None, bank_ledger_df=None):
    """
    consolidated_df/bank_ledger_df (2026-09-06, round 17, optional -
    existing callers that omit them keep this function's exact prior
    behaviour): when supplied, split_payment_pending_topups()'s rows are
    folded in alongside this function's own SUMIFS-style rows - a genuine
    split-payment order (e.g. #30456) whose order-level receipt_status
    already reads "Received" (masking a still-bank-pending leg - see that
    function's own docstring for the full root-cause story) would
    otherwise never reach this summary at all. See that function's
    docstring for the exact, disclosed scope of what it tops up.

    Rewritten 2026-08-31 (client-reported, item 1): the client's own
    Settlement Pending Summary sheet didn't match this report's numbers at
    all - not just a rounding difference, but a genuinely different
    Payment Gateway label ("Gokwik" here vs "Payu" there) and a different
    Amount Pending for every COD row. Root cause, confirmed by extracting
    the client's own workbook's ACTUAL cell formulas (not just its output
    values): their sheet is built directly off the Reco working sheet's
    own Recipt Remark/Query/Payment Provider columns with plain SUMIFS -
    e.g. Delhivery COD's Amount Pending is
    `SUMIFS(receipt_amount, Recipt Remark="Not Received", Query="Delhivery
    COD Setlment pending") + SUMIFS(total, Payment Provider="Delhivery
    COD", Query="COD Delivered Amount not Received")` - NOT this report's
    previous approach (a separately-computed "Settlement Amount" netted of
    that gateway's own deductions/refunds, still correct and still used
    for the Settlement Pending DETAIL sheet's own "how much will actually
    land" figure - see build_settlement_pending_report()'s docstring -
    just not what this SUMMARY sheet actually shows in their workbook).

    Rebuilt to mirror those exact formulas off reco_df directly instead:
      - A row is "pending" here iff its own Recipt Remark
        (engine.reco.attach_receipt_status()'s "receipt_status") is the
        bare "Not Received" - by construction of that function, this ONLY
        happens for orders classify_order_bank_status() puts in one of the
        three settlement-pending categories (COD - Delivered & Settlement
        Pending / Prepaid - Payment Received & Settlement Pending /
        Exception - Manual Reconciliation Required), never for a genuine
        RTO/Cancelled "nothing was ever collected" order - so this needs
        no separate category lookup of its own.
      - Grouped by the SAME resolved Payment Provider/Gateway attribution
        already shown on Reco working - one canonical attribution shared
        with that sheet's own Query text, rather than a second,
        independent attribution mechanism that could (and did) disagree
        with it. (2026-09-06, round 22: resolved per ROW, "Payment
        Provider" first and "Gateway" only as a per-row fallback, via
        _resolved_gateway_raw() - see that helper's own docstring for why
        preferring whichever COLUMN merely exists, the previous rule,
        wrongly called a courier-attributed-but-not-yet-reflecting order
        like #29284 "unresolved".)
      - A COD courier's own row sums receipt_amount (what the courier has
        already reported collecting, even if not yet bank-credited); a
        resolved prepaid gateway's row, and the catch-all "COD Delivered
        Amount not Received" row for orders with no resolvable gateway at
        all, both sum Total instead (nothing has been reported collected
        for these at all, so Total is the only meaningful "at stake"
        figure) - confirmed against the client's own formulas, which use
        the receipt_amount column only for the COD-courier-specific SUMIFS
        and Total everywhere else.
      - A handful of orders the client's workbook attributes to a specific
        COD courier via the Payment Provider column even though their
        Query is the generic catch-all (not that courier's own "pending"
        phrase) get folded into that courier's row too, using Total for
        just those rows - reproduced here the same way.

    Orders Pending is a live COUNTIF-equivalent of the same criteria as
    Amount Pending, rather than the client's own workbook's literal
    numbers - those turned out to be static, manually-typed figures that
    no longer match live rows in their own SUMIFS formulas (e.g. their own
    sheet shows "8" for the catch-all row while an actual SUMIFS with the
    same remark/query criteria used for the amount would count only 4)
    - a live count is more correct and stays right as the underlying data
    changes, which is the whole reason the client asked for formula-driven
    figures throughout this workbook.

    Disclosed gap, not guessed around: this still won't reproduce the
    client's numbers to the last rupee - a small number of Lost/Status
    Undefined orders in their own July data get a specific courier
    attribution (via Payment Provider) that engine/attribution.py's own
    resolution logic doesn't currently derive for those particular orders
    (it resolves cleanly for the ~640 genuinely-pending orders, just not
    for a handful of already-anomalous ones) - confirmed to affect under
    1% of the total Amount Pending figure on the client's own July data.
    Fixing that is an upstream Payment Provider/Gateway attribution
    question (engine/attribution.py), not a Settlement Pending Summary
    question, and safer to disclose than to guess a one-off rule from a
    handful of examples.
    """
    if reco_df is None or reco_df.empty or "receipt_status" not in reco_df.columns:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    # 2026-09-06 (round 17): computed up front, independent of the
    # order-level `pending` filter just below, since its whole point is to
    # surface split-payment orders THAT FILTER MISSES (see this function's
    # own note above and split_payment_pending_topups()'s docstring) -
    # must not be short-circuited by an early "nothing order-level pending"
    # return.
    topups = split_payment_pending_topups(reco_df, consolidated_df, bank_ledger_df, gateway_configs)

    df = reco_df.copy()
    # 2026-09-04 (round 10): engine.reco.attach_receipt_status() now also
    # returns "Received Bank settlement pending" (not just "Not Received")
    # for a still-settlement-pending order whose receipt_amount is ALREADY
    # populated (a COD order the courier's own report already confirms
    # collected, just not yet bank-credited - see that function's own
    # docstring, point 2). Both labels are the same underlying "still
    # settlement-pending" set classify_order_bank_status() puts an order
    # in - only the wording differs based on whether money has already
    # been reported collected - so both must count here, or an order like
    # this would silently vanish from this summary entirely (regression
    # caught in engine's own test suite) rather than simply moving group/
    # amount as its receipt_amount now correctly reflects.
    pending = df["receipt_status"].fillna("").isin(["Not Received", "Received Bank settlement pending"])
    if not pending.any():
        rows = topups.rename(columns={"order_id": "_oid"}) if not topups.empty else pd.DataFrame(columns=["_oid", "Group", "Payment Gateway", "Amount Pending"])
    else:
        df = df.loc[pending].copy()

        # 2026-09-06 (round 22): resolved per row (Payment Provider first,
        # Gateway as fallback), not per column - see _resolved_gateway_
        # raw()'s own docstring for the order #29284 root-cause story this
        # replaces (a courier-resolved order was falling into the
        # "unresolved" catch-all whenever "Gateway" - populated only from
        # that courier's OWN settlement file - happened to be blank for
        # that specific row).
        gateway_raw = _resolved_gateway_raw(df)
        unresolved = gateway_raw.isna()

        # 2026-09-05: amount now comes from pending_amount_by_order() - the
        # SAME per-order helper engine/reco.py::attach_settlement_pending()
        # calls for the Dashboard/Executive Summary headline figure - so this
        # by-gateway breakdown and that single headline number can never
        # disagree again (see that helper's own docstring for the full
        # client-reported story). Only the Group/Payment Gateway LABELS are
        # still resolved here (this function's own concern - which sheet row
        # an order's amount lands under).
        pending_amounts = pending_amount_by_order(reco_df, gateway_configs)

        # 2026-09-06 (round 22, client-reported, order #29284): a second,
        # NEW category alongside the client's own existing "COD amount
        # reflecting in COD report but pending to bank credit" (the
        # per-courier rows just below, e.g. "Delhivery COD"/"Shiprocket
        # COD") - "COD amount not reflecting in COD report and pending to
        # bank credit", for a COD order that IS correctly attributed to a
        # specific delivering courier (unlike the genuinely-unresolved
        # catch-all above) but whose own COD/settlement report has no row
        # for it at all yet (engine.reco.py::apply_cod_report_gap_query()'s
        # own courier-specific phrasing - see _is_cod_report_absent_query()
        # for exactly which "not reflecting" text this is, and why it's
        # deliberately narrower than pending_amount_by_order()'s own
        # amount-selection check just above). Previously such an order had
        # nowhere of its own to land - it was either silently folded into
        # that courier's ordinary "reported collected, bank credit still
        # pending" row (indistinct from a genuinely-reflecting one) or,
        # before the gateway_raw fix just above, dropped into the generic
        # unresolved bucket instead - neither of which is the client's own
        # distinct category for it. Prepaid is untouched: a resolved
        # prepaid gateway with nothing collected at all already had
        # nowhere more specific than its own gateway row (there is no
        # separate "prepaid report" concept to be missing from), so only
        # the COD group gets this new split. A split-payment order whose
        # STILL-pending leg reads the OTHER "not reflecting" phrasing
        # (refine_queries_with_settlement_status()'s generic "{courier}
        # Setlment pending not reflecting", e.g. order #30456 - round 16's
        # own regression test) deliberately keeps grouping under its own
        # courier row instead, unaffected by this new category.
        query_col = df["query"] if "query" in df.columns else pd.Series("", index=df.index)
        is_gap_query = _is_cod_report_absent_query(query_col, df.index)

        groups, labels = [], []
        for idx in df.index:
            if unresolved.loc[idx]:
                groups.append("COD")
                labels.append("COD Delivered Amount not Received")
                continue
            raw_label = gateway_raw.loc[idx]
            grp = _gateway_group(raw_label, gateway_configs)
            if grp == "COD" and is_gap_query.loc[idx]:
                groups.append("COD")
                labels.append(NOT_REFLECTING_LABEL)
                continue
            groups.append(grp)
            labels.append(_display_gateway_label(raw_label))

        df["Group"] = groups
        df["Payment Gateway"] = labels
        df["_amount"] = pending_amounts.loc[df.index].fillna(0.0)

        rows = df.rename(columns={"order_id": "_oid", "_amount": "Amount Pending"})[["_oid", "Group", "Payment Gateway", "Amount Pending"]]
        if not topups.empty:
            rows = pd.concat(
                [rows, topups.rename(columns={"order_id": "_oid"})[["_oid", "Group", "Payment Gateway", "Amount Pending"]]],
                ignore_index=True,
            )

    if rows.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    grouped = rows.groupby(["Group", "Payment Gateway"]).agg(**{
        "Orders Pending": ("_oid", "count"),
        "Amount Pending": ("Amount Pending", "sum"),
    }).reset_index()
    grouped["Amount Pending"] = grouped["Amount Pending"].round(2)

    return grouped[SUMMARY_COLUMNS].sort_values(["Group", "Payment Gateway"]).reset_index(drop=True)


def reconciliation_health_by_gateway(reco_df, recon_status_df, receipt_detail_df, gateway_configs):
    """
    The dashboard-level bifurcation requested: for every gateway/delivery
    partner, how much is already received in bank vs pending settlement vs
    pending bank matching vs a genuine exception - one row per gateway,
    covering ALL orders (not just the pending ones the detail report
    above focuses on), so this is the "so how are we doing overall, by
    gateway" view.
    """
    from .bank import COD_BANK_MATCHED, PREPAID_BANK_MATCHED

    cols = ["Payment Gateway", "Group", "Received in Bank", "Pending Settlement",
            "Pending Bank Matching", "Exception", "Total"]
    if reco_df is None or reco_df.empty or recon_status_df is None or recon_status_df.empty:
        return pd.DataFrame(columns=cols)

    df = recon_status_df.copy()
    reco_cols = ["order_id", "payment_method", "total"]
    if "delivery_partner" in reco_df.columns:
        reco_cols.append("delivery_partner")
    reco = reco_df[reco_cols].copy()
    reco["order_id"] = reco["order_id"].astype(str)
    df["order_id"] = df["order_id"].astype(str)
    df = df.merge(reco, on="order_id", how="left")

    if receipt_detail_df is not None and not receipt_detail_df.empty:
        rd = receipt_detail_df[["order_id", "payment_gateway"]].copy()
        rd["order_id"] = rd["order_id"].astype(str)
        df = df.merge(rd, on="order_id", how="left")
    else:
        df["payment_gateway"] = None

    def _attributed_gateway(row):
        if pd.notna(row.get("payment_gateway")) and str(row["payment_gateway"]).strip():
            return row["payment_gateway"]
        if row["Reconciliation Category"] in PENDING_CATEGORIES:
            return expected_gateway_for_order(row.get("payment_method"), row.get("delivery_partner"), gateway_configs or [])
        return None  # "not delivered, no receipt expected" - not attributable, and shouldn't be

    df["Payment Gateway"] = df.apply(_attributed_gateway, axis=1)
    df = df[df["Payment Gateway"].notna()]
    if df.empty:
        return pd.DataFrame(columns=cols)

    df["Group"] = df["Payment Gateway"].apply(lambda g: _gateway_group(g, gateway_configs))
    is_matched = df["Reconciliation Category"].isin({COD_BANK_MATCHED, PREPAID_BANK_MATCHED})
    is_settlement_pending = df["Reconciliation Category"].isin({COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING}) & (~df["has_settlement_row"])
    is_bank_matching_pending = df["Reconciliation Category"].isin({COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING}) & df["has_settlement_row"]
    is_exception = df["Reconciliation Category"] == EXCEPTION_MANUAL_REVIEW

    # Explicit per-row amount columns (rather than aggregating inside a
    # lambda) so the "how much" figure for each bucket is unambiguous: the
    # order's still-outstanding value while nothing has been collected yet
    # (Pending Settlement), and the gateway-reported receipt amount once
    # something HAS been collected/remitted but not yet bank-matched
    # (Pending Bank Matching) or already matched (Received in Bank).
    df["_received_amt"] = df["receipt_amount"].where(is_matched, 0.0)
    df["_pending_settlement_amt"] = df["total"].where(is_settlement_pending, 0.0)
    df["_pending_bank_amt"] = df["receipt_amount"].where(is_bank_matching_pending, 0.0)
    df["_exception_amt"] = df["receipt_amount"].where(is_exception, 0.0)

    grouped = df.groupby(["Payment Gateway", "Group"]).agg(**{
        "Received in Bank": ("_received_amt", "sum"),
        "Pending Settlement": ("_pending_settlement_amt", "sum"),
        "Pending Bank Matching": ("_pending_bank_amt", "sum"),
        "Exception": ("_exception_amt", "sum"),
    }).reset_index()

    for c in ["Received in Bank", "Pending Settlement", "Pending Bank Matching", "Exception"]:
        grouped[c] = grouped[c].fillna(0.0).round(2)
    grouped["Total"] = (grouped["Received in Bank"] + grouped["Pending Settlement"]
                         + grouped["Pending Bank Matching"] + grouped["Exception"]).round(2)

    return grouped[cols].sort_values(["Group", "Payment Gateway"]).reset_index(drop=True)
