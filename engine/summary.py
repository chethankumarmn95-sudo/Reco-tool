"""
summary.py
----------
Layer 4 - the equivalent of your "Reco Summary" sheet. Takes the finished
Reco working table and rolls it up into the management-ready views:
month-wise collection, delivery-status-wise breakup, and open queries.
"""

import pandas as pd


MONTH_CALENDAR_ORDER = {
    name: i for i, name in enumerate([
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ])
}


def month_summary(reco_df):
    """
    Client-reported (2026-08-21): rows came out shuffled ("August,
    December, July, ...") instead of calendar order. Root cause:
    groupby("month") sorts its groups alphabetically by default, and
    "month" (see engine/reco.py's `_derived_month`) only ever holds a bare
    month NAME like "August" - alphabetically, "August" < "December" <
    "July", which is exactly the reported symptom.

    Sorted here instead by each month's own earliest order date when
    "created_at" is available (true calendar order, so a fiscal year
    starting mid-calendar-year, e.g. Apr-Mar, still reads top-to-bottom in
    the order it actually happened) - falling back to plain Jan-Dec order
    only if no date column exists to anchor on at all.
    """
    g = reco_df.groupby("month").agg(
        orders=("order_id", "count"),
        order_value=("total", "sum"),
        receipt=("receipt_amount", "sum"),
        deduction=("total_deduction", "sum"),
        settlement=("settlement_amount", "sum"),
        diff=("diff", "sum"),
    ).reset_index()
    g["collection_rate"] = (g["receipt"] / g["order_value"]).round(4)

    if "created_at" in reco_df.columns:
        earliest_date_by_month = (
            pd.to_datetime(reco_df["created_at"], errors="coerce")
            .groupby(reco_df["month"]).min()
        )
        # .to_dict() first, not map()'d straight off the Series - see
        # engine/bank.py's classify_order_bank_status for the same fix
        # already applied there, with the full explanation: mapping an
        # EMPTY datetime64-typed Series via .map() crashes on some pandas
        # versions/dtype backends ("TypeError: Cannot cast DatetimeArray to
        # dtype float64") - pandas' internal map_array() builds a lookup
        # Series from the mapper and, for an empty mapper, defaults it to
        # float64, then tries to cast the (empty) DatetimeArray into that -
        # a real crash, not a hypothetical one, hit here whenever
        # month_summary() runs against an empty reco_df (client-reported
        # 2026-08-22: uploading a period with no data currently selected/
        # matched threw exactly this, crashing the whole Dashboard page
        # rather than just showing an empty summary). A plain dict sidesteps
        # this pandas dtype-inference path entirely, empty or not.
        g["_sort_key"] = g["month"].map(earliest_date_by_month.to_dict())
    else:
        g["_sort_key"] = g["month"].map(MONTH_CALENDAR_ORDER)

    g = g.sort_values("_sort_key", na_position="last").drop(columns="_sort_key").reset_index(drop=True)
    return g


def status_summary(reco_df):
    g = reco_df.groupby("final_delivery_status").agg(
        orders=("order_id", "count"),
        order_value=("total", "sum"),
        receipt=("receipt_amount", "sum"),
        diff=("diff", "sum"),
    ).reset_index()
    return g


def open_queries(reco_df):
    """
    Client-reported 2026-08-30 (item 4): rebuilt to match the client's own
    reference workbook's "Open queries" sheet, reverse-engineered against
    MOD.xlsx - columns Query | orders | order_value | Recipt | Refund |
    Bank receipt | exposure (display renames applied only on the exported
    copy in views/page_reports.py, same pattern as "Reco working" - this
    function's own column names stay query/orders/order_value/receipt/
    refund/bank_receipt/exposure so views/page_dashboard.py's existing
    "Top Exceptions" chart, which only reads "orders"/"query", is
    unaffected).

    "exposure" used to be diff.sum() (order value minus receipt, before
    any bank-matching or zeroing) - confirmed against MOD's own numbers to
    actually be order_value MINUS "Bank receipt", where "Bank receipt" is
    the sum of the EXPORTED "Bank credit" column - i.e. settlement_amount
    with the item-14 fix's zeroing already applied for any order still
    settlement-pending (see engine/reco.py::attach_settlement_pending's
    docstring and views/page_reports.py's _build_workbook() item-14
    comment). Every query that's still purely a settlement-pending
    category (Delhivery/Shiprocket/Payu Setlment pending, COD Delivered
    Amount not Received, ...) therefore shows Bank receipt = 0 in MOD, and
    exposure = order_value - 0 - exactly what's observed there; a refund/
    RTO/cancellation query instead shows a real (sometimes negative)
    Bank receipt, since those orders aren't zeroed.
    """
    q = reco_df[reco_df["query"] != "Okk"].copy()
    bank_receipt = q["settlement_amount"] if "settlement_amount" in q.columns else 0.0
    if "settlement_pending_amount" in q.columns:
        still_pending = q["settlement_pending_amount"].fillna(0) > 0.01
        bank_receipt = pd.Series(bank_receipt, index=q.index).where(~still_pending, 0.0)
    q["_bank_receipt"] = bank_receipt
    g = q.groupby("query").agg(
        orders=("order_id", "count"),
        order_value=("total", "sum"),
        receipt=("receipt_amount", "sum"),
        refund=("refund_amount", "sum"),
        bank_receipt=("_bank_receipt", "sum"),
    ).reset_index()
    g["exposure"] = (g["order_value"] - g["bank_receipt"]).round(2)
    for c in ["order_value", "receipt", "refund", "bank_receipt"]:
        g[c] = g[c].round(2)
    return g.sort_values("exposure", ascending=False).reset_index(drop=True)


def headline_totals(reco_df):
    """
    Client-reported (2026-08-27): "Net settlement" should ultimately match
    the Bank Credit amount (subject to normal timing differences), but it
    never did - it summed settlement_amount (Receipt - Deduction - Refund)
    for EVERY order, including ones whose money hasn't actually reached the
    bank yet (still sitting in "Settlement Pending" / exception categories -
    see engine/settlement_pending.py::build_settlement_pending_report()).
    That pending money was being counted as "settled" here even though the
    bank hadn't credited it, which is exactly why this figure never tied
    out to the bank statement.

    Fix: reco_df may optionally carry a "settlement_pending_amount" column
    (see engine/reco.py::attach_settlement_pending(), wired in by
    views/page_reconciliation.py and views/page_reports.py right after
    build_settlement_pending_report() runs) - the per-order amount still
    outstanding. When present, it's subtracted from the settlement_amount
    total so "Net settlement" reflects only money that has genuinely
    reached the bank. Reading it via `in reco_df.columns` (not a bare
    reco_df.get(...).sum(), which would error - a missing column returns a
    plain 0, and int has no .sum()) keeps this safe against an older saved
    reco_df from before this column existed: it simply falls back to the
    previous (pre-fix) behaviour for that saved period, a disclosed gap
    rather than a crash.

    Client-reported 2026-08-31 (item 8): "Bank Net Settlement" (this
    function's "Net settlement") still didn't tie to the actual Bank
    Credit, even after the fix above. Root cause: settlement_pending_
    amount (engine/settlement_pending.py::build_settlement_pending_
    report()'s own "Settlement Amount") is, BY DESIGN there, the order's
    full Total for a prepaid order the gateway hasn't even reported
    collecting yet (has_settlement_row False - nothing to net a deduction/
    refund FROM) - correct and meaningful for the Settlement Pending
    Summary/Detail sheets, where the point IS to show the full amount
    still outstanding. But settlement_amount (receipt_amount -
    total_deduction - refund_amount) for that SAME order is already 0
    (nothing received yet) - so a straight subtraction here removed the
    order's full Total from a total that never included it in the first
    place, DOUBLE-COUNTING the shortfall and understating Net Settlement
    by the sum of every such "nothing collected yet" pending order (the
    ENTIRE Payu-style prepaid-pending bucket, confirmed against the
    client's own July data). Fixed by capping each order's pending
    deduction at that same order's own settlement_amount (never below 0) -
    an order with nothing yet counted in settlement_amount now correctly
    contributes 0 to Net Settlement (excluded, not double-subtracted), and
    an order that DOES have a settlement row (e.g. Delhivery COD pending -
    settlement_pending_amount and settlement_amount are numerically
    identical there by construction) still nets to exactly 0 as before -
    Net Settlement now sums to precisely "settlement_amount for every
    order NOT still pending", which is what should tie to Bank Credit.
    The uncapped "Settlement pending" figure shown below is left as the
    full, genuine outstanding exposure (not capped) - only the Net
    Settlement arithmetic changes.
    """
    if "settlement_pending_amount" in reco_df.columns:
        settlement_amount_col = reco_df["settlement_amount"].fillna(0.0)
        pending_raw = reco_df["settlement_pending_amount"].fillna(0.0)
        settlement_pending_amount = round(pending_raw.sum(), 2)
        pending_deduction = pending_raw.combine(settlement_amount_col.clip(lower=0.0), min).sum()
    else:
        settlement_pending_amount = 0.0
        pending_deduction = 0.0
    net_settlement = round(reco_df["settlement_amount"].sum() - pending_deduction, 2)
    return {
        "Total orders": int(len(reco_df)),
        "Gross order value": round(reco_df["total"].sum(), 2),
        "Receipt before deduction": round(reco_df["receipt_amount"].sum(), 2),
        "Total deduction": round(reco_df["total_deduction"].sum(), 2),
        # Settlement = Receipt - Deduction - Refund (see engine/reco.py), so
        # refund has to be its own visible line here too - otherwise Receipt
        # minus Deduction silently doesn't equal Net Settlement and it looks
        # like the report is missing money rather than having refunded it.
        "Total refund": round(reco_df["refund_amount"].sum(), 2),
        # Money already claimed (settlement_amount) but not yet bank-credited -
        # see attach_settlement_pending() above. Shown as its own line so
        # Net settlement's subtraction is visible/auditable, not silent.
        "Settlement pending": settlement_pending_amount,
        "Net settlement": net_settlement,
        "Total diff (unreconciled)": round(reco_df["diff"].sum(), 2),
        "Open queries (orders)": int((reco_df["query"] != "Okk").sum()),
    }
