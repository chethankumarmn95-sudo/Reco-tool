"""
amazon_reco.py
--------------
Layer 3 of the Amazon pipeline (parallel to engine/reco.py for Shopify) -
joins the MTR revenue ledger (engine.amazon_loaders.load_mtr_reports) with
the Settlement Flat File's order-level expense ledger
(engine.amazon_consolidator.summarize_ledger_by_order) into one order-level
reconciliation table, and rolls the settlement-level (non-order) items and
settlement-to-bank position into the top-line waterfall.

Waterfall replicated (per the prior year's manual workbook, "Summary"
sheet, and confirmed with the client):
    Sales as per MTR (Invoice Value - see note below)
    less Refunds
    = Net Sales
    less Order-level Deductions (flat file, order-id-attributable)
    = Order-level Payable
    less Settlement-level Deductions (flat file, non-order-attributable:
        Advertising, MCF, Storage/Warehouse, Debt adj - Reserve movements
        excluded here, see note)
    = Receivable
    less Received to date (bank statement, via engine/amazon_bank.py)
    = Balance Receivable

Note on Invoice Value vs Principal Amount: this engine uses MTR's Invoice
Amount (Principal + Shipping + GST) as the revenue base throughout, per the
prior year's Audit Notes & Glossary Section B - Amazon computes commission/
closing fee on Invoice Value, and it's the GST-invoice, legally recognised
revenue figure. Principal Amount alone would understate revenue and
overstate deduction percentages.

Note on Reserve Movements: Current Reserve Amount (withheld) and Previous
Reserve Amount Balance (released) are cash-flow timing, not a real cost -
excluded from the Settlement-level Deductions figure used in the waterfall
above (matching the prior workbook's treatment: "NOT a deduction... no P&L
impact"), but still fully retained and visible in the expense ledger/
settlement-level detail for anyone who wants to trace them.

Note on MCF (Multi-Channel Fulfilment) references: some flat file rows
carry an order-id-shaped reference that isn't a real Amazon marketplace
order (e.g. "S02-xxxxxxx-xxxxxxx") - Amazon fulfilling a Shopify/D2C order
and collecting COD cash on its behalf. These never appear in MTR at all
(MTR only lists real marketplace orders), and the prior year's own
workbook explicitly booked this money under the Shopify P&L, not Amazon's.
build_order_reconciliation() below therefore only joins flat-file order-
level deductions/credits for order-ids that DO exist in MTR; anything left
over is surfaced separately by non_mtr_order_items() so it's still visible
(and still included in the settlement-to-bank tie-out, which works off the
full ledger regardless of MTR matching) rather than silently dropped.

Note on the reporting cut-off (confirmed with the client, 2026-08-19): the
MTR report and the Settlement Flat File don't necessarily cover the same
window - the client's MTR is bounded by financial-year-end (31-Mar), but
the Settlement Flat File is an export of whatever date range the client
happened to download, which can run weeks past that (their FY25-26 export
ran to early May). The client's own rule, given directly: "MTR cut-off
determines the order population. Settlement files can extend beyond the
MTR period, but settlements after the MTR cut-off must be classified as
subsequent settlements against the March receivable and must not be
included in the receivable calculation as of 31-Mar-2026. Bank receipts
must also be split by receipt date." classify_settlements_by_cutoff() and
subsequent_settlements_summary() below implement the settlement side of
that rule; the bank-receipt side is in engine/amazon_bank.py's
split_received_by_cutoff(). None of this is a "new methodology" invented
here - it's the client's own rule, made generic/configurable (auto-derived
default cut-off from the MTR file's own last date, but overridable) so it
applies correctly to every future period, not just this one.
"""

import pandas as pd

from .amazon_consolidator import (
    summarize_ledger_by_order, summarize_settlement_level_adjustments,
    non_mtr_order_level_items, reserve_movement_summary, _CASH_DEDUCTION_BUCKETS,
    order_ids_with_settlement_row,
)


def infer_cutoff_date(mtr_df):
    """
    Auto-derives the reporting cut-off from the MTR file's own coverage -
    "MTR cut-off determines the order population" (the client's own rule,
    see module docstring). Uses the later of MTR's own max order_date/
    invoice_date, at end-of-day, so a settlement deposited ON the cut-off
    date itself is still counted as "up to" the cut-off, not after it.
    Returns None if mtr_df has no usable dates (caller should fall back to
    treating everything as "within cut-off", i.e. today's pre-cutoff-aware
    behavior, rather than crash).
    """
    if mtr_df is None or mtr_df.empty:
        return None
    candidates = []
    for col in ("invoice_date", "order_date"):
        if col in mtr_df.columns:
            m = pd.to_datetime(mtr_df[col], errors="coerce").max()
            if pd.notna(m):
                candidates.append(m)
    if not candidates:
        return None
    cutoff = max(candidates).normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
    return cutoff


def classify_settlements_by_cutoff(settlement_summary_df, cutoff_date, date_col="deposit_date"):
    """
    Splits settlements into "within" (up to and including cutoff_date) and
    "subsequent" (after cutoff_date) settlement_id sets, per the client's
    own rule (see module docstring) - a settlement's classification is
    all-or-nothing based on its OWN date (deposit_date by default: when
    Amazon actually finalizes/pays out that settlement), not a per-line
    split. date_col is overridable in case a client's own convention keys
    off "end_date" (when the settlement's data window closes) instead.

    Returns (within_ids, subsequent_ids) - both sets of settlement_id
    strings, matching expense_ledger_df's settlement_id dtype (str) so
    callers can filter the ledger with `.isin(...)` directly.
    """
    if settlement_summary_df is None or settlement_summary_df.empty or cutoff_date is None:
        all_ids = set(settlement_summary_df["settlement_id"].astype(str)) if (
            settlement_summary_df is not None and not settlement_summary_df.empty) else set()
        return all_ids, set()

    dates = pd.to_datetime(settlement_summary_df[date_col], errors="coerce")
    ids = settlement_summary_df["settlement_id"].astype(str)
    within_ids = set(ids[dates <= cutoff_date])
    subsequent_ids = set(ids[dates > cutoff_date])
    # A settlement with an unparseable date is treated as "within" rather
    # than silently vanishing from the receivable - conservative default,
    # same "never silently drop a rupee" principle as non_mtr_order_items.
    subsequent_ids -= within_ids
    unclassified = set(ids) - within_ids - subsequent_ids
    within_ids |= unclassified
    return within_ids, subsequent_ids


def subsequent_settlements_summary(settlement_summary_df, expense_ledger_df, subsequent_ids):
    """
    One row per settlement classified as "subsequent" (after the reporting
    cut-off) - settlement_id | payment_mode | start_date | end_date |
    deposit_date | settlement_amount | order_level_deductions |
    settlement_level_deductions - so a subsequent settlement isn't just
    excluded from the March receivable, it's fully visible and traceable
    as "money/deductions that will land against the March receivable in a
    later period" (the client's own framing).
    """
    cols = ["settlement_id", "payment_mode", "start_date", "end_date", "deposit_date",
            "settlement_amount", "order_level_deductions", "settlement_level_deductions"]
    if not subsequent_ids or settlement_summary_df is None or settlement_summary_df.empty:
        return pd.DataFrame(columns=cols)

    df = settlement_summary_df.copy()
    df["settlement_id"] = df["settlement_id"].astype(str)
    df = df[df["settlement_id"].isin(subsequent_ids)].rename(columns={"total_amount": "settlement_amount"})

    if expense_ledger_df is not None and not expense_ledger_df.empty:
        ledger = expense_ledger_df[
            expense_ledger_df["settlement_id"].astype(str).isin(subsequent_ids)
            & expense_ledger_df["bucket"].isin(_CASH_DEDUCTION_BUCKETS)
        ]
        order_ded = ledger[ledger["level"] == "Order"].groupby("settlement_id")["amount"].sum()
        settlement_ded = ledger[ledger["level"] == "Settlement"].groupby("settlement_id")["amount"].sum()
    else:
        order_ded = pd.Series(dtype=float)
        settlement_ded = pd.Series(dtype=float)

    df["order_level_deductions"] = df["settlement_id"].map(order_ded).fillna(0.0).round(2)
    df["settlement_level_deductions"] = df["settlement_id"].map(settlement_ded).fillna(0.0).round(2)
    return df[cols].sort_values("deposit_date").reset_index(drop=True)


def build_mtr_order_summary(mtr_df):
    """
    Collapses the MTR ledger (one row per order-item / refund-item / cancel-
    item line) to one row per order_id:
        order_id | segment | order_date | invoice_date | quantity |
        invoice_amount | taxable | tax_amount | tcs_amount | payment_type |
        skus | mtr_status

    Summing Invoice Amount/Taxable/Tax Amount across every MTR transaction
    type for an order is safe without filtering by type first: Shipment
    rows carry the real positive revenue, Refund rows already carry the
    matching NEGATIVE revenue (Amazon's own MTR export convention -
    verified against the client's real MTR file), and Cancel/
    FreeReplacement rows are always zero (nothing was ever shipped/billed).
    So the sum is exactly "net sales as per MTR" with no special-casing.
    """
    cols = [
        "order_id", "segment", "order_date", "invoice_date", "quantity", "invoice_amount",
        "taxable", "tax_amount", "tcs_amount", "payment_type", "skus", "mtr_status",
        "has_refund", "refund_amount",
    ]
    if mtr_df is None or mtr_df.empty:
        return pd.DataFrame(columns=cols)

    df = mtr_df.copy()

    def _join_skus(s):
        return ", ".join(sorted({str(v) for v in s if pd.notna(v) and str(v).strip()}))

    def _status_by_latest_fallback(g):
        """
        Old severity-based guess, used ONLY for the (hopefully rare) orders
        that have no usable date on ANY of invoice_date/order_date/
        shipment_date for ANY of their lines - a genuinely incomplete/
        malformed export. Kept as a per-group Python function since this
        path is expected to run on a tiny minority of orders at most (see
        _status_by_order()'s docstring below for why the normal path
        doesn't need this).
        """
        types = set(str(v).strip().lower() for v in g["mtr_transaction_type"])
        if "refund" in types:
            return "Refund"
        if "shipment" in types:
            return "Delivered"
        if "cancel" in types and len(types) == 1:
            return "Cancelled"
        if "freereplacement" in types:
            return "Free Replacement"
        return "Other"

    def _status_by_order(df):
        """
        Vectorized replacement for a per-order `groupby().apply()` call -
        mtr_status reflects whichever transaction line is CHRONOLOGICALLY
        LATEST for this order, not a fixed severity ranking over the set
        of types present - per the client's own rule: the same order can
        legitimately show "Shipment" in one MTR report and "Refund" in a
        later one (payment made, then the order was undelivered/RTO'd or
        cancelled by the customer), or the reverse (a Cancelled line later
        superseded by a real Shipment), and the ORDER-LEVEL status must
        reflect whatever actually happened most recently - it must not
        get stuck showing "Refund" forever just because a fixed priority
        ranking always favoured it over "Delivered" regardless of which
        one actually happened last. Every individual line itself is still
        kept in full in the ledger (see engine/amazon_consolidator.py) -
        this only decides the one-line-per-order summary status.

        Uses each line's own Invoice Date as "when this transaction
        actually happened" (falling back to Order Date, then Shipment
        Date, if Invoice Date is blank for every line of THAT ORDER) -
        picking the type of whichever row has the latest of those.

        Performance note: an MTR export can run into the lakhs of lines for
        a full financial year, and this used to call a Python function once
        per DISTINCT ORDER via `.groupby("order_id").apply(...)` - for a
        seller with hundreds of thousands of orders/year, that's hundreds
        of thousands of Python-level function calls, each constructing a
        small sub-DataFrame, purely for pandas dispatch overhead. Rewritten
        here as three `.groupby().transform("count")` passes (vectorized,
        C-level) to decide, per order, which single date column to rank on,
        then one `.idxmax()` per order via groupby (also vectorized) to
        pick the winning row - no per-order Python call for the vast
        majority of orders. Only orders with literally no usable date on
        ANY line (a malformed export - expected to be rare to nonexistent
        in a real file) fall back to the old per-group severity guess, and
        that fallback only ever runs against that small leftover subset.
        """
        has_invoice = df.groupby("order_id")["invoice_date"].transform("count") > 0
        has_order_date = df.groupby("order_id")["order_date"].transform("count") > 0
        rank_date = df["invoice_date"].where(
            has_invoice, df["order_date"].where(has_order_date, df["shipment_date"])
        )
        return rank_date

    rank_date = _status_by_order(df)
    valid_mask = rank_date.notna()
    order_id_valid = df.loc[valid_mask, "order_id"]
    idxmax_per_order = (
        rank_date[valid_mask].groupby(order_id_valid).idxmax()
        if valid_mask.any() else pd.Series(dtype="int64")
    )
    # Vectorized dict-lookup instead of .map(_classify_type) (a per-value
    # Python function call) - a plain string->string dict passed to
    # Series.map() is a fast, fully vectorized categorical lookup.
    _classify_map = {
        "refund": "Refund", "shipment": "Delivered", "cancel": "Cancelled",
        "freereplacement": "Free Replacement", "free replacement": "Free Replacement",
    }
    winning_types = df.loc[idxmax_per_order.to_numpy(), "mtr_transaction_type"].astype(str).str.strip().str.lower()
    status_valid = winning_types.map(_classify_map).fillna("Other")
    status_valid.index = idxmax_per_order.index  # index = order_id
    status_valid = status_valid.rename("mtr_status")

    orders_without_any_date = set(df["order_id"]) - set(idxmax_per_order.index)
    if orders_without_any_date:
        fallback_df = df[df["order_id"].isin(orders_without_any_date)]
        status_fallback = fallback_df.groupby("order_id", group_keys=False).apply(
            _status_by_latest_fallback, include_groups=False
        ).rename("mtr_status")
        status_by_order = pd.concat([status_valid, status_fallback])
    else:
        status_by_order = status_valid

    agg = df.groupby("order_id").agg(
        segment=("segment", "first"),
        order_date=("order_date", "min"),
        invoice_date=("invoice_date", "min"),
        quantity=("quantity", "sum"),
        invoice_amount=("invoice_amount", "sum"),
        taxable=("taxable", "sum"),
        tax_amount=("tax_amount", "sum"),
        tcs_amount=("tcs_amount", "sum"),
        payment_type=("payment_type", "first"),
        skus=("sku", _join_skus),
    ).reset_index()

    agg = agg.merge(status_by_order, on="order_id", how="left")

    refund_rows = df[df["mtr_transaction_type"].str.lower() == "refund"]
    refund_by_order = refund_rows.groupby("order_id")["invoice_amount"].sum().rename("refund_amount")
    agg = agg.merge(refund_by_order, on="order_id", how="left")
    agg["refund_amount"] = agg["refund_amount"].fillna(0.0)
    agg["has_refund"] = agg["refund_amount"] != 0

    return agg[cols]


def build_order_reconciliation(mtr_df, expense_ledger_df, full_expense_ledger_df=None):
    """
    The order-level reconciliation table - one row per Amazon order_id,
    joining MTR revenue with the flat file's order-level deduction/credit
    totals (engine.amazon_consolidator.summarize_ledger_by_order):
        order_id | segment | order_date | skus | quantity | mtr_status |
        invoice_amount | refund_amount | net_sales | order_deductions |
        order_credits | net_payout | has_settlement_row

    "has_settlement_row" is False for an order that's in MTR (so it was
    genuinely shipped/invoiced) but hasn't appeared in the flat file AT
    ALL yet - i.e. genuinely settlement pending, the Amazon-side
    equivalent of the Shopify pipeline's "Delivered but not yet in gateway
    settlement" case.

    full_expense_ledger_df (optional): pass the FULL, unfiltered ledger
    here (before any reporting-cut-off filtering - see
    classify_settlements_by_cutoff()/infer_cutoff_date() above) when the
    caller applies one. has_settlement_row is computed from this full
    ledger, via order_ids_with_settlement_row() (ANY order-level row, any
    bucket - not just the cash-deduction buckets order_deductions/
    order_credits are restricted to), so it always means "has Amazon
    settled this order at ALL, ever" rather than "has Amazon settled this
    order within the current cut-off window" - a subsequent settlement (a
    real settlement, just excluded from THIS period's receivable per the
    client's cut-off rule) should never make an order look like it's
    stuck in limbo with no settlement at all. order_deductions/
    order_credits themselves still come from `expense_ledger_df` (the
    cut-off-scoped one, if the caller passed a filtered one) - only the
    flag's SOURCE set is broadened, not the receivable-relevant totals.
    Falls back to `expense_ledger_df` itself when not supplied (the
    original, pre-cut-off-feature behaviour, and the correct default when
    the caller never filtered anything to begin with).
    """
    cols = [
        "order_id", "segment", "order_date", "invoice_date", "skus", "quantity", "mtr_status",
        "payment_type", "invoice_amount", "refund_amount", "net_sales", "order_deductions",
        "order_credits", "net_payout", "has_settlement_row",
    ]
    mtr_summary = build_mtr_order_summary(mtr_df)
    if mtr_summary.empty:
        return pd.DataFrame(columns=cols)

    known_order_ids = set(mtr_summary["order_id"])
    order_ded = summarize_ledger_by_order(expense_ledger_df, known_order_ids=known_order_ids)
    settled_order_ids = order_ids_with_settlement_row(
        full_expense_ledger_df if full_expense_ledger_df is not None else expense_ledger_df
    )

    df = mtr_summary.merge(order_ded, on="order_id", how="left")
    df["order_deductions"] = df["order_deductions"].fillna(0.0)
    df["order_credits"] = df["order_credits"].fillna(0.0)
    df["has_settlement_row"] = df["order_id"].astype(str).isin(settled_order_ids)

    df["net_sales"] = df["invoice_amount"]  # refund_amount is already netted into invoice_amount (see build_mtr_order_summary)
    df["net_payout"] = df["net_sales"] + df["order_deductions"] + df["order_credits"]

    return df[cols]


def settlement_level_summary(expense_ledger_df):
    """
    Settlement-level (non-order) adjustments, rolled up by category across
    ALL settlements - the Amazon equivalent of the prior workbook's
    "11. Monthly S2 Analysis" sheet, minus the Reserve Movement bucket
    (see module docstring - excluded from the deduction total, kept
    visible separately).
        category | bucket | amount
    """
    cols = ["category", "bucket", "amount"]
    detail = summarize_settlement_level_adjustments(expense_ledger_df)
    if detail.empty:
        return pd.DataFrame(columns=cols)
    grouped = detail.groupby(["category", "bucket"])["amount"].sum().reset_index()
    return grouped[cols]


def build_waterfall(order_reco_df, expense_ledger_df, mtr_df=None, received_to_date=0.0,
                     received_in_transit=0.0, subsequent_settlements_total=None):
    """
    The single top-line waterfall (Amazon equivalent of the prior
    workbook's "Summary" sheet section B), as a list of
    {"Particular": ..., "Amount": ...} rows ready to display/export in
    order.

    IMPORTANT: expense_ledger_df here should already be scoped to
    settlements "within" the reporting cut-off (see
    classify_settlements_by_cutoff() above) - this function itself does
    no date filtering, it just totals whatever ledger it's handed. The
    caller (views/page_reconciliation.py) is responsible for the split,
    per the client's own rule (see module docstring): settlements after
    the MTR cut-off must not be included in the receivable calculation.

    received_to_date should likewise already be restricted to bank
    receipts dated on/before the cut-off (see
    engine/amazon_bank.py's split_received_by_cutoff) - "Bank receipts
    must also be split by receipt date," independently of which
    settlement they belong to.

    received_in_transit (optional): bank receipts for "within cut-off"
    settlements whose ACTUAL bank credit landed after the cut-off (a
    timing lag between Amazon's stated deposit-date and the real bank
    credit) - shown as a memo, since it's not yet cash-received as of the
    cut-off, but the corresponding deduction IS already recognized above
    (the settlement itself is "within cut-off"), so Balance Receivable
    already correctly includes it as still-outstanding.

    subsequent_settlements_total (optional): the total settlement amount
    of settlements classified as "subsequent" (after cut-off) - shown as
    a memo only, per the client's rule that these must NOT be included in
    the receivable calculation as of the cut-off; see
    subsequent_settlements_summary() above for the full traceable detail.

    Two more memo lines are appended after the main waterfall - Reserve
    Movement and non-MTR (MCF/Shopify pass-through) order items - neither
    is part of the Amazon P&L total above, but both are shown so nothing
    this engine parsed from the flat file is ever invisible; see this
    module's docstring for why each is excluded from the main total.
    """
    if order_reco_df is None or order_reco_df.empty:
        sales, refunds, order_ded = 0.0, 0.0, 0.0
    else:
        sales = float(order_reco_df["invoice_amount"].sum() - order_reco_df["refund_amount"].sum())
        refunds = float(order_reco_df["refund_amount"].sum())
        order_ded = float((order_reco_df["order_deductions"] + order_reco_df["order_credits"]).sum())

    settlement_detail = summarize_settlement_level_adjustments(expense_ledger_df)
    settlement_ded = float(settlement_detail["amount"].sum()) if not settlement_detail.empty else 0.0
    reserve_net = reserve_movement_summary(expense_ledger_df)

    known_order_ids = set(mtr_df["order_id"]) if mtr_df is not None and not mtr_df.empty else set(
        order_reco_df["order_id"]) if order_reco_df is not None and not order_reco_df.empty else set()
    non_mtr = non_mtr_order_level_items(expense_ledger_df, known_order_ids)
    non_mtr_total = float(non_mtr["amount"].sum()) if not non_mtr.empty else 0.0

    net_sales = sales + refunds
    order_level_payable = net_sales + order_ded
    receivable = order_level_payable + settlement_ded
    balance_receivable = receivable - received_to_date

    rows = [
        {"Particular": "Sales as per MTR (Invoice Value)", "Amount": round(sales, 2)},
        {"Particular": "Less: Refunds", "Amount": round(refunds, 2)},
        {"Particular": "Net Sales", "Amount": round(net_sales, 2)},
        {"Particular": "Less: Order-level Deductions (Flat File, incl. TDS/TCS)", "Amount": round(order_ded, 2)},
        {"Particular": "Order-level Payable", "Amount": round(order_level_payable, 2)},
        {"Particular": "Less: Settlement-level Deductions (Flat File)", "Amount": round(settlement_ded, 2)},
        {"Particular": "Receivable", "Amount": round(receivable, 2)},
        {"Particular": "Less: Received to date (Bank Statement)", "Amount": round(-abs(received_to_date), 2)},
        {"Particular": "Balance Receivable", "Amount": round(balance_receivable, 2)},
        {"Particular": "(Memo) Net Reserve Movement - cash-flow timing, not in the total above", "Amount": round(reserve_net, 2)},
        {"Particular": "(Memo) Non-MTR order-level items (MCF/Shopify pass-through) - not in the total above", "Amount": round(non_mtr_total, 2)},
    ]
    if received_in_transit:
        rows.append({
            "Particular": "(Memo) Receipts in transit - settlement within cut-off, bank credit landed after it",
            "Amount": round(received_in_transit, 2),
        })
    if subsequent_settlements_total is not None:
        rows.append({
            "Particular": "(Memo) Subsequent settlements (after cut-off) - against this receivable, not yet due",
            "Amount": round(subsequent_settlements_total, 2),
        })
    return pd.DataFrame(rows)
