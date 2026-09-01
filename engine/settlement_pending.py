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

Note on scope, consistent with this whole engine's disclosed-heuristic
approach (see engine/bank.py's module docstring and engine/settlement.py's
expected_gateway_for_order()): "Expected Settlement Date" here is a
best-effort estimate (order date + that gateway's configured
expected_settlement_days, or a generic fallback if not configured) for
planning/ageing purposes only - it is not a claim about any gateway's
actual contractual SLA. Confirm the real settlement terms with each
gateway/courier directly before relying on this for anything beyond
prioritising follow-ups.
"""

import pandas as pd

from .bank import (
    COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW,
)
from .settlement import expected_gateway_for_order
from .reco import _COD_SETTLEMENT_PENDING_PHRASES

PENDING_CATEGORIES = {COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW}

DETAIL_COLUMNS = [
    "Order ID", "Payment Gateway", "Payment/Transaction Reference (UTR)",
    "Order Date", "Payment/Receipt Date", "Order Amount", "Gateway Amount",
    "Settlement Amount", "Expected Settlement Date", "Actual Settlement Date",
    "Bank Credit Date", "Settlement Status", "Days Pending",
]

# Client-reported 2026-08-30 (item 8): matches the client's own reference
# workbook's "Settlement Pending Summary" sheet exactly - 4 columns, no
# separate "Exceptions"/"Exception Amount" breakout (Orders Pending/Amount
# Pending already includes every still-outstanding order, exceptions
# included - see settlement_pending_summary_by_gateway()'s own docstring).
SUMMARY_COLUMNS = ["Group", "Payment Gateway", "Orders Pending", "Amount Pending"]

DEFAULT_EXPECTED_DAYS = {"COD": 20, "Prepaid": 10, "Unknown": 10}


def build_settlement_pending_report(reco_df, recon_status_df, receipt_detail_df, gateway_configs, receipt_summary_df=None):
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
    """
    if reco_df is None or reco_df.empty or recon_status_df is None or recon_status_df.empty:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    df = recon_status_df[recon_status_df["Reconciliation Category"].isin(PENDING_CATEGORIES)].copy()
    if df.empty:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    reco_cols = ["order_id", "created_at", "total", "payment_method"]
    if "delivery_partner" in reco_df.columns:
        reco_cols.append("delivery_partner")
    reco = reco_df[reco_cols].copy()
    reco["order_id"] = reco["order_id"].astype(str)
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
    # Orders with no settlement row yet (has_settlement_row == False) have
    # receipt_amount == 0 - there's nothing yet for the gateway/courier to
    # have deducted anything FROM, so a receipt-based net figure would
    # wrongly show these orders as ₹0 pending instead of the full order
    # value that's genuinely still outstanding. Use the order's own Total
    # for those; only orders that DO have a settlement row get the net
    # (Gateway Amount minus that gateway's own deduction, minus refund)
    # figure, same "final payment" concept used elsewhere in this engine.
    _net_settlement = (df["receipt_amount"] - df["total_deduction"] - df["refund_amount"]).round(2)
    df["_settlement_amount"] = df["total"].where(~df["has_settlement_row"], _net_settlement)

    # No settlement row yet -> no gateway name to read off the receipt
    # ledger. Attribute it the same best-effort way engine/settlement.py
    # already does for "Pending for Settlement" (Payment Method text, or
    # the courier that delivered it for COD), rather than inventing a
    # second, different guess here - one methodology, disclosed once.
    def _attributed_gateway(row):
        if pd.notna(row.get("payment_gateway")) and str(row["payment_gateway"]).strip():
            return row["payment_gateway"]
        return expected_gateway_for_order(row.get("payment_method"), row.get("delivery_partner"), gateway_configs or [])

    df["Payment Gateway"] = df.apply(_attributed_gateway, axis=1)

    gateway_days = {cfg["label"]: cfg.get("expected_settlement_days") for cfg in (gateway_configs or [])}

    def expected_settlement_date(row):
        base = pd.to_datetime(row["created_at"], errors="coerce")
        if pd.isna(base):
            return None
        days = gateway_days.get(row["Payment Gateway"])
        if days is None:
            days = DEFAULT_EXPECTED_DAYS.get(row["payment_type"], 10)
        return base + pd.Timedelta(days=days)

    df["_expected_settlement_date"] = df.apply(expected_settlement_date, axis=1)

    out = pd.DataFrame({
        "Order ID": df["order_id"],
        "Payment Gateway": df["Payment Gateway"],
        "Payment/Transaction Reference (UTR)": df["utr"],
        "Order Date": df["created_at"],
        "Payment/Receipt Date": df["receipt_date"],
        "Order Amount": df["total"],
        "Gateway Amount": df["receipt_amount"],
        "Settlement Amount": df["_settlement_amount"],
        "Expected Settlement Date": df["_expected_settlement_date"],
        "Actual Settlement Date": df["receipt_date"],
        "Bank Credit Date": df["bank_credit_date"],
        "Settlement Status": df["Reconciliation Category"],
        "Days Pending": df["days_pending"],
    })
    return out.sort_values("Days Pending", ascending=False, na_position="last").reset_index(drop=True)


def _gateway_group(label, gateway_configs):
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
    """
    label = str(label or "").strip()
    return label.capitalize() if label and label.islower() else label


def settlement_pending_summary_by_gateway(reco_df, gateway_configs):
    """
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
      - Grouped by the SAME resolved "Gateway" column already on Reco
        working (falling back to "Payment Provider" if Gateway isn't
        present) - one canonical attribution shared with the Reco working
        sheet's own Query text, rather than a second, independent
        attribution mechanism that could (and did) disagree with it.
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

    df = reco_df.copy()
    pending = df["receipt_status"].fillna("") == "Not Received"
    if not pending.any():
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    df = df.loc[pending].copy()

    if "Gateway" in df.columns:
        gateway_raw = df["Gateway"]
    elif "Payment Provider" in df.columns:
        gateway_raw = df["Payment Provider"]
    else:
        gateway_raw = pd.Series(None, index=df.index)
    gateway_text = gateway_raw.astype(str).str.strip()
    unresolved = gateway_raw.isna() | gateway_text.isin(["", "nan", "None", "#N/A", "NA"])

    query_col = df["query"] if "query" in df.columns else pd.Series("", index=df.index)
    is_catch_all_query = query_col.fillna("") == "COD Delivered Amount not Received"

    groups, labels, amounts = [], [], []
    for idx in df.index:
        if unresolved.loc[idx]:
            groups.append("COD")
            labels.append("COD Delivered Amount not Received")
            amounts.append(df.at[idx, "total"])
            continue
        raw_label = gateway_raw.loc[idx]
        grp = _gateway_group(raw_label, gateway_configs)
        labels.append(_display_gateway_label(raw_label))
        groups.append(grp)
        if grp == "COD" and not is_catch_all_query.loc[idx]:
            amounts.append(df.at[idx, "receipt_amount"])
        else:
            amounts.append(df.at[idx, "total"])

    df["Group"] = groups
    df["Payment Gateway"] = labels
    df["_amount"] = amounts
    df["_amount"] = df["_amount"].fillna(0.0)

    grouped = df.groupby(["Group", "Payment Gateway"]).agg(**{
        "Orders Pending": ("order_id", "count"),
        "Amount Pending": ("_amount", "sum"),
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
