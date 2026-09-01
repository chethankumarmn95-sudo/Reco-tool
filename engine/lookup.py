"""
lookup.py
---------
Builds the "Order Lookup Dashboard" table - one row per order with every
field a client might ask about, pulled together from all the layers of
the pipeline. This is deliberately separate from reco.py's financial
Reco working table: that one is optimized for totals, this one is
optimized for "tell me everything about order #12345".
"""

import pandas as pd
from .loaders import normalize_order_id, resolve_col_or_raise


def build_sku_detail(source_frames, sku_cfg):
    """
    Pulls SKU names and quantity per order from the configured SKU source
    (Unicommerce for ESCA - it's line-item level, so each order's rows are
    grouped back together here).
    """
    label = sku_cfg["label"]
    if label not in source_frames:
        return pd.DataFrame(columns=["order_id", "skus", "quantity"])

    df = source_frames[label].copy()
    order_id_col = resolve_col_or_raise(df, sku_cfg["order_id_col"], label)
    df["order_id"] = normalize_order_id(df[order_id_col])
    sku_col = sku_cfg["sku_col"]

    grouped = df.groupby("order_id").agg(
        skus=(sku_col, lambda s: ", ".join(sorted(set(str(v) for v in s if pd.notna(v))))),
        quantity=(sku_col, "count"),  # one row per unit in Unicommerce's export
    ).reset_index()

    return grouped


def reconciliation_status(query):
    return "Reconciled" if query == "Okk" else query


def refund_status(row):
    """
    Only flags "Refund Pending" when money was actually received and not
    yet refunded - not for every RTO/Cancelled order. A COD order that
    never got collected has nothing to refund; that's normal, not pending.
    """
    if row["refund_amount"] > 1:
        return f"Refunded (₹{row['refund_amount']:,.2f})"
    if row["final_delivery_status"] in ("RTO", "Cancelled", "Lost", "Refunded"):
        if row["receipt_amount"] > 1:
            return "Refund Pending"
        return "N/A (nothing collected - COD not yet received)"
    return "N/A"


def build_exception_detail(reco_df, lookup_df, channel_name):
    """Merges reco_df (raw per-partner statuses, money) with lookup_df
    (payment/bank/refund detail) into one order-level table with every
    column the exception report asks for - Order ID, delivery partner,
    payment/bank status, exception type/reason, and reconciliation status,
    enough to trace an exception back to its source. Shared by the
    Exceptions page (interactive, filterable) and the Reports page's
    downloaded workbook (a static "Exceptions" sheet), so both always agree
    on exactly what counts as an exception."""
    df = reco_df.copy()
    df["order_id"] = df["order_id"].astype(str)

    if lookup_df is not None:
        lk = lookup_df.copy()
        lk["Order ID"] = lk["Order ID"].astype(str)
        merge_cols = ["Order ID", "Payment method", "Payment gateway", "Bank credit date",
                      "Bank status", "Reconciliation status", "Pending amount", "Refund status"]
        # Reconciliation Category / Reason (engine.bank.classify_order_bank_status) -
        # only present once a run has produced them; older saved months won't have
        # these columns, so they're included opportunistically rather than required.
        merge_cols += [c for c in ("Reconciliation Category", "Category Reason") if c in lk.columns]
        df = df.merge(lk[merge_cols], left_on="order_id", right_on="Order ID", how="left")

    df["Sales Channel"] = channel_name or ""
    df["Payment/Gateway Status"] = df["receipt_amount"].apply(lambda x: "Received" if x > 1 else "Not Received")
    df["Exception Type"] = df["final_delivery_status"]
    df["Exception Reason"] = df["query"]

    unicommerce_col = "unicommerce_raw_status" if "unicommerce_raw_status" in df.columns else None

    out = pd.DataFrame({
        "Order ID": df["order_id"],
        "Order Date": df["created_at"],
        "Sales Channel": df["Sales Channel"],
        "Delivery Partner": df["delivery_partner"],
        "Unicommerce Status": df[unicommerce_col] if unicommerce_col else None,
        "Delivery Status": df["final_delivery_status"],
        "Payment Method": df.get("Payment method"),
        "Payment/Gateway Status": df["Payment/Gateway Status"],
        "Bank/UTR Status": df.get("Bank status"),
        "Bank Credit Date": df.get("Bank credit date"),
        "Reconciliation Category": df.get("Reconciliation Category"),
        "Category Reason": df.get("Category Reason"),
        "Exception Type": df["Exception Type"],
        "Exception Reason": df["Exception Reason"],
        "Order Value": df["total"],
        "Receipt Amount": df["receipt_amount"],
        "Pending Amount": df.get("Pending amount", df["diff"]),
        "Reconciliation Status": df.get("Reconciliation status", df["query"].apply(lambda q: "Reconciled" if q == "Okk" else q)),
    })
    return out


# Client-reported 2026-08-30 (item 10): the DOWNLOADED "Exceptions" sheet's
# own column set/order, reverse-engineered header-by-header against
# MOD.xlsx - matches Reco working's own internal column names verbatim
# where MOD does (order_id, created_at, delivery_partner, final_delivery_
# status, subtotal, shipping, taxes, total, receipt_amount, total_deduction,
# refund_amount, diff, unicommerce_raw_status), and its own display renames
# elsewhere (Recipt Remark/Bank credit - same as Reco working's own export
# rename; Exception Reason - Query's rename here specifically). Deliberately
# NOT the richer lookup-merged table build_exception_detail() below returns
# (Bank/UTR Status, Reconciliation Category, Category Reason, ...) - that
# richer shape stays exactly as-is for the interactive Exceptions page's own
# filter widgets (Exception Type/Reconciliation Status columns it depends
# on), which is out of scope for matching the client's reference workbook.
EXCEPTION_EXPORT_COLUMNS = [
    "order_id", "created_at", "Sales Channel", "unicommerce_raw_status",
    "delivery_partner", "final_delivery_status", "Payment Method", "Payment Provider",
    "subtotal", "shipping", "taxes", "total", "Recipt Remark", "receipt_amount",
    "total_deduction", "refund_amount", "diff", "Bank credit", "Exception Reason",
]


def build_exception_export_view(reco_df, order_ids, channel_name):
    """
    Builds the DOWNLOADED "Exceptions" sheet exactly as the client's own
    reference workbook has it - see EXCEPTION_EXPORT_COLUMNS above for the
    full rationale. `order_ids` should be the same order_id set the
    interactive Exceptions page's own "Reconciliation Status != Reconciled"
    filter already selected (via build_exception_detail() below), so the
    interactive page and this exported sheet never disagree about WHICH
    orders count as exceptions - only the columns SHOWN differ.

    "Bank credit" applies the same item-14 zeroing views/page_reports.py's
    _build_workbook() already applies to the "Reco working" sheet's own
    "Bank credit" column (settlement_amount, zeroed for any order still
    settlement-pending) - this sheet would otherwise show a nonzero amount
    for money that, per that fix, hasn't actually reached the bank yet.
    """
    if reco_df is None or reco_df.empty:
        return pd.DataFrame(columns=EXCEPTION_EXPORT_COLUMNS)

    ids = set(str(i) for i in order_ids) if order_ids is not None else set()
    df = reco_df[reco_df["order_id"].astype(str).isin(ids)].copy()
    if df.empty:
        return pd.DataFrame(columns=EXCEPTION_EXPORT_COLUMNS)

    df["Sales Channel"] = channel_name or ""
    bank_credit = df["settlement_amount"] if "settlement_amount" in df.columns else 0.0
    if "settlement_pending_amount" in df.columns:
        still_pending = df["settlement_pending_amount"].fillna(0) > 0.01
        bank_credit = pd.Series(bank_credit, index=df.index).where(~still_pending, 0.0)
    df["Bank credit"] = bank_credit
    df["Exception Reason"] = df.get("query")
    df["Recipt Remark"] = df.get("receipt_status")

    for col in EXCEPTION_EXPORT_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[EXCEPTION_EXPORT_COLUMNS].reset_index(drop=True)


def build_order_lookup(reco_df, sku_detail_df, receipt_detail_df, recon_status_df=None):
    """
    Joins the Reco working table with SKU detail, receipt detail (gateway,
    mode, receipt date), and the per-order reconciliation-category
    classification (engine.bank.classify_order_bank_status) into the full
    Order Lookup table.

    recon_status_df should always be supplied now (page_reconciliation.py
    computes it unconditionally, whether or not a bank statement was
    uploaded this run) - it's what replaces the old blanket "Not checked -
    no bank statement uploaded" fallback, which used to fire for ANY order
    with no gateway receipt row at all - including COD orders that were
    never delivered, where no receipt was ever expected in the first place
    (see engine/bank.py's module-section docstring for the full story).
    Only a genuinely missing classification (e.g. a much older saved month
    from before this existed) falls back to a clearly-labelled "not
    available" state rather than silently mislabeling it as "not checked".
    """
    df = reco_df.merge(sku_detail_df, on="order_id", how="left")
    df = df.merge(receipt_detail_df, on="order_id", how="left")

    df["reconciliation_status"] = df["query"].apply(reconciliation_status)
    df["pending_amount"] = df["diff"].round(2)
    df["refund_status"] = df.apply(refund_status, axis=1)

    if recon_status_df is not None and not recon_status_df.empty:
        rs = recon_status_df[["order_id", "Reconciliation Category", "Category Reason",
                               "bank_credit_date", "days_pending"]].copy()
        rs["order_id"] = rs["order_id"].astype(str)
        df["order_id"] = df["order_id"].astype(str)
        df = df.merge(rs, on="order_id", how="left")
        df["bank_status"] = df["Reconciliation Category"].fillna(
            "Reconciliation category not available for this saved run - re-run reconciliation to populate it"
        )
    else:
        df["Reconciliation Category"] = None
        df["Category Reason"] = None
        df["bank_credit_date"] = None
        df["days_pending"] = None
        df["bank_status"] = "Reconciliation category not available for this saved run - re-run reconciliation to populate it"

    # Payment Method / Payment Provider / Gateway (engine.attribution.
    # attach_payment_columns / build_payment_gateway_lookups) and the Bank
    # UTR Detail columns (engine.bank.build_order_level_utr_detail) are
    # both wired into reco_df upstream (views/page_reconciliation.py,
    # views/page_reports.py) - client-reported 2026-08-27, and the earlier
    # "Gateway" column client request (2026-08-25). Included here
    # opportunistically (only if actually present on reco_df) so Order
    # Lookup shows the same columns as the Reco working export, without
    # requiring every caller to have run the newer pipeline steps.
    # Client-reported 2026-08-30 (item 11): "receipt_status" used to be
    # included here too (as "Receipt Status") - verified against MOD.xlsx's
    # own "Order Lookup" sheet directly, it has no such column (28 columns
    # total, ending at "Refund Date") - removed to match exactly. It still
    # lives on the "Reco working" sheet ("Recipt Remark") and the
    # "Exceptions" sheet ("Recipt Remark"), both of which DO carry it in
    # MOD's own reference workbook.
    optional_passthrough_cols = [
        "Gateway", "Payment Method", "Payment Provider",
        "Payment Date (Bank Date)", "Payment UTR", "Setlment Remarks",
        "Refund UTR", "Refund Date",
    ]
    display_cols = [
        "order_id", "created_at", "final_delivery_status", "delivery_partner",
        "delivered_date", "rto_date", "skus", "quantity", "total", "receipt_amount",
        "payment_mode", "payment_gateway", "receipt_date", "bank_credit_date",
        "bank_status", "Reconciliation Category", "Category Reason", "days_pending",
        "reconciliation_status", "pending_amount", "refund_status",
    ] + [c for c in optional_passthrough_cols if c in reco_df.columns]
    for c in display_cols:
        if c not in df.columns:
            df[c] = None

    return df[display_cols].rename(columns={
        "order_id": "Order ID",
        "created_at": "Order date",
        "final_delivery_status": "Delivery status",
        "delivery_partner": "Delivery partner",
        "delivered_date": "Delivery date",
        "rto_date": "RTO date",
        "skus": "SKU(s)",
        "quantity": "Quantity",
        "total": "Order value",
        "receipt_amount": "Receipt amount",
        "payment_mode": "Payment method",
        "payment_gateway": "Payment gateway",
        "receipt_date": "Receipt date",
        "bank_credit_date": "Bank credit date",
        "bank_status": "Bank status",
        "days_pending": "Days pending",
        "reconciliation_status": "Reconciliation status",
        "pending_amount": "Pending amount",
        "refund_status": "Refund status",
        "receipt_status": "Receipt Status",
    })
