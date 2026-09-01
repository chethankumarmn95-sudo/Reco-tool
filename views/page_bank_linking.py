"""
page_bank_linking.py
---------------------
Shows how well gateway receipts have been traced through to actual bank
credits via UTR matching. Bank statement itself is uploaded on the Upload
Data page (it's just another source file); this page is where you review
the result of that linking.

Amazon/marketplace channel has no per-order UTR at all (see
engine/amazon_bank.py) - a settlement is one lump bank credit covering many
orders, matched to the bank statement by amount+deposit date instead, the
same way Shopify's COD courier batches are. _render_amazon below is that
equivalent view, reading amazon_settlement_register_df directly from
session_state (populated live by a Reconciliation run, or restored by
loading a saved period on Data Management) rather than from
theme.require_data_or_prompt()'s DTC-shaped autoload.
"""

import streamlit as st

from views.theme import require_data_or_prompt
from engine.bank import (
    COD_NOT_DELIVERED, COD_SETTLEMENT_PENDING, COD_BANK_MATCHED,
    PREPAID_SETTLEMENT_PENDING, PREPAID_BANK_MATCHED, EXCEPTION_MANUAL_REVIEW,
)
from engine.amazon_bank import settlement_register_summary


def _render_amazon():
    settlement_register_df = st.session_state.get("amazon_settlement_register_df")
    if settlement_register_df is None or settlement_register_df.empty:
        st.info(
            "No settlement-to-bank matching available yet. Go to **Upload Data** to upload the "
            "Settlement Flat File(s) and a bank statement, then **Reconciliation** to run it - or "
            "**Data Management** to reload a previously saved period."
        )
        return

    st.caption(
        "Amazon settlements have no per-order UTR - each settlement is one lump bank credit "
        "covering many orders, matched here to the bank statement by amount + deposit date "
        "(exactly like Shopify's COD courier remittance batches). See engine/amazon_bank.py."
    )

    counts = settlement_register_df["status"].value_counts()
    c1, c2, c3 = st.columns(3)
    c1.metric("Settlements matched", f"{counts.get('Matched', 0):,}")
    c2.metric("Nil / Negative (nothing due)", f"{counts.get('Nil / Negative', 0):,}")
    c3.metric("Unmatched (needs review)", f"{counts.get('Unmatched', 0):,}")

    st.divider()
    st.subheader("Settlement Summary (payment-mode-wise)")
    st.caption(
        "Mirrors the prior workbook's own \"Recon Summary\" section: Settlement Total vs Bank "
        "Receipts vs Variance, per payment mode (COD/Online) and overall."
    )
    st.dataframe(settlement_register_summary(settlement_register_df), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Settlements needing manual review")
    unmatched_df = settlement_register_df[settlement_register_df["status"] == "Unmatched"]
    if unmatched_df.empty:
        st.success("Every settlement with money due was successfully traced to a bank credit.")
    else:
        st.dataframe(unmatched_df, use_container_width=True, height=350)

    st.divider()
    st.subheader("Full settlement-by-settlement detail")
    st.dataframe(settlement_register_df, use_container_width=True, height=420)

    expense_ledger_df = st.session_state.get("amazon_expense_ledger_df")
    tie_out_df = st.session_state.get("amazon_tie_out_df")
    if tie_out_df is not None and not tie_out_df.empty:
        st.divider()
        st.subheader("Settlement parsing tie-out")
        st.caption(
            "Every line parsed from the Settlement Flat File, summed per settlement, checked "
            "against Amazon's own reported total for that settlement - the interim stand-in for "
            "matching expenses to Amazon's own fee/tax invoices until that report is available "
            "(see engine/amazon_invoice_check.py)."
        )
        tied = int((tie_out_df["status"] == "Tied out").sum())
        c1, c2 = st.columns(2)
        c1.metric("Settlements tied out", f"{tied:,} / {len(tie_out_df):,}")
        c2.metric("Needs review", f"{len(tie_out_df) - tied:,}")
        if tied < len(tie_out_df):
            st.dataframe(tie_out_df[tie_out_df["status"] != "Tied out"], use_container_width=True, hide_index=True)


def render():
    st.title("Bank Statement / UTR Linking")

    config = st.session_state.get("config") or {}
    if config.get("channel_type") == "marketplace":
        _render_amazon()
        return

    if not require_data_or_prompt("Upload Data"):
        return

    lookup_df = st.session_state.get("lookup_df")
    recon_status_df = st.session_state.get("recon_status_df")
    if lookup_df is None or "Reconciliation Category" not in lookup_df.columns:
        st.info("Run a reconciliation first to see bank linking status here.")
        return

    counts = lookup_df["Reconciliation Category"].value_counts()
    matched_count = counts.get(COD_BANK_MATCHED, 0) + counts.get(PREPAID_BANK_MATCHED, 0)
    pending_count = counts.get(COD_SETTLEMENT_PENDING, 0) + counts.get(PREPAID_SETTLEMENT_PENDING, 0)
    exception_count = counts.get(EXCEPTION_MANUAL_REVIEW, 0)
    not_delivered_count = counts.get(COD_NOT_DELIVERED, 0)

    st.caption(
        "Six reconciliation categories replace the old three-bucket \"Matched / Bank "
        "receipt not identified / Not checked\" view - see engine/bank.py's module-section "
        "docstring for exactly how each one is decided."
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Bank matched", f"{matched_count:,}")
    c2.metric("Settlement pending", f"{pending_count:,}")
    c3.metric("Exceptions (manual review)", f"{exception_count:,}")
    c4.metric("COD - not delivered (no receipt expected)", f"{not_delivered_count:,}")

    st.divider()
    st.subheader("Full category breakdown")
    st.dataframe(
        counts.rename_axis("Reconciliation Category").reset_index(name="Orders"),
        use_container_width=True, hide_index=True,
    )

    st.divider()
    st.subheader("Orders needing manual reconciliation")
    st.caption(
        "Delivered/settled orders that couldn't be traced to a bank credit within the "
        "expected window - by order UTR or by settlement-batch amount/date - and genuinely "
        "need a human to look, as distinct from orders that are simply still within their "
        "normal settlement/bank-credit timing."
    )
    exceptions_df = lookup_df[lookup_df["Reconciliation Category"] == EXCEPTION_MANUAL_REVIEW]
    if exceptions_df.empty:
        st.success("No orders are currently flagged as needing manual reconciliation.")
    else:
        cols = ["Order ID", "Payment gateway", "Order value", "Receipt amount",
                "Receipt date", "Days pending", "Category Reason"]
        st.dataframe(exceptions_df[[c for c in cols if c in exceptions_df.columns]],
                     use_container_width=True, height=400)

    cod_batches_df = st.session_state.get("cod_batches_df")
    cod_batch_match_df = st.session_state.get("cod_batch_match_df")
    if cod_batches_df is not None and not cod_batches_df.empty:
        st.divider()
        st.subheader("COD settlement batches (amount/date fallback match)")
        st.caption(
            "COD couriers remit collected cash as one lump bank credit covering many "
            "orders, and the per-order UTR on their settlement export is often blank or "
            "not the bank's own reference. This groups each courier's same-day settlement "
            "rows into a batch and matches the BATCH total to a bank credit by amount and "
            "date proximity, so a consolidated remittance can still be traced back to its "
            "orders even without a clean per-row UTR. Matched via amount/date, not a hard "
            "reference - confirm manually, same disclosed-heuristic approach as the rest of "
            "this page (see engine/bank.py's docstring)."
        )
        merged = cod_batches_df.merge(cod_batch_match_df, on="batch_id", how="left")
        merged["order_count"] = merged["order_ids"].apply(len)
        display_cols = ["batch_id", "source", "settlement_date", "order_count",
                         "batch_amount", "matched", "bank_amount", "bank_date", "match_note"]
        st.dataframe(merged[[c for c in display_cols if c in merged.columns]].rename(columns={
            "batch_id": "Batch", "source": "Gateway/Partner", "settlement_date": "Settlement date",
            "order_count": "Orders in batch", "batch_amount": "Batch amount", "matched": "Matched",
            "bank_amount": "Bank amount", "bank_date": "Bank date", "match_note": "Note",
        }), use_container_width=True, height=350)

    st.divider()
    st.subheader("Orders with unmatched bank receipts (order-level UTR view)")
    st.caption(
        "Gateway shows a receipt for these orders, but no matching UTR was found in the "
        "bank statement directly (before the settlement-batch fallback above)."
    )
    unmatched = lookup_df.iloc[0:0]
    if recon_status_df is not None and not recon_status_df.empty:
        unmatched_ids = set(recon_status_df.loc[
            (~recon_status_df["bank_matched"]) & recon_status_df["has_settlement_row"], "order_id"
        ].astype(str))
        unmatched = lookup_df[lookup_df["Order ID"].astype(str).isin(unmatched_ids)]
    if unmatched.empty:
        st.success("Every gateway receipt with a UTR was successfully traced to the bank statement (directly, or via a matched settlement batch).")
    else:
        cols = ["Order ID", "Payment gateway", "Receipt date", "Order value", "Reconciliation Category"]
        st.dataframe(unmatched[[c for c in cols if c in unmatched.columns]], use_container_width=True, height=400)

    utr_bank_reco_df = st.session_state.get("utr_bank_reco_df")
    if utr_bank_reco_df is not None and not utr_bank_reco_df.empty:
        st.divider()
        st.subheader("Bank Reconciliation by Settlement (UTR)")
        st.caption(
            "One row per bank credit (UTR) - not per order. A single settlement UTR very "
            "often bundles more than one order's payout together, and not always from the "
            "same reconciliation period (e.g. a delayed remittance for last month's orders "
            "landing in this month's bank credit). This splits each UTR's total into "
            "this-period vs other-period amounts and compares it to what the bank statement "
            "actually shows for that UTR."
        )
        # "Matched" as a literal Remarks value is now only a rare
        # degenerate fallback (2026-08-23: Remarks is a composed phrase
        # string - see engine.bank._classify_utr_remark) - a cleanly-tied-
        # out row is any Remarks ending in that suffix.
        matched = int(utr_bank_reco_df["Remarks"].str.endswith("Ties to bank credit").sum())
        needs_review = len(utr_bank_reco_df) - matched
        c1, c2 = st.columns(2)
        c1.metric("UTRs matched cleanly", f"{matched:,}")
        c2.metric("UTRs needing review", f"{needs_review:,}")
        st.caption(
            "Rows not ending \"Ties to bank credit\" are a starting point for the reconciling "
            "accountant's manual review, not a final answer - see engine/bank.py's "
            "docstring for exactly how each Remarks category is decided."
        )
        st.dataframe(utr_bank_reco_df, use_container_width=True, height=420)
    elif st.session_state.get("bank_ledger_df") is not None:
        st.divider()
        st.info(
            "No settlement UTRs could be matched between the gateway receipts and the "
            "uploaded bank statement for this run."
        )
