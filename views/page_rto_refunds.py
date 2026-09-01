"""
page_rto_refunds.py
--------------------
Focused view on RTO and refund status - the COD-aware logic that only
flags a refund as pending when money was actually received and not yet
refunded (see engine/reco.py flag_queries for the full explanation).
"""

import streamlit as st

from engine.summary import status_summary
from views.theme import require_data_or_prompt
from views.charts import donut, horizontal_bar


def render():
    st.title("RTO & Refunds")
    if not require_data_or_prompt("Upload Data"):
        return

    reco_df = st.session_state["reco_df"]
    s_df = status_summary(reco_df)

    rto_df = reco_df[reco_df["final_delivery_status"] == "RTO"]
    cancelled_df = reco_df[reco_df["final_delivery_status"] == "Cancelled"]
    lost_df = reco_df[reco_df["final_delivery_status"] == "Lost"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total RTO", f"{len(rto_df):,}")
    c2.metric("Total Cancelled", f"{len(cancelled_df):,}")
    c3.metric("Total Lost", f"{len(lost_df):,}")
    refund_pending_df = reco_df[reco_df["query"].str.contains("not refunded", na=False)]
    c4.metric("Refund Pending", f"{len(refund_pending_df):,}", help="Money received, not yet refunded")

    st.divider()

    left, right = st.columns(2)
    with left:
        donut(s_df, "final_delivery_status", "orders", "Delivery status distribution")
    with right:
        rto_reasons = reco_df[reco_df["final_delivery_status"].isin(["RTO", "Cancelled", "Lost"])]
        by_query = rto_reasons.groupby("query").size().reset_index(name="orders")
        if len(by_query) > 0:
            horizontal_bar(by_query, "orders", "query", "RTO/Cancelled/Lost breakdown by outcome")
        else:
            st.info("No RTO, Cancelled, or Lost orders in this dataset.")

    st.divider()
    st.subheader("Refund Pending — needs action")
    st.caption(
        "Only shows cases where money was actually received and not yet refunded. "
        "COD orders that were RTO'd with nothing collected are correctly excluded — "
        "there's nothing to refund on those."
    )
    if refund_pending_df.empty:
        st.success("No refund-pending orders — nothing outstanding.")
    else:
        cols = ["order_id", "month", "final_delivery_status", "total", "receipt_amount", "refund_amount", "query"]
        st.dataframe(refund_pending_df[[c for c in cols if c in refund_pending_df.columns]],
                     use_container_width=True, height=350)
