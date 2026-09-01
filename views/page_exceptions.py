"""
page_exceptions.py
--------------------
Full order-level exception detail - not just a summary count. Every
applicable order gets one row with delivery partner, Unicommerce status,
payment method, bank/UTR status, exception type/reason, amount, and
reconciliation status - enough to trace the exception back to source.
Filterable by Financial Year, Month, Date range, Sales Channel, Exception
Type, and Reconciliation Status before download.
"""

import io
import pandas as pd
import streamlit as st

from engine import storage
from engine.lookup import build_exception_detail
from engine.formatting import strip_tz
from views.filters import render_filter_controls, load_combined, trim_to_date_range


def render():
    st.title("Exceptions")
    client_key = st.session_state.get("client_key")
    if not client_key:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    runs = storage.list_runs(client_key)
    if not runs:
        st.info("No reconciliation saved yet. Go to Upload Data / Reconciliation first.")
        return

    st.markdown("##### Filters")
    filtered_runs, (date_from, date_to) = render_filter_controls(runs, key_prefix="exceptions")

    if not filtered_runs:
        st.warning("No saved months match the current filters.")
        return

    reco_df, lookup_df = load_combined(client_key, filtered_runs)
    reco_df = trim_to_date_range(reco_df, "created_at", date_from, date_to)

    config = st.session_state.get("config") or {}
    detail_df = build_exception_detail(reco_df, lookup_df, config.get("channel_name"))

    c1, c2 = st.columns(2)
    with c1:
        exc_types = sorted(detail_df["Exception Type"].dropna().unique())
        exc_type_choice = st.multiselect("Exception Type", exc_types, default=exc_types)
    with c2:
        recon_statuses = sorted(detail_df["Reconciliation Status"].dropna().unique())
        recon_choice = st.multiselect("Reconciliation Status", recon_statuses, default=recon_statuses)

    only_exceptions = st.checkbox("Show only orders with an open exception (exclude Reconciled)", value=True)

    view_df = detail_df[
        detail_df["Exception Type"].isin(exc_type_choice) &
        detail_df["Reconciliation Status"].isin(recon_choice)
    ]
    if only_exceptions:
        view_df = view_df[view_df["Reconciliation Status"] != "Reconciled"]

    st.caption(f"{len(view_df):,} order(s) in the current filtered view.")
    st.dataframe(view_df, use_container_width=True, height=450)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            # Same timezone-aware-datetime guard as Reports' workbook
            # builders (see engine/formatting.py's strip_tz) - Excel can't
            # write a tz-aware datetime column at all.
            strip_tz(view_df).to_excel(writer, sheet_name="Exceptions", index=False)
        st.download_button(
            "Download as Excel (.xlsx)",
            data=excel_buf.getvalue(),
            file_name="exception_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        csv_data = view_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download as CSV",
            data=csv_data,
            file_name="exception_report.csv",
            mime="text/csv",
        )
