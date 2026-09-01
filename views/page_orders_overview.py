"""
page_orders_overview.py
------------------------
Order-level browsing: search a specific order, or scan the full table.
"""

import pandas as pd
import streamlit as st

from views.theme import require_data_or_prompt


def render():
    st.title("Orders Overview")
    if not require_data_or_prompt("Upload Data"):
        return

    lookup_df = st.session_state.get("lookup_df")
    reco_df = st.session_state["reco_df"]

    tab1, tab2 = st.tabs(["Order Lookup", "Full detail"])

    with tab1:
        if lookup_df is None:
            st.info("Order lookup detail isn't available for this run.")
        else:
            with st.form("order_search_form"):
                search_input = st.text_input("Search by Order ID (e.g. 12345 or #12345)")
                submitted = st.form_submit_button("Search")

            if submitted:
                st.session_state["last_order_search"] = search_input.strip()

            current_search = st.session_state.get("last_order_search", "")

            if current_search:
                clean_search = current_search.replace("#", "")
                match = lookup_df[lookup_df["Order ID"].astype(str).str.replace("#", "") == clean_search]
                if match.empty:
                    st.warning(f"No order found matching '{current_search}'.")
                else:
                    row = match.iloc[0]
                    c1, c2, c3 = st.columns(3)
                    fields = list(row.index)
                    third = len(fields) // 3 + 1
                    for col, chunk in zip((c1, c2, c3), [fields[:third], fields[third:2*third], fields[2*third:]]):
                        for field in chunk:
                            col.metric(field, str(row[field]) if pd.notna(row[field]) else "—")
                    if st.button("Clear search"):
                        st.session_state["last_order_search"] = ""
                        st.rerun()
            else:
                st.caption("Type an Order ID above and click Search, or browse the full lookup table below.")
                st.dataframe(lookup_df, use_container_width=True, height=450)

    with tab2:
        st.dataframe(reco_df, use_container_width=True, height=500)
