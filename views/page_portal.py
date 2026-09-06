"""
page_portal.py
---------------
The "Authorised Portal" hub. Shown right after a successful login, before
the Reconciliation Tool itself - this is the landing spot inside the
authenticated area, structured so more tools (MIS & Analytics, etc.) can
be added here later as additional cards without changing how login or
navigation works.

Selecting a tool sets st.session_state["entered_portal"] = True, which is
what app.py checks to decide whether to show this hub or the tool's own
sidebar navigation.
"""

import streamlit as st


def render():
    username = st.session_state.get("username", "")

    st.markdown(
        f"""
        <div style="padding: 4px 0 6px;">
            <div style="font-size:0.78rem; letter-spacing:0.06em; text-transform:uppercase; color:#8888A0;">
                Authorised Portal
            </div>
            <h1 style="margin:6px 0 4px;">Welcome, {username}</h1>
            <p style="color:#6B6B80; margin:0;">Choose a tool to open.</p>
        </div>
        <br>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        with st.container(border=True):
            st.markdown("#### 📊 Ecom Reco 360")
            st.caption(
                "Order, payment and settlement matching across every channel — "
                "Shopify, Amazon, Flipkart, Tata 1mg, First Cry and bank statements."
            )
            st.markdown("")
            if st.button(
                "Open Ecom Reco 360 →",
                key="open_reco_tool",
                type="primary",
                use_container_width=True,
            ):
                st.session_state["entered_portal"] = True
                st.rerun()

    with col2:
        with st.container(border=True):
            st.markdown("#### 📈 MIS & Analytics")
            st.caption("Cross-tool financial reporting and trend analysis.")
            st.markdown("")
            st.button(
                "Coming soon",
                key="mis_analytics_soon",
                disabled=True,
                use_container_width=True,
            )
