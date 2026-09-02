"""
app.py
------
The website entrypoint. Run it with:  streamlit run app.py

This file is now just a thin shell: it sets up the sidebar navigation and
hands off to whichever page is selected. All the actual page content lives
in views/page_*.py - see that folder for the real logic.
"""

import streamlit as st

from views.auth import check_login, logout_button
from views.state_init import init_state
from views.theme import apply_theme
from views import page_portal
from views import (
    page_dashboard,
    page_upload,
    page_reconciliation,
    page_orders_overview,
    page_rto_refunds,
    page_bank_linking,
    page_reports,
    page_data_management,
    page_exceptions,
    page_settings,
    page_activity_log,
)

st.set_page_config(page_title="Reco Tool", layout="wide", page_icon="📊")

if not check_login():
    st.stop()

init_state()
apply_theme()

# The Authorised Portal hub - shown right after login, before any tool.
# Structured so more tools can be added here later without touching the
# login flow or the Reconciliation Tool's own navigation below.
if not st.session_state.get("entered_portal"):
    page_portal.render()
    with st.sidebar:
        st.markdown("### Reco Tool")
        st.caption("E-commerce Reconciliation")
        logout_button()
    st.stop()

pages = [
    st.Page(page_dashboard.render, title="Dashboard", icon="🏠", default=True, url_path="dashboard"),
    st.Page(page_upload.render, title="Upload Data", icon="📤", url_path="upload-data"),
    st.Page(page_reconciliation.render, title="Reconciliation", icon="🔄", url_path="reconciliation"),
    st.Page(page_orders_overview.render, title="Orders Overview", icon="📦", url_path="orders-overview"),
    st.Page(page_rto_refunds.render, title="RTO & Refunds", icon="↩️", url_path="rto-refunds"),
    st.Page(page_bank_linking.render, title="Bank Statement / UTR Linking", icon="🏦", url_path="bank-linking"),
    st.Page(page_reports.render, title="Reports", icon="📑", url_path="reports"),
    st.Page(page_data_management.render, title="Data Management", icon="🗂️", url_path="data-management"),
    st.Page(page_exceptions.render, title="Exceptions", icon="⚠️", url_path="exceptions"),
    st.Page(page_settings.render, title="Settings", icon="⚙️", url_path="settings"),
    st.Page(page_activity_log.render, title="Activity Log", icon="🕒", url_path="activity-log"),
]

with st.sidebar:
    st.markdown("### Reco Tool")
    st.caption("E-commerce Reconciliation")
    if st.button("‹ Back to Portal", use_container_width=True):
        st.session_state["entered_portal"] = False
        st.rerun()
    logout_button()

nav = st.navigation(pages, position="sidebar")
nav.run()
