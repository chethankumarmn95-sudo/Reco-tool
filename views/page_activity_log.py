"""
page_activity_log.py
----------------------
A simple history view built from saved-month metadata. Not a full audit
trail yet (that would need to log every action, not just saves) - this is
a first, honest version: it shows what's been saved and when.
"""

import streamlit as st

from engine import storage


def render():
    st.title("Activity Log")
    client_key = st.session_state.get("client_key")
    if not client_key:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    runs = storage.list_runs(client_key)
    if not runs:
        st.caption("No activity yet. Save a reconciliation from Data Management to see it here.")
        return

    st.caption("Every reconciliation saved for this client/channel, most recent first.")
    for r in runs:
        st.write(f"📁 **{r['month_label']}** — {r['order_count']:,} orders — saved {r['saved_at']}")
