"""
page_settings.py
------------------
Client/channel selection. Each config file under configs/ is one
client+channel combination - adding Amazon or Flipkart later is a new
config file here, not new code, and it'll show up in this list automatically.
"""

import streamlit as st

from views.state_init import set_active_config
from views.page_upload import reset_upload_state


def render():
    st.title("Settings")

    st.subheader("Active client / channel")
    labels = list(st.session_state["config_labels"].keys())
    current = st.session_state.get("chosen_label", labels[0] if labels else None)
    chosen = st.selectbox("Client / Channel", labels, index=labels.index(current) if current in labels else 0)

    if chosen != current:
        set_active_config(chosen)
        # Was previously a short, hand-rolled reset here that only cleared
        # a handful of DTC-shaped keys (orders_df/delivery_frames/
        # gateway_frames/bank_df) and never restored anything from disk -
        # so switching channel from THIS page (as opposed to Upload Data's
        # own platform selector, which already went through the shared,
        # correct path) silently dropped every already-confirmed raw
        # upload out of session_state for good, made Data Management's
        # per-file Delete option disappear, and never even touched
        # mtr_files/settlement_files/attribution_frames at all - so an
        # Amazon channel's raw files could leak into view after switching
        # away and back. reset_upload_state() (views/page_upload.py) is
        # the single correct implementation of "switching channel" - it
        # clears every raw-upload AND derived-result key for every channel
        # shape, then restores whatever's still persisted on disk for the
        # newly-active client_key - so this page now behaves identically
        # to Upload Data's own channel switch instead of a second,
        # incomplete copy of it.
        reset_upload_state()
        st.success(f"Switched to {chosen}. Upload files for this channel to get started.")
        st.rerun()

    st.divider()
    st.subheader("Available channels")
    st.caption(
        "Every config file in the configs/ folder shows up here as a selectable channel. "
        "To add Amazon, Flipkart, or another client, add a new config JSON with that "
        "source's column mappings - no code changes needed."
    )
    for label in labels:
        st.write(f"• {label}")
