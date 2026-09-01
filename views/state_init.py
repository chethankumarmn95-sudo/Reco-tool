"""
state_init.py
-------------
Runs once at the top of app.py to make sure session_state has the basics
every page relies on: which client/channel config is active. Individual
pages don't need to worry about "what if config isn't set yet" - this
guarantees it always is, defaulting to the first config found.
"""

import glob
import json
import os
import streamlit as st

from engine import storage


def init_state():
    if "config_labels" not in st.session_state:
        config_files = sorted(glob.glob("configs/*.json"))
        config_labels, config_paths = {}, {}
        for path in config_files:
            with open(path) as f:
                cfg = json.load(f)
            label = f"{cfg['client_name']} — {cfg['channel_name']}"
            config_labels[label] = cfg
            config_paths[label] = path
        st.session_state["config_labels"] = config_labels
        st.session_state["config_paths"] = config_paths

        if config_labels:
            first_label = list(config_labels.keys())[0]
            st.session_state["chosen_label"] = first_label
            st.session_state["config"] = config_labels[first_label]
            st.session_state["client_key"] = os.path.splitext(
                os.path.basename(config_paths[first_label])
            )[0]
            # Client-reported (2026-08-22): raw files that had genuinely
            # already been uploaded and confirmed - Shopify orders,
            # Razorpay, etc. - looked completely gone (Data Management's
            # per-file delete list empty; Razorpay's own upload immediately
            # re-failing its "Shopify order report needs to be uploaded
            # first" prerequisite check) the moment the app process
            # restarted or the browser session expired, purely because
            # those uploads lived only in st.session_state and nowhere
            # durable. Restoring here - the one place a fresh session's
            # default client_key gets decided - means the very first page
            # a returning user lands on (not only Upload Data) already has
            # everything that client had confirmed before, exactly as if
            # the app had never restarted. See engine.storage's
            # RAW_UPLOAD_SLOTS/"Persisted raw uploads" section for the full
            # mechanism (this mirrors views/page_upload.py's
            # reset_upload_state, which handles the OTHER case - explicitly
            # switching client/platform mid-session, from either the
            # Upload Data or Settings page).
            #
            # Client-reported again (2026-08-26): the Gokwik Order/
            # Transaction "attribution" reports (added 2026-08-25) went
            # through the exact same "Delete option vanishes after
            # reopening" symptom, because this loop still hard-coded the
            # OLDER, shorter key list rather than importing
            # storage.RAW_UPLOAD_SESSION_KEYS - attribution_frames was
            # simply never in it. Iterating the shared constant instead of
            # a separately hand-maintained tuple is what keeps this from
            # silently drifting out of sync again the next time a new raw
            # upload type is added anywhere in the app.
            restored = storage.load_raw_uploads(st.session_state["client_key"])
            for key in storage.RAW_UPLOAD_SESSION_KEYS:
                if key in restored:
                    st.session_state[key] = restored[key]

            # Recycle Bin items past their retention window get purged as a
            # lazy sweep (see storage.purge_expired_recycle_bin_items's
            # docstring - this app has no background scheduler) whenever
            # Data Management is visited, scoped to whichever client/
            # channel is active there. Sweeping every configured client/
            # channel once here too, at the one guaranteed point every
            # fresh session passes through, means a channel the user
            # doesn't happen to open Data Management for this session still
            # gets its own expired items purged automatically rather than
            # only whenever someone next visits that specific channel's
            # Data Management page.
            for other_label, other_path in config_paths.items():
                other_key = os.path.splitext(os.path.basename(other_path))[0]
                storage.purge_expired_recycle_bin_items(other_key)

    for key, default in [
        ("reco_df", None), ("lookup_df", None), ("totals", None),
        ("consolidated", None), ("bank_status_df", None),
        ("gateway_settlement_df", None), ("utr_bank_reco_df", None), ("bank_ledger_df", None),
        ("recon_status_df", None), ("settlement_pending_df", None),
        ("settlement_pending_summary_df", None), ("gateway_health_df", None),
        ("cod_batches_df", None), ("cod_batch_match_df", None), ("bank_statement_uploaded", False),
        ("amazon_order_reco_df", None), ("amazon_waterfall_df", None), ("amazon_expense_ledger_df", None),
        ("amazon_settlement_register_df", None), ("amazon_settlement_summary_df", None),
        ("amazon_tie_out_df", None), ("amazon_non_mtr_df", None),
        ("amazon_cutoff_date", None), ("amazon_subsequent_settlements_df", None),
        # Every raw-upload slot (orders_df, bank_df, delivery_frames,
        # gateway_frames, attribution_frames, mtr_files, settlement_files)
        # defaults from the same canonical registry the restore loop above
        # uses, so a slot no config has ever produced a raw upload for
        # still gets correctly initialized ({} vs None) instead of being
        # missing from session_state entirely.
        *storage.raw_upload_defaults().items(),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default


def set_active_config(label):
    st.session_state["chosen_label"] = label
    st.session_state["config"] = st.session_state["config_labels"][label]
    path = st.session_state["config_paths"][label]
    st.session_state["client_key"] = os.path.splitext(os.path.basename(path))[0]


def get_client_channels(client_name):
    """Every config (channel) that belongs to the given client_name, as
    {"label", "client_key", "config"} dicts.

    Each channel (Amazon, the Shopify/DTC channel, future Flipkart/Tata
    1mg/First Cry, ...) is its own configs/*.json with its own client_key
    derived from the filename - so their saved runs already live in
    completely separate data_store/<client_key>/ folders and never
    overwrite one another. The only thing that was ever "single channel
    at a time" was the VIEW: Dashboard/Reports used to look at just
    st.session_state["client_key"] (whichever channel happens to be
    active from Settings/Upload Data), so switching channels made the
    other one's already-saved numbers disappear from the screen even
    though the underlying data was untouched. This groups every channel
    for one client together so Dashboard/Reports can show all of them at
    once (see their use of st.tabs) instead of just the active one.
    """
    config_labels = st.session_state.get("config_labels", {})
    config_paths = st.session_state.get("config_paths", {})
    channels = []
    for label, cfg in config_labels.items():
        if cfg.get("client_name") == client_name:
            path = config_paths.get(label, "")
            key = os.path.splitext(os.path.basename(path))[0]
            channels.append({"label": label, "client_key": key, "config": cfg})
    return channels
