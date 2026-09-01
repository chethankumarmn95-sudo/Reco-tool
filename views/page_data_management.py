"""
page_data_management.py
-------------------------
Save the current reconciliation as a named month, load one or more saved
months back (combining them for cross-month reporting), or delete a saved
month with a confirmation step.

Branches on channel_type: the DTC/Shopify flow (_render_dtc) saves/loads
reco_df+lookup_df via engine.storage.save_run/list_runs/combine_runs; the
marketplace/Amazon flow (_render_amazon) saves/loads the very different
order_reco_df/waterfall_df/expense_ledger_df/... result set via the
parallel engine.storage.save_amazon_run/list_amazon_runs/combine_amazon_runs
added alongside those. Saving here is what makes an Amazon run show up on
Dashboard/Reports/Bank Linking too - those pages read from the SAME saved
periods (via list_amazon_runs), exactly like the DTC pages already do.

Also home to individual raw-file deletion (2026-08-21 client request) -
see the module-section docstring just below for the full story: this is
DIFFERENT from "delete a saved month/period" above, which removes an
entire already-saved dataset. The two are the "two levels of deletion"
the client explicitly asked for - individual raw file (this section,
works on the CURRENT in-progress upload before/after running
reconciliation) vs entire saved dataset (the saved-month/period delete
above, for intentionally starting completely over).

Both deletion levels are SOFT deletes as of 2026-08-22 (see the "Recycle
Bin" section further below) - neither one immediately, permanently erases
anything; both move their data into engine.storage's Recycle Bin, where it
sits for a retention window (default 30 days) fully excluded from
reconciliation/duplicate-checking, restorable at any point during that
window, and auto-purged for good once it expires.
"""

import streamlit as st

from engine import dedup, storage
from engine.summary import headline_totals
from views import page_reconciliation


# ---------------------------------------------------------------------------
# Individual raw file deletion
# ---------------------------------------------------------------------------
# Client-reported (2026-08-21): the only "Delete" option in Data Management
# removed an entire saved dataset/period - there was no way to remove just
# ONE raw report (e.g. "the Razorpay report was uploaded incorrectly")
# while keeping every other uploaded report (Shopify orders, Shiprocket,
# Delhivery, Unicommerce, Gokwik, ...) exactly as they were.
#
# Deleting one raw file correctly needs FOUR things to happen together,
# or the fix is incomplete:
#   1. Remove that file's data from session_state - and ONLY that file's
#      data (a dict-shaped slot like gateway_frames/delivery_frames/
#      mtr_files/settlement_files loses just its one entry; a scalar slot
#      like orders_df/bank_df is cleared entirely, since there's only ever
#      one file in that slot to begin with).
#   2. Clear that report type's ingestion ledger (see engine/storage.py's
#      clear_seen_keys/clear_latest_records) - otherwise re-uploading the
#      exact same file afterwards still shows "already reconciled - 0 new
#      rows" against a ledger entry whose underlying data was just deleted
#      from the session entirely (the client's own explicit complaint).
#   3. Move a snapshot of both (the removed data AND the exact ledger
#      entries just cleared) into the Recycle Bin (engine.storage's
#      soft_delete_raw_file/restore_raw_file_from_recycle_bin - see that
#      module's "Recycle Bin" section) rather than just discarding them -
#      2026-08-22 client request for a 30-day-retention safety net on
#      every delete, not only the "delete a saved month/period" flow
#      further below.
#   4. Recalculate every downstream reconciliation layer that depended on
#      the deleted file, automatically, right away - not leave the
#      Reconciliation page showing stale numbers until someone happens to
#      revisit it and click "Run reconciliation" again. Done here by
#      calling the exact same run_dtc_reconciliation()/
#      run_marketplace_reconciliation() functions the Reconciliation
#      page's own button calls (see views/page_reconciliation.py) - one
#      shared implementation of "run the pipeline", triggered from two
#      places, so they can never quietly drift apart.
#
# If there's no longer enough data to reconcile at all after the deletion
# (e.g. the orders file itself was just deleted, or every MTR segment
# was), step 4 is skipped and the existing reconciliation results are
# simply cleared instead - showing stale numbers computed from data that
# no longer fully exists would be worse than showing nothing.

_DERIVED_DTC_KEYS = (
    "reco_df", "lookup_df", "consolidated", "totals", "recon_status_df",
    "cod_batches_df", "cod_batch_match_df", "gateway_settlement_df",
    "settlement_pending_df", "settlement_pending_summary_df",
    "utr_bank_reco_df", "bank_ledger_df", "bank_statement_uploaded",
)

_DERIVED_AMAZON_KEYS = (
    "amazon_cutoff_date", "amazon_order_reco_df", "amazon_waterfall_df",
    "amazon_expense_ledger_df", "amazon_settlement_summary_df",
    "amazon_settlement_register_df", "amazon_tie_out_df", "amazon_non_mtr_df",
    "amazon_subsequent_settlements_df", "bank_statement_uploaded",
)


def _clear_derived_results(derived_keys):
    for key in derived_keys:
        st.session_state[key] = None


def _dtc_raw_file_specs(config):
    """
    Every individually-deletable raw upload slot for the DTC/Shopify flow,
    in the shape the delete UI below needs: display label, where its data
    lives in session_state (dict_entry=None for a scalar slot like
    orders_df/bank_df; the dict key within session_key for a multi-file
    slot like delivery_frames/gateway_frames), and which ingestion-ledger
    bucket(s) (see engine/storage.py) need clearing alongside it. Ledger
    kind ("seen" vs "latest") and report_key naming mirror exactly what
    views/page_upload.py's render_duplicate_check() and
    views/page_reconciliation.py's _update_dtc_ledger() already use for
    this same report, so a delete here and an upload/reconcile there can
    never disagree about which ledger file a given report's data lives in.
    Bank statement has no ledger entry at all - it's never run through
    render_duplicate_check (see page_upload.py's _render_bank_upload).
    """
    specs = [{
        "label": config.get("orders", {}).get("label", "Orders"),
        "session_key": "orders_df", "dict_entry": None,
        "ledger": [("orders", "seen")],
    }]
    for d_cfg in config.get("delivery_partners", []):
        specs.append({
            "label": d_cfg["label"], "session_key": "delivery_frames", "dict_entry": d_cfg["label"],
            "ledger": [(f"delivery__{d_cfg['label']}", "latest")],
        })
    for g_cfg in config.get("gateways", []):
        specs.append({
            "label": g_cfg["label"], "session_key": "gateway_frames", "dict_entry": g_cfg["label"],
            "ledger": [(f"gateway__{g_cfg['label']}", "latest")],
        })
    for a_cfg in config.get("attribution_sources", []):
        specs.append({
            "label": a_cfg["label"], "session_key": "attribution_frames", "dict_entry": a_cfg["label"],
            "ledger": [(f"attribution__{a_cfg['label']}", "latest")],
        })
    if "bank_statement" in config:
        specs.append({
            "label": config["bank_statement"]["label"], "session_key": "bank_df", "dict_entry": None,
            "ledger": [],
        })
    return specs


def _amazon_raw_file_specs(config):
    """_dtc_raw_file_specs()'s equivalent for the Amazon/marketplace flow -
    MTR reports (one per segment) and Settlement Flat Files (one per
    payment mode) instead of orders/delivery-partners/gateways, mirroring
    _update_amazon_ledger()'s report_key naming exactly."""
    specs = []
    for r_cfg in config.get("mtr_reports", []):
        specs.append({
            "label": r_cfg["label"], "session_key": "mtr_files", "dict_entry": r_cfg["segment"],
            "ledger": [(f"mtr__{r_cfg['segment']}", "seen")],
        })
    for s_cfg in config.get("settlement_files", []):
        specs.append({
            "label": s_cfg["label"], "session_key": "settlement_files", "dict_entry": s_cfg["payment_mode"],
            "ledger": [(f"settlement__{s_cfg['payment_mode']}", "latest")],
        })
    if "bank_statement" in config:
        specs.append({
            "label": config["bank_statement"]["label"], "session_key": "bank_df", "dict_entry": None,
            "ledger": [],
        })
    return specs


def _is_uploaded(spec):
    """Whether this raw file slot currently holds real data - a dict slot
    with no entry for this label, or a scalar slot that's None/empty,
    both mean "not currently uploaded" and are simply left out of the
    delete list (nothing to delete)."""
    if spec["dict_entry"] is not None:
        frames = st.session_state.get(spec["session_key"]) or {}
        df = frames.get(spec["dict_entry"])
    else:
        df = st.session_state.get(spec["session_key"])
    return df is not None and not (hasattr(df, "empty") and df.empty)


def _delete_raw_file(client_key, spec):
    """
    Removes exactly one raw file's data from the current session, leaving
    every other uploaded report untouched, clears its ingestion-ledger
    entries (see engine/storage.py), and moves a snapshot of both into the
    Recycle Bin so the delete is undoable within the retention window -
    see this module's section docstring above for why all three matter
    together.
    """
    if spec["dict_entry"] is not None:
        frames = dict(st.session_state.get(spec["session_key"]) or {})
        removed_data = frames.pop(spec["dict_entry"], None)
        st.session_state[spec["session_key"]] = frames
    else:
        removed_data = st.session_state.get(spec["session_key"])
        st.session_state[spec["session_key"]] = None

    ledger_removals = []
    for report_key, ledger_kind in spec["ledger"]:
        if ledger_kind == "seen":
            removed_keys = storage.get_seen_keys(client_key, report_key)
            if removed_keys:
                ledger_removals.append({"report_key": report_key, "ledger_kind": "seen", "keys": sorted(removed_keys)})
            storage.clear_seen_keys(client_key, report_key)
        else:
            removed_records = storage.get_latest_records(client_key, report_key)
            if removed_records:
                ledger_removals.append({"report_key": report_key, "ledger_kind": "latest", "keys": sorted(removed_records.keys())})
            storage.clear_latest_records(client_key, report_key)

    storage.soft_delete_raw_file(
        client_key, spec["label"], spec["session_key"], spec["dict_entry"], removed_data, ledger_removals,
    )
    # Also remove the separate always-current on-disk copy this file's
    # upload maintained for restart durability (see engine.storage's
    # "Persisted raw uploads" section) - it's already safely captured in
    # the Recycle Bin snapshot above, so this is purely "stop treating it
    # as currently uploaded", not a second, unprotected delete.
    storage.clear_raw_upload(client_key, spec["session_key"], spec["dict_entry"])


def _recalculate_or_clear(config, plumbing, prefix):
    """
    Shared "clear stale results, then either recalculate or admit there's
    not enough left to" step, used after BOTH a raw-file deletion and a
    raw-file RESTORE from the Recycle Bin - the two are symmetric (each
    changes what's in session_state and needs the same follow-up), so this
    is the one place that logic lives rather than two copies that could
    quietly drift apart.

    plumbing: the dict of channel-specific callables built in _render_dtc/
    _render_amazon (derived_keys/can_recalculate/run_fn/build_message/
    no_data_message - see _render_raw_file_deletion below for what each
    one means). prefix: the flash message's leading clause, e.g. "Deleted
    Razorpay." or "Restored Razorpay.".

    Sets st.session_state["_data_mgmt_flash"] for render() to show after
    the rerun this always triggers via its caller.
    """
    _clear_derived_results(plumbing["derived_keys"])
    if plumbing["can_recalculate"]():
        with st.spinner("Recalculating reconciliation..."):
            result = plumbing["run_fn"](config)
        st.session_state["_data_mgmt_flash"] = ("success", f"{prefix} {plumbing['build_message'](result)}")
    else:
        st.session_state["_data_mgmt_flash"] = ("info", f"{prefix} {plumbing['no_data_message']}")


def _render_raw_file_deletion(client_key, config, specs, plumbing):
    """
    Shared "Uploaded raw files" list + per-file Delete button + confirm
    step, used by both _render_dtc and _render_amazon below (identical UI
    and flow either way - only the specs list and plumbing differ between
    the two channel shapes).

    plumbing: {"derived_keys", "can_recalculate", "run_fn", "build_message",
    "no_data_message"} - see _recalculate_or_clear above for what each does;
    built once per channel in _render_dtc/_render_amazon and shared with
    the Recycle Bin's restore action too, so a delete and its own restore
    always recalculate the exact same way.
    """
    st.subheader("Uploaded raw files")
    st.caption(
        "Delete one specific raw report without touching any of the others - the "
        "reconciliation is automatically recalculated from what's left, right away. "
        "Deleted reports move to the Recycle Bin further down this page rather than "
        "disappearing immediately - restore one from there if this was a mistake."
    )
    uploaded_specs = [s for s in specs if _is_uploaded(s)]
    if not uploaded_specs:
        st.caption("No raw files uploaded yet for this channel.")
        st.divider()
        return

    for spec in uploaded_specs:
        c1, c2, c3 = st.columns([3, 1, 1])
        c1.write(spec["label"])
        c2.write("Uploaded")
        widget_id = f"{client_key}_{spec['session_key']}_{spec['dict_entry']}"
        confirm_key = f"confirm_delete_raw_{widget_id}"
        if c3.button("Delete", key=f"delbtn_raw_{widget_id}"):
            st.session_state[confirm_key] = True
        if st.session_state.get(confirm_key):
            st.warning(
                f"Delete **{spec['label']}** raw report? This moves the file and its related "
                "reconciliation results to the Recycle Bin (kept for "
                f"{storage.RECYCLE_BIN_RETENTION_DAYS} days, restorable any time before then). "
                "Other uploaded reports will not be affected."
            )
            yc, nc = st.columns(2)
            if yc.button("Yes, delete", key=f"yes_raw_{widget_id}"):
                _delete_raw_file(client_key, spec)
                st.session_state[confirm_key] = False
                _recalculate_or_clear(config, plumbing, f"Deleted {spec['label']}.")
                st.rerun()
            if nc.button("No, cancel", key=f"no_raw_{widget_id}"):
                st.session_state[confirm_key] = False
                st.rerun()

    st.divider()


def _render_recycle_bin(client_key, config, plumbing):
    """
    Lists every soft-deleted item still within its retention window
    (raw-file deletes from _render_raw_file_deletion above, AND whole
    saved-period deletes from the "Delete a saved month/period" section
    below - both land in the same per-client bin, see engine/storage.py's
    "Recycle Bin" section) with a Restore and a "Delete permanently"
    action each - the client's explicit ask for a visible bin showing
    "deleted files/data and their deletion date", plus the restore path
    that's the actual point of calling it a Recycle Bin rather than just a
    delayed permanent delete.

    Also runs the automatic 30-day purge as a lazy sweep (see
    storage.purge_expired_recycle_bin_items's docstring for why a lazy
    sweep-on-visit is the right model for an app with no background
    scheduler) - cheap even with several items in the bin, since it only
    reads each item's small manifest.json.

    A saved-period item's restore doesn't need plumbing at all (putting
    its .pkl back doesn't touch session_state or require recalculating
    anything - see storage.restore_run_from_recycle_bin's docstring);
    plumbing is only actually used for a "raw_file" item's restore, which
    changes session_state and does need to recalculate.
    """
    purged = storage.purge_expired_recycle_bin_items(client_key)
    if purged:
        st.caption(f"({purged} item(s) past their {storage.RECYCLE_BIN_RETENTION_DAYS}-day retention were purged automatically.)")

    items = storage.list_recycle_bin(client_key)
    with st.expander(f"🗑️ Recycle Bin ({len(items)})", expanded=False):
        st.caption(
            f"Deleted reports and saved periods are kept here for {storage.RECYCLE_BIN_RETENTION_DAYS} days "
            "before being purged automatically. Nothing in here is used for reconciliation or "
            "duplicate-checking while it sits in the bin."
        )
        if not items:
            st.caption("Recycle Bin is empty.")
            return

        for item in items:
            c1, c2, c3, c4 = st.columns([3, 2, 1, 1])
            kind_label = {"raw_file": "Raw file", "dtc": "Saved period", "amazon": "Saved period"}.get(item["kind"], item["kind"])
            detail = f"{item['label']}" if item["kind"] == "raw_file" else f"{item.get('month_label', item['label'])} — {item.get('order_count', 0):,} orders"
            c1.write(f"**{kind_label}**: {detail}")
            c2.write(f"Deleted {item['deleted_at'][:10]} — purges in {item['days_remaining']} day(s)")
            widget_id = f"{client_key}_{item['item_id']}"
            if c3.button("Restore", key=f"restore_{widget_id}"):
                if item["kind"] == "raw_file":
                    restored = storage.restore_raw_file_from_recycle_bin(client_key, item["item_id"])
                    if restored is not None:
                        if restored["dict_entry"] is not None:
                            frames = dict(st.session_state.get(restored["session_key"]) or {})
                            frames[restored["dict_entry"]] = dedup.accumulate_df(
                                frames.get(restored["dict_entry"]), restored["data"],
                            )
                            st.session_state[restored["session_key"]] = frames
                            # Symmetric with _delete_raw_file's clear_raw_upload
                            # call - a restored raw file needs to survive the
                            # next app restart too, same as any other current
                            # upload (see engine.storage's "Persisted raw
                            # uploads" section).
                            storage.save_raw_upload(
                                client_key, restored["session_key"], restored["dict_entry"],
                                frames[restored["dict_entry"]],
                            )
                        else:
                            st.session_state[restored["session_key"]] = dedup.accumulate_df(
                                st.session_state.get(restored["session_key"]), restored["data"],
                            )
                            storage.save_raw_upload(
                                client_key, restored["session_key"], None,
                                st.session_state[restored["session_key"]],
                            )
                        _recalculate_or_clear(config, plumbing, f"Restored {item['label']}.")
                else:
                    fname = storage.restore_run_from_recycle_bin(client_key, item["item_id"])
                    if fname:
                        st.session_state["_data_mgmt_flash"] = ("success", f"Restored '{item.get('month_label', fname)}'.")
                st.rerun()
            confirm_key = f"confirm_purge_{widget_id}"
            if c4.button("Delete permanently", key=f"purge_{widget_id}"):
                st.session_state[confirm_key] = True
            if st.session_state.get(confirm_key):
                st.warning(f"Permanently delete **{item['label']}** from the Recycle Bin? This can't be undone.")
                yc, nc = st.columns(2)
                if yc.button("Yes, delete permanently", key=f"yespurge_{widget_id}"):
                    storage.permanently_delete_recycle_bin_item(client_key, item["item_id"])
                    st.session_state[confirm_key] = False
                    st.rerun()
                if nc.button("No, cancel", key=f"nopurge_{widget_id}"):
                    st.session_state[confirm_key] = False
                    st.rerun()


def _render_dtc(client_key):
    config = st.session_state.get("config") or {}
    plumbing = {
        "derived_keys": _DERIVED_DTC_KEYS,
        "can_recalculate": lambda: st.session_state.get("orders_df") is not None,
        "run_fn": page_reconciliation.run_dtc_reconciliation,
        "build_message": lambda result: (
            f"Reconciliation recalculated - {result['order_count']:,} orders reconciled."
            f"{result['bank_note']} {result['save_note']}"
        ),
        "no_data_message": (
            "Orders data is no longer available, so reconciliation results were cleared - "
            "upload an orders file and run reconciliation again on the Reconciliation page."
        ),
    }

    _render_raw_file_deletion(client_key, config, specs=_dtc_raw_file_specs(config), plumbing=plumbing)

    st.subheader("Save current reconciliation")
    reco_df = st.session_state.get("reco_df")
    if reco_df is None:
        st.info("No reconciliation loaded right now - run one on the Reconciliation page first.")
    else:
        default_label = storage.detect_month_label(reco_df)
        save_label = st.text_input("Save as", value=default_label, key="save_label_input")
        if st.button("Save this reconciliation"):
            totals = st.session_state.get("totals") or headline_totals(reco_df)
            config = st.session_state.get("config") or {}
            storage.save_run(client_key, save_label, reco_df, st.session_state.get("lookup_df"), totals,
                              config.get("channel_name"))
            st.success(f"Saved as '{save_label}'.")
            st.rerun()

    st.divider()
    st.subheader("Saved months")
    saved_runs = storage.list_runs(client_key)

    if not saved_runs:
        st.caption("No saved months yet.")
    else:
        run_labels = {
            f"{r['month_label']} — {r['order_count']:,} orders — saved {r['saved_at'][:10]}": r["file"]
            for r in saved_runs
        }
        selected = st.multiselect("Select one or more saved months to load", list(run_labels.keys()))
        if st.button("Load selected", disabled=(len(selected) == 0)):
            fnames = [run_labels[s] for s in selected]
            combined_reco, combined_lookup = storage.combine_runs(client_key, fnames)
            st.session_state["reco_df"] = combined_reco
            st.session_state["lookup_df"] = combined_lookup
            st.session_state["totals"] = headline_totals(combined_reco)
            st.success(f"Loaded {len(selected)} saved month(s) — {len(combined_reco):,} orders total. "
                       "Go to Dashboard to view.")

        st.divider()
        st.caption("Delete a saved month (entire dataset for that period)")
        for r in saved_runs:
            dcol1, dcol2 = st.columns([4, 1])
            dcol1.write(f"{r['month_label']} — {r['order_count']:,} orders — saved {r['saved_at'][:10]}")
            confirm_key = f"confirm_delete_{r['file']}"
            if dcol2.button("Delete", key=f"delete_btn_{r['file']}"):
                st.session_state[confirm_key] = True
            if st.session_state.get(confirm_key):
                st.warning(
                    f"Delete '{r['month_label']}'? This moves it to the Recycle Bin (kept for "
                    f"{storage.RECYCLE_BIN_RETENTION_DAYS} days, restorable any time before then)."
                )
                yc, nc = st.columns(2)
                if yc.button("Yes, delete", key=f"yes_{r['file']}"):
                    storage.soft_delete_run(client_key, r["file"], report_keys_to_scope=["orders"])
                    st.session_state[confirm_key] = False
                    st.rerun()
                if nc.button("No, cancel", key=f"no_{r['file']}"):
                    st.session_state[confirm_key] = False
                    st.rerun()

    st.divider()
    _render_recycle_bin(client_key, config, plumbing)


def _render_amazon(client_key, config):
    plumbing = {
        "derived_keys": _DERIVED_AMAZON_KEYS,
        "can_recalculate": lambda: bool(st.session_state.get("mtr_files")) and bool(st.session_state.get("settlement_files")),
        "run_fn": page_reconciliation.run_marketplace_reconciliation,
        "build_message": lambda result: (
            f"Reconciliation recalculated - {result['order_count']:,} MTR orders reconciled."
            f"{result['bank_note']}{result['cutoff_note']}{result['subsequent_note']} {result['save_note']}"
        ),
        "no_data_message": (
            "MTR report and/or Settlement Flat File data is no longer available, so "
            "reconciliation results were cleared - upload the missing report(s) and run "
            "reconciliation again on the Reconciliation page."
        ),
    }
    mtr_report_keys = [f"mtr__{r_cfg['segment']}" for r_cfg in config.get("mtr_reports", [])]

    _render_raw_file_deletion(client_key, config, specs=_amazon_raw_file_specs(config), plumbing=plumbing)

    st.subheader("Save current reconciliation")
    order_reco_df = st.session_state.get("amazon_order_reco_df")
    if order_reco_df is None or order_reco_df.empty:
        st.info("No reconciliation loaded right now - run one on the Reconciliation page first.")
    else:
        default_label = storage.detect_amazon_month_label(order_reco_df)
        save_label = st.text_input("Save as", value=default_label, key="save_label_input_amazon")
        if st.button("Save this reconciliation", key="save_amazon_btn"):
            storage.save_amazon_run(
                client_key, save_label,
                order_reco_df=order_reco_df,
                waterfall_df=st.session_state.get("amazon_waterfall_df"),
                expense_ledger_df=st.session_state.get("amazon_expense_ledger_df"),
                settlement_summary_df=st.session_state.get("amazon_settlement_summary_df"),
                settlement_register_df=st.session_state.get("amazon_settlement_register_df"),
                non_mtr_df=st.session_state.get("amazon_non_mtr_df"),
                tie_out_df=st.session_state.get("amazon_tie_out_df"),
                subsequent_settlements_df=st.session_state.get("amazon_subsequent_settlements_df"),
                cutoff_date=st.session_state.get("amazon_cutoff_date"),
                channel_name=config.get("channel_name"),
            )
            st.success(f"Saved as '{save_label}'.")
            st.rerun()

    st.divider()
    st.subheader("Saved periods")
    saved_runs = storage.list_amazon_runs(client_key)

    if not saved_runs:
        st.caption("No saved periods yet.")
    else:
        run_labels = {
            f"{r['month_label']} — {r['order_count']:,} orders — saved {r['saved_at'][:10]}": r["file"]
            for r in saved_runs
        }
        selected = st.multiselect("Select one or more saved periods to load", list(run_labels.keys()), key="amazon_load_select")
        if st.button("Load selected", disabled=(len(selected) == 0), key="amazon_load_btn"):
            fnames = [run_labels[s] for s in selected]
            combined = storage.combine_amazon_runs(client_key, fnames)
            st.session_state["amazon_order_reco_df"] = combined["order_reco_df"]
            st.session_state["amazon_expense_ledger_df"] = combined["expense_ledger_df"]
            st.session_state["amazon_settlement_summary_df"] = combined["settlement_summary_df"]
            st.session_state["amazon_settlement_register_df"] = combined["settlement_register_df"]
            st.session_state["amazon_non_mtr_df"] = combined["non_mtr_df"]
            st.session_state["amazon_tie_out_df"] = combined["tie_out_df"]
            st.session_state["amazon_subsequent_settlements_df"] = combined["subsequent_settlements_df"]
            st.session_state["amazon_waterfall_df"] = combined["waterfall_df"]
            order_count = len(combined["order_reco_df"]) if combined["order_reco_df"] is not None else 0
            st.success(f"Loaded {len(selected)} saved period(s) — {order_count:,} orders total. "
                       "Go to Dashboard to view.")

        st.divider()
        st.caption("Delete a saved period (entire dataset for that period)")
        for r in saved_runs:
            dcol1, dcol2 = st.columns([4, 1])
            dcol1.write(f"{r['month_label']} — {r['order_count']:,} orders — saved {r['saved_at'][:10]}")
            confirm_key = f"confirm_delete_amazon_{r['file']}"
            if dcol2.button("Delete", key=f"delete_btn_amazon_{r['file']}"):
                st.session_state[confirm_key] = True
            if st.session_state.get(confirm_key):
                st.warning(
                    f"Delete '{r['month_label']}'? This moves it to the Recycle Bin (kept for "
                    f"{storage.RECYCLE_BIN_RETENTION_DAYS} days, restorable any time before then)."
                )
                yc, nc = st.columns(2)
                if yc.button("Yes, delete", key=f"yes_amazon_{r['file']}"):
                    storage.soft_delete_run(client_key, r["file"], report_keys_to_scope=mtr_report_keys)
                    st.session_state[confirm_key] = False
                    st.rerun()
                if nc.button("No, cancel", key=f"no_amazon_{r['file']}"):
                    st.session_state[confirm_key] = False
                    st.rerun()

    st.divider()
    _render_recycle_bin(client_key, config, plumbing)


def render():
    st.title("Data Management")

    flash = st.session_state.get("_data_mgmt_flash")
    if flash:
        kind, message = flash
        getattr(st, kind)(message)
        st.session_state["_data_mgmt_flash"] = None

    client_key = st.session_state.get("client_key")
    if not client_key:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    config = st.session_state.get("config") or {}
    if config.get("channel_type") == "marketplace":
        _render_amazon(client_key, config)
    else:
        _render_dtc(client_key)
