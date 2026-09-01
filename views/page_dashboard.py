"""
page_dashboard.py
------------------
The home page. Simplified filters (Financial Year + Sales Channel, with a
custom date range tucked away for when it's really needed - no Month
dropdown clutter) and a fuller set of KPIs, charts, and a reconciliation
health breakdown, styled closer to a real SaaS finance dashboard.

Every filter here applies live - Streamlit reruns the page on each widget
change, so picking a date range (or a channel, or a financial year)
updates every KPI/chart on this page immediately, no "Run" button needed.
"""

import pandas as pd
import streamlit as st

from engine.summary import headline_totals, month_summary, status_summary, open_queries
from engine.formatting import indian_number
from engine.consolidator import receipt_detail_by_order
from engine.bank import classify_order_bank_status
from engine.settlement_pending import reconciliation_health_by_gateway
from views.charts import bar_grouped, donut, horizontal_bar, health_gauge
from views.filters import (
    render_dashboard_filter_controls, load_combined_with_settlement,
    current_financial_year, trim_to_date_range,
)
from views.state_init import get_client_channels
from engine import storage


def _render_amazon(client_key, config):
    """
    Amazon/marketplace-channel Dashboard - mirrors the DTC dashboard's own
    "list saved periods, filter, combine, show KPIs" shape, but sourced
    from engine.storage.list_amazon_runs/combine_amazon_runs (see
    engine/storage.py's save_amazon_run docstring for why this is a
    separate function family rather than reusing list_runs/combine_runs).
    render_dashboard_filter_controls itself needed no changes at all - it
    only ever looks at "financial_year"/"channel_name" keys, which
    list_amazon_runs deliberately returns in the same shape as list_runs.
    """
    runs = storage.list_amazon_runs(client_key)
    if not runs:
        st.info(
            "No reconciliation saved yet for this channel. Go to **Upload Data** to upload files, "
            "then **Reconciliation** to run one, then **Data Management** to save it — it'll show "
            "up here automatically."
        )
        return

    filtered_runs, (date_from, date_to) = render_dashboard_filter_controls(runs, key_prefix=f"dash_amz_{client_key}")
    if not filtered_runs:
        st.warning("No saved periods match the current filters.")
        return

    fnames = [r["file"] for r in filtered_runs]
    combined = storage.combine_amazon_runs(client_key, fnames)
    order_reco_df = combined["order_reco_df"]
    order_reco_df = trim_to_date_range(order_reco_df, "order_date", date_from, date_to)

    if order_reco_df is None or order_reco_df.empty:
        st.warning("No orders fall inside the selected date range.")
        return

    fy_set = {r.get("financial_year") for r in filtered_runs}
    fy_label = filtered_runs[0].get("financial_year", current_financial_year())
    months_included = sorted({r["month_label"] for r in filtered_runs})
    range_caption = ""
    if date_from or date_to:
        range_caption = f" · {date_from or '…'} to {date_to or '…'}"
    st.caption(
        f"{', '.join(months_included)} · {fy_label if len(fy_set) == 1 else 'Multiple FYs'} · "
        f"{len(order_reco_df):,} MTR orders{range_caption}"
    )
    st.divider()

    waterfall_df = combined["waterfall_df"]
    wf = dict(zip(waterfall_df["Particular"], waterfall_df["Amount"])) if waterfall_df is not None and not waterfall_df.empty else {}

    def _wf(label, default=0.0):
        return wf.get(label, default)

    pending = int((~order_reco_df["has_settlement_row"]).sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total MTR Orders", f"{len(order_reco_df):,}")
    c2.metric("Net Sales (MTR)", f"₹{indian_number(_wf('Net Sales'))}")
    c3.metric("Receivable", f"₹{indian_number(_wf('Receivable'))}")
    c4.metric("Balance Receivable", f"₹{indian_number(_wf('Balance Receivable'))}")

    c5, c6, c7 = st.columns(3)
    c5.metric("Order-level Deductions (incl. TDS/TCS)", f"₹{indian_number(abs(_wf('Less: Order-level Deductions (Flat File, incl. TDS/TCS)')))}")
    c6.metric("Settlement-level Deductions", f"₹{indian_number(abs(_wf('Less: Settlement-level Deductions (Flat File)')))}")
    c7.metric("Settlement Pending (orders)", f"{pending:,}")

    st.divider()
    st.markdown("##### Top-line Waterfall")
    st.caption(
        "Sales as per MTR → Net Sales → Order-level Deductions → Settlement-level Deductions → "
        "Receivable → Balance Receivable, plus two memo lines (Reserve Movement, non-MTR/MCF "
        "pass-through items) excluded from the total - see engine/amazon_reco.py for why."
    )
    if waterfall_df is not None and not waterfall_df.empty:
        st.dataframe(waterfall_df, use_container_width=True, hide_index=True)

    subsequent_df = combined["subsequent_settlements_df"]
    if subsequent_df is not None and not subsequent_df.empty:
        st.caption(
            f"{len(subsequent_df):,} settlement(s) totalling ₹{indian_number(subsequent_df['settlement_amount'].sum())} "
            "fell after the reporting cut-off and are excluded from the figures above - see "
            "**Reports** for the full Subsequent Settlements detail."
        )

    st.divider()

    settlement_register_df = combined["settlement_register_df"]
    left, right = st.columns([1, 2])
    if settlement_register_df is not None and not settlement_register_df.empty:
        status_counts = settlement_register_df["status"].value_counts()
        with left:
            match_rate = (status_counts.get("Matched", 0) / max(len(settlement_register_df), 1)) * 100
            st.markdown("##### Settlement-to-Bank Match Rate")
            health_gauge(round(match_rate, 1), title="Settlements Matched")
        with right:
            st.markdown("##### Settlement status breakdown")
            status_df = status_counts.rename_axis("status").reset_index(name="count")
            donut(status_df, "status", "count", "Settlement Status")
    else:
        st.info("Upload and match a bank statement (Upload Data / Reconciliation) to see settlement-to-bank health here.")

    st.divider()
    st.markdown("##### Orders needing attention (no matching settlement line yet)")
    pending_df = order_reco_df[~order_reco_df["has_settlement_row"]]
    if pending_df.empty:
        st.success("Every MTR order already has a matching settlement line - nothing pending.")
    else:
        cols = ["order_id", "segment", "order_date", "skus", "quantity", "mtr_status",
                "invoice_amount", "refund_amount"]
        st.dataframe(pending_df[[c for c in cols if c in pending_df.columns]],
                     use_container_width=True, height=350)

    with st.expander("📅 Period breakdown", expanded=False):
        rows = []
        for r in sorted(filtered_runs, key=lambda x: x.get("date_min") or ""):
            rows.append({
                "Period": r["month_label"],
                "Orders": r["order_count"],
                "Saved": r["saved_at"][:10],
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.caption(
        "Figures update automatically whenever you run a new period's Amazon reconciliation - "
        "no manual save step needed. Use **Data Management** if you want to re-save under a "
        "different period label, or load multiple saved periods together."
    )


def _channel_has_data(channel):
    """Whether this channel (a {"client_key", "config", ...} dict from
    get_client_channels) has any saved reconciliation at all - used to
    decide which channel tabs to show on the combined Dashboard below."""
    cfg = channel["config"]
    key = channel["client_key"]
    if (cfg or {}).get("channel_type") == "marketplace":
        return bool(storage.list_amazon_runs(key))
    return bool(storage.list_runs(key))


def _render_dtc(client_key, config):
    runs = storage.list_runs(client_key)

    if not runs:
        st.info(
            "No reconciliation saved yet for this channel. Go to **Upload Data** to upload files, "
            "then **Reconciliation** to run one — it'll show up here automatically."
        )
        return

    filtered_runs, (date_from, date_to) = render_dashboard_filter_controls(runs, key_prefix=f"dash_{client_key}")

    if not filtered_runs:
        st.warning("No saved months match the current filters.")
        return

    reco_df, lookup_df, consolidated_df, bank_ledger_df = load_combined_with_settlement(client_key, filtered_runs)

    # A saved month can span more than the custom date range picked above -
    # trim down to the exact days requested, not just whichever whole
    # month(s) happen to overlap it. Without this the KPIs never actually
    # moved when you picked a custom range.
    reco_df = trim_to_date_range(reco_df, "created_at", date_from, date_to)
    if lookup_df is not None and "Order date" in lookup_df.columns:
        lookup_df = trim_to_date_range(lookup_df, "Order date", date_from, date_to)

    if reco_df is None or len(reco_df) == 0:
        st.warning("No orders fall inside the selected date range.")
        return

    totals = headline_totals(reco_df)

    fy_set = {r.get("financial_year") for r in filtered_runs}
    fy_label = filtered_runs[0].get("financial_year", current_financial_year())
    months_included = sorted({r["month_label"] for r in filtered_runs})
    range_caption = ""
    if date_from or date_to:
        range_caption = f" · {date_from or '…'} to {date_to or '…'}"
    st.caption(
        f"{', '.join(months_included)} · {fy_label if len(fy_set) == 1 else 'Multiple FYs'} · "
        f"{len(reco_df):,} orders{range_caption}"
    )

    st.divider()

    # --- KPI cards (top row - the headline numbers) --------------------
    delivered = int((reco_df["final_delivery_status"] == "Delivered").sum())
    rto = int((reco_df["final_delivery_status"] == "RTO").sum())
    refund_pending = int(reco_df["query"].str.contains("not refunded", na=False).sum())
    n = max(len(reco_df), 1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Orders", f"{len(reco_df):,}")
    c2.metric("Delivered", f"{delivered:,}", f"{delivered/n*100:.1f}%")
    c3.metric("RTO", f"{rto:,}", f"{rto/n*100:.1f}%")
    c4.metric("Refund Pending", f"{refund_pending:,}")

    c5, c6 = st.columns(2)
    c5.metric("Gross Order Value", f"₹{indian_number(totals['Gross order value'])}")
    c6.metric("Net Settlement", f"₹{indian_number(totals['Net settlement'])}")

    st.divider()

    # --- Delivered-but-pending breakdown (partial vs fully unpaid) -----
    fully_unpaid = int((reco_df["query"] == "Delivered but amount not received - reason?").sum())
    partial_paid = int(reco_df["query"].str.startswith("Partial Payment Received", na=False).sum())
    excess_paid = int(reco_df["query"].str.startswith("Excess Payment Received", na=False).sum())
    open_q = int((reco_df["query"] != "Okk").sum())

    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Delivered, Fully Unpaid", f"{fully_unpaid:,}")
    p2.metric("Delivered, Partially Paid", f"{partial_paid:,}")
    p3.metric("Delivered, Excess Received", f"{excess_paid:,}")
    p4.metric("Open Exceptions (total)", f"{open_q:,}")

    st.divider()

    # --- Charts ----------------------------------------------------------
    m_df = month_summary(reco_df)
    s_df = status_summary(reco_df)
    q_df = open_queries(reco_df)

    left, right = st.columns([2, 1])
    with left:
        bar_grouped(m_df, "month", ["receipt", "settlement"], "Order & Settlement Overview")
    with right:
        donut(s_df, "final_delivery_status", "orders", "Order Status Distribution")

    c_left, c_right = st.columns(2)
    with c_left:
        payment_df = pd.DataFrame({
            "bucket": ["Received / Settled", "Pending Amount", "Refund Pending"],
            "amount": [
                totals["Net settlement"],
                max(totals["Gross order value"] - totals["Receipt before deduction"], 0),
                reco_df.loc[reco_df["query"].str.contains("not refunded", na=False), "receipt_amount"].sum(),
            ],
        })
        donut(payment_df, "bucket", "amount", "Payment Summary")
    with c_right:
        if len(q_df) > 0:
            horizontal_bar(q_df, "orders", "query", "Top Exceptions")
        else:
            st.success("No open exceptions — everything reconciled cleanly.")

    st.divider()

    # --- Reconciliation health ------------------------------------------
    hc1, hc2 = st.columns([1, 2])
    with hc1:
        health = round(min(100, max(0, (1 - open_q / n) * 100)), 1)
        st.markdown("##### Reconciliation Health")
        health_gauge(health)
    with hc2:
        st.markdown("##### Health breakdown")
        data_completeness = (reco_df["delivery_partner"].notna().sum() / n) * 100
        order_matching = ((reco_df["final_delivery_status"] != "Status Undefined").sum() / n) * 100
        payment_reconciliation = ((reco_df["query"] == "Okk").sum() / n) * 100

        for label, pct in [
            ("Data Completeness", data_completeness),
            ("Order Matching", order_matching),
            ("Payment Reconciliation", payment_reconciliation),
        ]:
            st.write(f"**{label}**: {pct:.1f}%")
            st.progress(min(pct, 100) / 100)

    st.divider()

    # --- Gateway-wise / delivery-partner-wise bank & settlement health ----
    bank_statement_uploaded = bank_ledger_df is not None and not bank_ledger_df.empty
    recon_status_df = classify_order_bank_status(
        reco_df, consolidated_df, bank_ledger_df, config.get("gateways", []), bank_statement_uploaded,
    )
    receipt_detail_df = receipt_detail_by_order(consolidated_df) if consolidated_df is not None else None
    gateway_health_df = reconciliation_health_by_gateway(
        reco_df, recon_status_df, receipt_detail_df, config.get("gateways", []),
    )
    if gateway_health_df is not None and not gateway_health_df.empty:
        st.markdown("##### Reconciliation health by gateway / delivery partner")
        st.caption(
            "How much is already received in bank vs pending settlement vs pending bank "
            "matching vs a genuine exception, for each payment gateway / COD delivery "
            "partner - see engine/bank.py and engine/settlement_pending.py for exactly how "
            "each bucket is decided."
        )
        st.dataframe(
            gateway_health_df.drop(columns=["Group"]) if "Group" in gateway_health_df.columns else gateway_health_df,
            use_container_width=True, hide_index=True,
        )
        st.divider()

    with st.expander("📅 Monthly breakdown", expanded=False):
        rows = []
        for r in sorted(filtered_runs, key=lambda x: x.get("date_min") or ""):
            payload_totals = None
            try:
                payload = storage.load_run(client_key, r["file"])
                payload_totals = payload.get("totals")
            except Exception:
                pass
            rows.append({
                "Month": r["month_label"],
                "Orders": r["order_count"],
                "Gross value": payload_totals.get("Gross order value") if payload_totals else None,
                "Receipt": payload_totals.get("Receipt before deduction") if payload_totals else None,
                "Settlement": payload_totals.get("Net settlement") if payload_totals else None,
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.caption(
        "Figures update automatically whenever you run a new month's reconciliation - "
        "no manual save step needed."
    )


def _render_channel(channel):
    cfg = channel["config"] or {}
    key = channel["client_key"]
    if cfg.get("channel_type") == "marketplace":
        _render_amazon(key, cfg)
    else:
        _render_dtc(key, cfg)


def render():
    st.title("Reconciliation Dashboard")
    config = st.session_state.get("config")
    client_key = st.session_state.get("client_key")

    if not client_key:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    # Show every channel belonging to this client that has ANY saved
    # reconciliation - not just whichever single channel happens to be
    # "active" right now via Settings/Upload Data. Uploading/reconciling a
    # new channel only ever changes which channel is active for the next
    # upload; it never touches another channel's already-saved data (each
    # channel is saved under its own client_key / data_store folder), so
    # the previously uploaded channel's numbers must keep showing up here
    # too, side by side, instead of dropping off the Dashboard.
    all_channels = get_client_channels((config or {}).get("client_name")) or [
        {"label": st.session_state.get("chosen_label", ""), "client_key": client_key, "config": config or {}}
    ]
    channels_with_data = [ch for ch in all_channels if _channel_has_data(ch)]

    if not channels_with_data:
        if config:
            st.caption(f"{config['client_name']} — {config['channel_name']}")
        st.info(
            "No reconciliation saved yet. Go to **Upload Data** to upload files, then run a "
            "reconciliation - it'll show up here automatically for each channel."
        )
        return

    if len(channels_with_data) == 1:
        only = channels_with_data[0]
        cfg = only["config"] or {}
        if cfg:
            st.caption(f"{cfg.get('client_name', '')} — {cfg.get('channel_name', '')}")
        _render_channel(only)
        return

    client_name = channels_with_data[0]["config"].get("client_name", "")
    st.caption(f"{client_name} — all channels with saved reconciliations")
    tabs = st.tabs([ch["config"].get("channel_name", ch["label"]) for ch in channels_with_data])
    for tab, channel in zip(tabs, channels_with_data):
        with tab:
            _render_channel(channel)
