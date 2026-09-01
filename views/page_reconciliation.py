"""
page_reconciliation.py
-----------------------
Runs the actual reconciliation engine against whatever was uploaded on the
Upload Data page. Kept separate from uploading so "did my files load ok"
and "run the numbers" are two distinct, checkable steps.
"""

import pandas as pd
import streamlit as st

from engine.consolidator import (
    build_consolidated_receipt, summarize_receipts_by_order, receipt_detail_by_order, _blank_order_id_mask,
)
from engine.reco import (
    run_shopify_pipeline, attach_settlement_pending, refine_queries_with_settlement_status,
    attach_receipt_status,
)
from engine.lookup import build_sku_detail, build_order_lookup
from engine.bank import (
    load_bank_statement, classify_order_bank_status,
    build_cod_settlement_batches, match_batches_to_bank, matched_order_level_utrs,
    build_settlement_ledger, bank_reconciliation_by_utr, build_order_level_utr_detail,
    build_refund_utr_detail, to_naive_datetime_series,
    COD_BANK_MATCHED, PREPAID_BANK_MATCHED, EXCEPTION_MANUAL_REVIEW,
)
from engine.period import classify_order_periods
from engine.attribution import build_gokwik_payment_provider_map, build_payment_gateway_lookups, attach_payment_columns
from engine.razorpay_settlement import remap_unmapped_rows
from engine.settlement import gateway_settlement_summary
from engine.settlement_pending import build_settlement_pending_report, settlement_pending_summary_by_gateway
from engine.summary import headline_totals
from engine.validation import check_mandatory_sources
from engine.formatting import indian_number
from engine.loaders import resolve_col
from engine import dedup, storage
from engine.dedup import AWB_COLUMN_ALIASES

from engine.amazon_loaders import load_mtr_reports, load_settlement_files
from engine.amazon_consolidator import build_settlement_summary, build_expense_ledger
from engine.amazon_reco import (
    build_order_reconciliation, build_waterfall, infer_cutoff_date,
    classify_settlements_by_cutoff, subsequent_settlements_summary,
)
from engine.amazon_bank import settlement_register, settlement_register_summary, split_received_by_cutoff
from engine.amazon_invoice_check import settlement_tie_out
from engine.amazon_consolidator import non_mtr_order_level_items


def render():
    st.title("Reconciliation")
    config = st.session_state.get("config")
    if not config:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    if config.get("channel_type") == "marketplace":
        _render_marketplace(config)
    else:
        _render_dtc(config)


def _update_dtc_ledger(client_key, config, orders_df, delivery_frames, gateway_frames):
    """
    Records this run's contribution to the client's ingestion ledger (see
    engine/storage.py) right after a successful save - so the NEXT upload
    of a wider/overlapping date range can tell what's already been
    reconciled. Orders use "seen" (skip-dup) ledger entries keyed on Order
    ID alone (a sales record is permanent once reconciled); delivery
    partner and gateway files use "latest" ledger entries (Order ID + AWB
    for delivery, Order ID + Transaction Type for gateways, per the
    client's own confirmed key choices) - a "latest" entry is never used to
    exclude anything, only to flag "this is an update" on the NEXT upload's
    preview (see views/page_upload.py's render_duplicate_check).
    """
    order_key_col = resolve_col(orders_df, config["orders"]["order_id_col"])
    if order_key_col:
        storage.add_seen_keys(
            client_key, "orders", set(dedup.build_row_key(orders_df, [order_key_col]))
        )

    for d_cfg in config.get("delivery_partners", []):
        df = delivery_frames.get(d_cfg["label"])
        if df is None or df.empty:
            continue
        key_cols = [c for c in [resolve_col(df, d_cfg["order_id_col"]), resolve_col(df, AWB_COLUMN_ALIASES)] if c]
        if not key_cols:
            continue
        keys = set(dedup.build_row_key(df, key_cols))
        storage.record_latest_records(client_key, f"delivery__{d_cfg['label']}", {k: True for k in keys})

    for g_cfg in config.get("gateways", []):
        df = gateway_frames.get(g_cfg["label"])
        if df is None or df.empty:
            continue
        # Same three-piece key (order id + type + amount) as the actual
        # merge key in views/page_upload.py's gateway accumulate_df call -
        # see that call site's docstring for why amount was added
        # (2026-08-23 overlapping-settlement-file money-loss fix). Kept in
        # lockstep so this ledger's "is this row an update to something
        # already recorded" signal never disagrees with what accumulate_df
        # actually treats as the same settlement event.
        key_cols = [c for c in [
            resolve_col(df, g_cfg["order_id_col"]), resolve_col(df, g_cfg.get("type_col")),
            resolve_col(df, g_cfg.get("amount_col")),
        ] if c]
        if not key_cols:
            continue
        keys = set(dedup.build_row_key(df, key_cols))
        storage.record_latest_records(client_key, f"gateway__{g_cfg['label']}", {k: True for k in keys})


def run_dtc_reconciliation(config):
    """
    Executes the full DTC/Shopify reconciliation pipeline - consolidated
    receipts, order-level reco, bank/settlement classification, the
    Gateway Settlement Report, Settlement Pending Report, UTR-level bank
    reconciliation, order lookup detail - entirely fresh from whatever is
    CURRENTLY in st.session_state (orders_df/delivery_frames/
    gateway_frames/bank_df), writes every result back to session_state,
    and auto-saves the run (split by financial year - see
    storage.financial_year_labels' docstring) exactly as the "Run
    reconciliation" button below has always done.

    Pulled out into its own standalone function (2026-08-21, client's
    "delete one raw file -> automatically recalculate affected
    reconciliation" request) so the exact same code path can be triggered
    from two different places: the button in _render_dtc below, and
    views/page_data_management.py's individual-raw-file-delete action -
    deleting e.g. just the Razorpay gateway file needs to immediately
    recompute every downstream layer (Payment Gateway Settlement Report,
    Settlement Pending Report, UTR bank reconciliation, etc.) from the
    reduced dataset, not silently leave the Reconciliation page's last
    results stale until someone happens to revisit it and click the
    button again. Having exactly one implementation of "run the DTC
    pipeline" (rather than copy-pasting this into the delete handler too)
    means the two trigger points can never quietly drift apart.

    Caller's responsibility to first check st.session_state["orders_df"]
    is not None - this function assumes there's at least an orders file to
    reconcile against (both call sites already guard on this: _render_dtc
    via its own early-return below, page_data_management via checking
    orders_df survived the deletion before calling this).

    Returns {"order_count": int, "bank_note": str, "save_note": str} -
    the pieces each call site's own success message is built from (the
    two want slightly different wording around the delete).
    """
    orders_df = st.session_state.get("orders_df")
    delivery_frames = st.session_state.get("delivery_frames") or {}
    gateway_frames = st.session_state.get("gateway_frames") or {}
    bank_df = st.session_state.get("bank_df")
    client_key = st.session_state.get("client_key")

    # Client-reported (2026-08-23), confirmed against her own live raw
    # Razorpay data: a settlement row whose Shopify order genuinely exists
    # and maps correctly (same token, order number, amount, UTR all
    # verified) was still coming through as entirely unmatched -
    # receipt_amount = 0 in Reco working, wrongly "Settlement Pending",
    # and its share of the UTR's bank credit missing from Bank Reco
    # (UTR-wise)'s "this period" total. Root cause: the Razorpay token-to-
    # order mapping (engine.razorpay_settlement.map_razorpay_settlement_
    # to_shopify) only ever runs ONCE, at upload time, against whatever
    # orders_df existed at that exact moment - a Razorpay settlement
    # export routinely covers orders that haven't been uploaded yet
    # (payout lag), and once that token fails to resolve it stays blank
    # forever, since nothing ever asks the question again. Retried here,
    # on every reconciliation run, against the CURRENT orders_df - a
    # previously-unmapped row gets a fresh chance every time, without ever
    # touching a row that's already mapped (see remap_unmapped_rows'
    # docstring). Persisted back to session_state AND to the raw-upload
    # store immediately once resolved, so the fix sticks (future runs
    # don't need to redo the lookup for this same row, and any later
    # accumulate_df merge sees the corrected order id, not a stale blank).
    razorpay_labels = [
        g_cfg["label"] for g_cfg in config.get("gateways", [])
        if g_cfg.get("raw_transform") == "razorpay_settlement"
    ]
    for label in razorpay_labels:
        current_df = gateway_frames.get(label)
        remapped_df = remap_unmapped_rows(current_df, orders_df)
        if remapped_df is not None and current_df is not None and not remapped_df.equals(current_df):
            gateway_frames[label] = remapped_df
            storage.save_raw_upload(client_key, "gateway_frames", label, remapped_df)
    st.session_state["gateway_frames"] = gateway_frames

    with st.spinner("Building consolidated receipt ledger..."):
        consolidated = build_consolidated_receipt(gateway_frames, config["gateways"])

        bank_statement_uploaded = bank_df is not None and "bank_statement" in config
        bank_ledger = None
        if bank_statement_uploaded:
            bank_ledger = load_bank_statement(bank_df, config["bank_statement"])

        # Note: receipt_summary here covers EVERY order_id that shows up
        # anywhere in the consolidated gateway receipts - including ones
        # not part of this run's own order file (e.g. a delayed
        # settlement for an order sold last period). That's needed below
        # for the UTR-level bank reconciliation, which has to see those
        # "other period" settlements too, not just this run's orders.
        receipt_summary = summarize_receipts_by_order(consolidated)

    with st.spinner("Matching orders to delivery status and receipts..."):
        reco_df = run_shopify_pipeline(orders_df, delivery_frames, receipt_summary, config)

    with st.spinner("Classifying bank / settlement reconciliation status..."):
        # Computed unconditionally - whether or not a bank statement was
        # uploaded this run - so COD orders that were never delivered
        # correctly show "no receipt expected" instead of the old
        # blanket "Not checked - no bank statement uploaded", and
        # delivered orders still get a Settlement Pending / Bank Matched
        # / Exception verdict even without a bank statement to check
        # against. See engine/bank.py's module-section docstring.
        recon_status_df = classify_order_bank_status(
            reco_df, consolidated, bank_ledger, config["gateways"], bank_statement_uploaded,
        )
        # Kept separately too (not just folded inside the classifier)
        # so the Bank Linking page can show exactly which COD
        # settlement batches got matched to which bank credit, and why.
        # Uses the same matched_order_level_utrs() exclusion set as
        # classify_order_bank_status() above, so this display-only view
        # can never show a bank credit as "matched" here AND as a direct
        # order-level UTR match elsewhere - one shared exclusion set,
        # not two that could disagree and double-count the same credit.
        cod_batches_df = build_cod_settlement_batches(consolidated, config["gateways"], bank_ledger)
        already_matched_utrs = matched_order_level_utrs(consolidated, bank_ledger)
        cod_batch_match_df = match_batches_to_bank(cod_batches_df, bank_ledger, exclude_utrs=already_matched_utrs)

    with st.spinner("Preparing order lookup inputs..."):
        sku_detail_df = build_sku_detail(delivery_frames, config["sku_source"]) if "sku_source" in config else pd.DataFrame(columns=["order_id", "skus", "quantity"])
        receipt_detail_df = receipt_detail_by_order(consolidated)

    # "This period" = whatever order(s) are in THIS run's own order
    # file (confirmed auto-detect approach - see engine/period.py).
    # Computed once here (rather than separately further down too) so
    # the Gateway Settlement Report's own period/timing split and the
    # UTR-level bank reconciliation's period split always agree on
    # exactly the same classification for every order.
    period_labels = classify_order_periods(
        receipt_summary["order_id"], reco_df[["order_id", "created_at"]], client_key,
    )
    period_by_order_id = dict(zip(receipt_summary["order_id"].astype(str), period_labels))

    with st.spinner("Building Payment Gateway Settlement Report..."):
        gateway_settlement_df = gateway_settlement_summary(
            consolidated, reco_df, config["gateways"], period_by_order_id=period_by_order_id,
        )

    with st.spinner("Building Settlement Pending Report..."):
        settlement_pending_df = build_settlement_pending_report(
            reco_df, recon_status_df, receipt_detail_df, config["gateways"],
            receipt_summary_df=receipt_summary,
        )
        # settlement_pending_summary_df moved below (client-reported
        # 2026-08-31, item 1) - it now reads reco_df's own Gateway/query/
        # receipt_status columns directly (see
        # engine/settlement_pending.py::settlement_pending_summary_by_
        # gateway()'s own docstring for why), none of which exist yet at
        # this point in the pipeline.
        # Client-reported 2026-08-27: "Net Settlement" must net out money
        # that hasn't reached the bank yet - see
        # engine/reco.py::attach_settlement_pending() and
        # engine/summary.py::headline_totals() for the full story. Must run
        # after settlement_pending_df exists (it's the source), and before
        # headline_totals(reco_df) below (the consumer).
        reco_df = attach_settlement_pending(reco_df, settlement_pending_df)

    with st.spinner("Attributing payment gateway / provider..."):
        # Gokwik is a checkout aggregator, not the rail that actually moves
        # the money - see engine/attribution.py's module docstring.
        # Computed unconditionally (NOT gated on a bank statement being
        # uploaded, unlike the UTR-level bank reco below): the Reco working
        # "Gateway" column and the Payment Method/Payment Provider columns
        # (client-reported 2026-08-27) are independent of bank matching.
        gokwik_provider_map = build_gokwik_payment_provider_map(
            st.session_state.get("attribution_frames"), config.get("attribution_sources", []),
        )
        # Client-reported 2026-08-27: this lookup's order_id half used to be
        # silently discarded (`_, utr_gateway_lookup = ...`), even though
        # engine/attribution.py's own docstring already promised a "Gateway"
        # column on Reco working - it was built but never wired in. Fixed
        # here: both halves are now used.
        order_id_gateway_lookup, utr_gateway_lookup = build_payment_gateway_lookups(consolidated, gokwik_provider_map)
        reco_df["Gateway"] = reco_df["order_id"].astype(str).map(order_id_gateway_lookup.to_dict())
        reco_df = attach_payment_columns(
            reco_df, st.session_state.get("attribution_frames"), config.get("attribution_sources", []),
            gokwik_provider_map=gokwik_provider_map,
            # 2026-08-31 (round 8): lets a COD order with no settlement
            # row yet still resolve "<courier> COD" as its Payment
            # Provider - see attach_payment_columns()'s own docstring.
            recon_status_df=recon_status_df, gateway_configs=config.get("gateways", []),
        )
        # Client-reported 2026-08-30: "Payment Gateway identification/
        # details are not coming correctly" - the underlying join logic
        # (engine.attribution.build_gokwik_payment_provider_map) was
        # verified correct in isolation against the client's own Gokwik
        # Order/Transaction Report data, so an EMPTY result here despite
        # both reports genuinely being uploaded points at a session-state
        # issue (e.g. attribution_frames not actually populated/restored
        # for this run) rather than the join itself. Surfaced visibly here
        # - rather than silently degrading to the plain "Gokwik" label, as
        # engine/attribution.py's own fallback philosophy otherwise
        # intends - specifically so this is easy to catch and report
        # rather than only showing up as a wrong figure three sheets away.
        attribution_sources_cfg = config.get("attribution_sources", [])
        if attribution_sources_cfg:
            attribution_frames_now = st.session_state.get("attribution_frames") or {}
            expected_labels = {c.get("label") for c in attribution_sources_cfg if c.get("label")}
            uploaded_labels = {
                label for label, df in attribution_frames_now.items()
                if df is not None and not df.empty
            }
            if expected_labels and not uploaded_labels:
                st.info(
                    "Payment Gateway Attribution (Gokwik Order Report / Transaction Report) hasn't been "
                    "uploaded yet this run - Gateway/Payment Provider will show the bare gateway/aggregator "
                    "name until both are uploaded on the Upload Data page (optional, but needed for the "
                    "easebuzz/payu-level detail)."
                )
            elif expected_labels - uploaded_labels:
                st.info(
                    "Payment Gateway Attribution: only "
                    + ", ".join(sorted(uploaded_labels)) + " is uploaded - "
                    + ", ".join(sorted(expected_labels - uploaded_labels))
                    + " is still missing, so Gokwik-routed orders can't be refined to their downstream "
                    "processor (e.g. easebuzz/payu) yet."
                )
            elif gokwik_provider_map is None or gokwik_provider_map.empty:
                st.warning(
                    "Payment Gateway Attribution: both Gokwik Order Report and Gokwik Transaction Report "
                    "are uploaded, but no orders could be matched between them - Payment Provider/Gateway "
                    "will show the bare \"Gokwik\" label for every Gokwik-routed order. This usually means "
                    "one of the two files' Order ID / Payment ID columns doesn't line up with the other's "
                    "(e.g. a re-exported file with renamed columns) - please double check both files were "
                    "uploaded for the SAME period and report the mismatch if it persists."
                )
            else:
                st.caption(
                    f"Payment Gateway Attribution: refined {len(gokwik_provider_map):,} Gokwik-routed "
                    "orders to their downstream processor (easebuzz/payu/...)."
                )
        # Client-reported 2026-08-30: an order flag_queries() marked "Okk"
        # purely by diff can still be sitting with the gateway/courier,
        # not yet bank-credited - see
        # engine/reco.py::refine_queries_with_settlement_status()'s own
        # docstring. Must run after recon_status_df (already computed
        # above) AND the "Gateway" column just attached, since the
        # specific pending label depends on both.
        reco_df = refine_queries_with_settlement_status(reco_df, recon_status_df)

        # Client-reported 2026-08-30 (item 2): new "receipt_status" column
        # ("Recipt Remark" on export - see engine/reco.py::
        # attach_receipt_status()'s own docstring), same recon_status_df
        # and same pipeline point as the query refinement immediately
        # above.
        reco_df = attach_receipt_status(reco_df, recon_status_df)

        # Client-reported 2026-08-31 (item 1): now built directly off
        # reco_df's own (just-finalised) Gateway/query/receipt_status
        # columns - see engine/settlement_pending.py::
        # settlement_pending_summary_by_gateway()'s own docstring - so it
        # must run after every one of those is in place.
        settlement_pending_summary_df = settlement_pending_summary_by_gateway(
            reco_df, config["gateways"],
        )

    with st.spinner("Reconciling bank statement by settlement (UTR)..."):
        utr_bank_reco_df = pd.DataFrame()
        if bank_ledger is not None and not bank_ledger.empty:
            # Reuses the same period_by_order_id computed above (before
            # the Gateway Settlement Report), so this and that report
            # never disagree about which orders belong to "this period".
            settlement_ledger = build_settlement_ledger(consolidated, period_by_order_id)
            utr_bank_reco_df = bank_reconciliation_by_utr(
                settlement_ledger, bank_ledger, consolidated_df=consolidated,
                utr_gateway_lookup=utr_gateway_lookup,
            )
            # Client's own "Bank UTR Detail" section (2026-08-27 request) -
            # traces each order straight through to the bank credit that
            # paid it. Built from the settlement_ledger/utr_bank_reco_df
            # already computed just above - never recomputed.
            # Client-reported 2026-08-30 (item 1): Refund UTR/Refund Date
            # - see engine/bank.py::build_refund_utr_detail()'s own
            # docstring for what these values represent.
            refund_utr_detail_df = build_refund_utr_detail(consolidated)
            utr_detail_df = build_order_level_utr_detail(
                settlement_ledger, utr_bank_reco_df, refund_utr_detail_df=refund_utr_detail_df,
            )
            if not utr_detail_df.empty:
                utr_detail_lookup = utr_detail_df.set_index("order_id")
                reco_df["order_id"] = reco_df["order_id"].astype(str)
                for col in ["Payment Date (Bank Date)", "Payment UTR", "Setlment Remarks",
                            "Refund UTR", "Refund Date"]:
                    reco_df[col] = reco_df["order_id"].map(utr_detail_lookup[col].to_dict())

    with st.spinner("Building order lookup detail..."):
        # Built AFTER every reco_df enrichment step above (delivery
        # partner, settlement pending, Gateway/Payment Method/Payment
        # Provider, Bank UTR Detail) - not right after run_shopify_pipeline()
        # like before - so Order Lookup shows the same final columns as the
        # Reco working export, not a stale snapshot from before this run's
        # own attribution/bank-matching steps ran (client-reported
        # 2026-08-27: new columns should be consistent everywhere reco_df
        # is shown, not just in the downloaded workbook).
        lookup_df = build_order_lookup(reco_df, sku_detail_df, receipt_detail_df, recon_status_df)

    st.session_state["reco_df"] = reco_df
    st.session_state["lookup_df"] = lookup_df
    st.session_state["consolidated"] = consolidated
    st.session_state["totals"] = headline_totals(reco_df)
    st.session_state["recon_status_df"] = recon_status_df
    st.session_state["cod_batches_df"] = cod_batches_df
    st.session_state["cod_batch_match_df"] = cod_batch_match_df
    st.session_state["gateway_settlement_df"] = gateway_settlement_df
    st.session_state["settlement_pending_df"] = settlement_pending_df
    st.session_state["settlement_pending_summary_df"] = settlement_pending_summary_df
    st.session_state["utr_bank_reco_df"] = utr_bank_reco_df
    st.session_state["bank_ledger_df"] = bank_ledger
    st.session_state["bank_statement_uploaded"] = bank_statement_uploaded

    # Auto-save under the detected month label, so the Dashboard's
    # cumulative view updates without a separate manual step. Saving
    # again under the same month label overwrites that month (re-running
    # May updates May, doesn't duplicate it).
    #
    # Saved SPLIT BY CALENDAR MONTH (see engine.storage.month_labels'
    # docstring for the full client-reported story) rather than as one
    # file per financial year - a single upload/run's orders_df is
    # CUMULATIVE across every upload ever confirmed this client (by
    # design - see engine.dedup.accumulate_df), not just this run's own
    # newly-added rows, so splitting any coarser than one calendar month
    # per saved period let an earlier month's orders get re-saved
    # wholesale into a LATER month's file every time more data was
    # uploaded on top - two "different" saved periods silently becoming
    # cumulative supersets of one another. engine.storage.combine_runs
    # (used by Dashboard/Reports to combine multiple saved periods) is a
    # plain concatenation with no cross-period dedup - it exists
    # specifically to combine genuinely DISJOINT months - so that overlap
    # meant any order saved under more than one label got double- (or
    # triple-, ...-) counted the moment more than one of those periods was
    # selected together, while whichever period a filter happened to
    # exclude looked like its own data had vanished, even though it
    # existed - just buried inside a later period's cumulative snapshot
    # instead of its own. A calendar month is always entirely inside one
    # financial year, so this split keeps the original "never mix
    # financial years in one saved period" guarantee too, not just adds
    # to it. In the normal case (one run = one calendar month, the vast
    # majority of the time), this loop runs exactly once and behaves
    # identically to the single save_run() call this replaced.
    month_per_order = storage.month_labels(reco_df["created_at"])

    # consolidated_df rows with NO resolvable order_id (engine.consolidator.
    # normalize_gateway_df's "blank order id but real money" case - e.g. a
    # Razorpay settlement row this tool's own order-linking step couldn't
    # trace to a Shopify order) can't be assigned a month via order_id
    # lookup at all - they'd never match any month_order_ids set below, no
    # matter how it's built. They still carry a real, matchable UTR and a
    # real settled amount, so they need their OWN month assignment - by
    # their own receipt_date (the gateway's own settlement/receipt date),
    # the same "fall back to the row's own date when there's no order to
    # key off of" pattern already used on the Amazon side (see
    # run_marketplace_reconciliation's ledger_month_by_order.fillna(...)
    # chain below).
    consolidated_orphan_mask = pd.Series(False, index=consolidated.index) \
        if consolidated is not None and not consolidated.empty else pd.Series([], dtype=bool)
    month_per_consolidated_blank = pd.Series([], dtype=object)
    if consolidated is not None and not consolidated.empty and "order_id" in consolidated.columns:
        # Same helper normalize_gateway_df/summarize_receipts_by_order/
        # receipt_detail_by_order already use, so "no usable order id"
        # means exactly the same thing everywhere in this engine - see
        # _blank_order_id_mask's own docstring for every shape a "blank"
        # can take (real NaN, literal "nan" text, empty string).
        consolidated_blank_order_mask = _blank_order_id_mask(consolidated["order_id"])

        # Client-reported (2026-08-23): "For FY 2025-26, I uploaded the
        # Shopify Order Report up to March 2026, but the corresponding
        # Settlement Reports are available up to May 2026 ... the tool
        # should not assume the Order Report and Settlement/Payment
        # Gateway Reports will always cover the same period." A
        # consolidated row can carry a perfectly real, NON-blank order id
        # (the settlement file genuinely names a real April/May Shopify
        # order) that still can't be assigned a month via month_order_ids
        # below, for a different reason than the blank-order case above:
        # that order simply hasn't been uploaded into orders_df AT ALL
        # yet (the order report's own date range stops short of the
        # settlement report's). Before this fix, such a row matched
        # neither `by_order` (its order isn't in ANY month's order set)
        # nor the blank-order `by_own_date` fallback (its order id isn't
        # blank) - so it silently fell out of every saved period's
        # consolidated snapshot, permanently, even though the order (and
        # its settlement) both genuinely exist and the order itself will
        # show up correctly the moment it's finally uploaded. Treated
        # identically to a blank-order-id row from here on - its own
        # receipt_date decides which month it rides along in - since from
        # this function's point of view "no order I currently know about"
        # and "no order id at all" need the exact same fallback.
        all_known_order_ids = set(reco_df["order_id"].astype(str))
        consolidated_unknown_order_mask = ~consolidated_blank_order_mask & (
            ~consolidated["order_id"].astype(str).isin(all_known_order_ids)
        )
        consolidated_orphan_mask = consolidated_blank_order_mask | consolidated_unknown_order_mask
        if consolidated_orphan_mask.any() and "receipt_date" in consolidated.columns:
            month_per_consolidated_blank = storage.month_labels(
                consolidated.loc[consolidated_orphan_mask, "receipt_date"]
            )

    # Bank ledger has no order_id (it's a plain bank statement, not
    # order-keyed) - its own month comes from ITS OWN credit date
    # (bank_date), not from any order.
    month_per_bank = pd.Series([], dtype=object)
    if bank_ledger is not None and not bank_ledger.empty and "bank_date" in bank_ledger.columns:
        month_per_bank = storage.month_labels(bank_ledger["bank_date"])

    # The set of months to save is the UNION of every month any of the
    # three sources actually has data for - not just the months orders
    # exist in. A bank credit (or an unattributable gateway settlement)
    # dated in a month with no order data yet is completely normal - COD/
    # Prepaid payouts routinely land 1-4 weeks after the order's own month,
    # so the LATEST order month reconciled so far very often has no
    # matching bank credit yet, while an EARLIER month's orders get their
    # bank credit only after this run. Iterating only over order-months
    # (as this loop originally did) meant a bank credit or unattributable
    # settlement dated in a month with no saved order-period yet was never
    # attached to ANY saved period at all, and therefore silently vanished
    # from every future combined Reports run - the Reports page can only
    # ever combine what got saved somewhere. Client-reported (2026-08-22):
    # UTRs confirmed present in both the bank statement and the payment
    # gateway report still showing "Bank statement not found for this UTR"
    # in the downloaded "Bank Reco (UTR-wise)" sheet, and those same
    # orders wrongly appearing in "Settlement Pending Detail" - both
    # explained by exactly this gap once traced against the client's real
    # saved periods (their July-dated Delhivery/Shiprocket/Razorpay bank
    # credits had no saved period to live in at all, since no order data
    # had been reconciled for July yet).
    #
    # Chronological, not alphabetical ("April" < "June" < "May" would sort
    # wrong) - by each month's own earliest date across all three sources,
    # same fix as engine.summary.month_summary's own groupby("month")
    # docstring, and the same pattern run_marketplace_reconciliation below
    # already uses for order_date ∪ deposit_date.
    #
    # to_naive_datetime_series() on EVERY source here, not just
    # pd.to_datetime() - client-reported (2026-08-22): "TypeError: agg
    # function failed [how->min,dtype->object]" the moment a month had
    # BOTH an order (reco_df["created_at"] - Shopify's own export commonly
    # carries a "+05:30"-style timezone offset) and a bank credit or
    # unattributable settlement in it (bank_ledger["bank_date"] /
    # consolidated["receipt_date"] - already timezone-STRIPPED at their own
    # source, see load_bank_statement/normalize_gateway_df). Concatenating
    # a timezone-AWARE Series with a timezone-NAIVE one can't produce one
    # unified datetime64 dtype, so pandas falls back to generic "object"
    # dtype holding a mix of aware and naive Timestamps - fine as long as
    # no single month-group ever needs to compare one of each, but the
    # instant it does (this exact case: an order and a bank credit landing
    # in the same calendar month), .min() has to compare them and raises,
    # since a timezone-aware and a timezone-naive Timestamp are never
    # comparable to each other at all. Stripping every source to naive
    # BEFORE concatenating - same fix already applied at every OTHER
    # date-parsing site in this engine (see engine.bank.to_naive_timestamp/
    # to_naive_datetime_series's own docstrings) - means there's only ever
    # one dtype in play, so this can never come up again regardless of
    # which sources land in the same month.
    earliest_parts = [to_naive_datetime_series(reco_df["created_at"]).groupby(month_per_order).min()]
    if len(month_per_consolidated_blank):
        earliest_parts.append(
            to_naive_datetime_series(consolidated.loc[consolidated_orphan_mask, "receipt_date"])
            .groupby(month_per_consolidated_blank).min()
        )
    if len(month_per_bank):
        earliest_parts.append(to_naive_datetime_series(bank_ledger["bank_date"]).groupby(month_per_bank).min())
    earliest_by_month = pd.concat(earliest_parts).groupby(level=0).min()

    all_months = sorted(
        set(month_per_order.dropna().unique())
        | set(month_per_consolidated_blank.dropna().unique())
        | set(month_per_bank.dropna().unique()),
        key=lambda m: earliest_by_month.get(m) if pd.notna(earliest_by_month.get(m)) else pd.Timestamp.max,
    )

    saved_period_labels = []
    for month_label in all_months:
        month_mask = (month_per_order == month_label).to_numpy()
        month_reco_df = reco_df[month_mask]
        month_order_ids = set(month_reco_df["order_id"].astype(str))

        # lookup_df (engine.lookup's exception-report table) keys its order
        # column "Order ID" (capitalized), not "order_id" - reco_df/
        # consolidated_df use the lowercase form. Check both so this filter
        # actually fires instead of silently falling through to "save the
        # whole unfiltered lookup_df every time", which would reintroduce
        # the exact cross-month duplication this fix exists to remove.
        month_lookup_df = lookup_df
        if lookup_df is not None and not lookup_df.empty:
            lookup_id_col = "order_id" if "order_id" in lookup_df.columns else (
                "Order ID" if "Order ID" in lookup_df.columns else None)
            if lookup_id_col is not None:
                month_lookup_df = lookup_df[lookup_df[lookup_id_col].astype(str).isin(month_order_ids)]

        month_consolidated_df = consolidated
        if consolidated is not None and not consolidated.empty and "order_id" in consolidated.columns:
            by_order = consolidated["order_id"].astype(str).isin(month_order_ids)
            # Orphan rows (blank order id, OR a real order id that isn't
            # uploaded into orders_df at all yet - see
            # consolidated_orphan_mask above) ride along on their OWN
            # receipt_date's month, never on any order's month.
            by_own_date = pd.Series(False, index=consolidated.index)
            if len(month_per_consolidated_blank):
                by_own_date.loc[month_per_consolidated_blank.index] = (month_per_consolidated_blank == month_label)
            month_consolidated_df = consolidated[by_order | by_own_date]

        month_bank_ledger_df = bank_ledger
        if bank_ledger is not None and not bank_ledger.empty and "bank_date" in bank_ledger.columns:
            month_bank_ledger_df = bank_ledger[(month_per_bank == month_label).to_numpy()]

        if month_reco_df.empty and (month_consolidated_df is None or month_consolidated_df.empty) \
                and (month_bank_ledger_df is None or month_bank_ledger_df.empty):
            continue

        month_totals = headline_totals(month_reco_df)
        storage.save_run(client_key, month_label, month_reco_df, month_lookup_df, month_totals,
                          config.get("channel_name"),
                          consolidated_df=month_consolidated_df, bank_ledger_df=month_bank_ledger_df)
        saved_period_labels.append(month_label)

    # The ingestion ledger (dedup/updated-record tracking - see
    # engine/dedup.py) is keyed on the raw uploaded files themselves,
    # not on the computed/month-sliced reco_df - "have we already seen
    # this order/AWB/transaction" doesn't care which month a row landed
    # in, so this runs once over the full original upload regardless of
    # how many monthly slices were just saved above.
    _update_dtc_ledger(client_key, config, orders_df, delivery_frames, gateway_frames)

    bank_note = " Bank statement linked." if bank_statement_uploaded else ""
    if len(saved_period_labels) > 1:
        save_note = (
            f"This upload covered more than one calendar month, so it was saved as "
            f"**{len(saved_period_labels)} separate periods** so figures never mix across "
            f"months: {', '.join(saved_period_labels)}."
        )
    elif saved_period_labels:
        save_note = f"Saved as **{saved_period_labels[0]}** — Dashboard will include this automatically."
    else:
        save_note = "Nothing to save (no orders found)."
    return {"order_count": len(reco_df), "bank_note": bank_note, "save_note": save_note}


def _render_dtc(config):
    orders_df = st.session_state.get("orders_df")
    delivery_frames = st.session_state.get("delivery_frames") or {}
    gateway_frames = st.session_state.get("gateway_frames") or {}

    if orders_df is None:
        st.warning("No orders file uploaded yet. Go to **Upload Data** first.")
        return

    missing_sources = check_mandatory_sources(config, orders_df, delivery_frames, gateway_frames)
    if missing_sources:
        st.warning(
            "Some mandatory reports are still missing: "
            f"**{', '.join(missing_sources)}**. You can still run with what you have, "
            "but results will be incomplete."
        )

    run_clicked = st.button("Run reconciliation", type="primary")

    if run_clicked:
        result = run_dtc_reconciliation(config)
        st.success(
            f"Done. {result['order_count']:,} orders reconciled.{result['bank_note']} {result['save_note']}"
        )

    if st.session_state.get("reco_df") is not None:
        st.divider()
        st.subheader("Last run summary")
        totals = st.session_state.get("totals") or headline_totals(st.session_state["reco_df"])
        items = list(totals.items())
        row1, row2 = items[:4], items[4:]
        for row in (row1, row2):
            cols = st.columns(len(row))
            for col, (k, v) in zip(cols, row):
                display_val = f"₹{indian_number(v)}" if "orders" not in k.lower() else f"{v:,}"
                col.metric(k, display_val)

        recon_status_df = st.session_state.get("recon_status_df")
        if recon_status_df is not None and not recon_status_df.empty:
            st.divider()
            st.subheader("Reconciliation Category breakdown")
            st.caption(
                "Every order, classified into one of six reconciliation categories - "
                "replaces the old \"Not checked - no bank statement uploaded\" / \"Bank "
                "receipt not identified\" labels with the actual reason. See engine/bank.py's "
                "module-section docstring for exactly how each category is decided."
            )
            counts = recon_status_df["Reconciliation Category"].value_counts()
            for cat, cnt in counts.items():
                st.write(f"**{cat}**: {cnt:,}")
            matched = int(recon_status_df["Reconciliation Category"].isin({COD_BANK_MATCHED, PREPAID_BANK_MATCHED}).sum())
            exceptions = int((recon_status_df["Reconciliation Category"] == EXCEPTION_MANUAL_REVIEW).sum())
            c1, c2, c3 = st.columns(3)
            c1.metric("Bank Matched", f"{matched:,}")
            c2.metric("Exceptions (manual review)", f"{exceptions:,}")
            c3.metric("Total orders classified", f"{len(recon_status_df):,}")

        gateway_settlement_df = st.session_state.get("gateway_settlement_df")
        if gateway_settlement_df is not None and not gateway_settlement_df.empty:
            st.divider()
            st.subheader("Payment Gateway Settlement Report")
            st.caption(
                "Settled, pending, and deducted amounts shown separately for each payment "
                "gateway. \"Pending for Settlement\" is Delivered orders whose money hasn't "
                "shown up in that gateway's settlement file yet - a best-effort attribution "
                "(see Reports for the full workbook and methodology notes)."
            )
            st.dataframe(gateway_settlement_df, use_container_width=True, hide_index=True)

        settlement_pending_summary_df = st.session_state.get("settlement_pending_summary_df")
        if settlement_pending_summary_df is not None and not settlement_pending_summary_df.empty:
            st.divider()
            st.subheader("Settlement Pending Report - gateway/partner-wise summary")
            st.caption(
                "Prepaid gateways (\"Payment Gateway Settlement Pending\") and COD partners "
                "(\"COD Settlement Pending\") shown separately. Full order-wise detail is on "
                "the Reports page download."
            )
            for group_name in ("Prepaid", "COD"):
                group_df = settlement_pending_summary_df[settlement_pending_summary_df["Group"] == group_name]
                if not group_df.empty:
                    st.markdown(f"**{'Payment Gateway Settlement Pending' if group_name == 'Prepaid' else 'COD Settlement Pending'}**")
                    st.dataframe(group_df.drop(columns=["Group"]), use_container_width=True, hide_index=True)

        utr_bank_reco_df = st.session_state.get("utr_bank_reco_df")
        if utr_bank_reco_df is not None and not utr_bank_reco_df.empty:
            st.divider()
            st.subheader("Bank Reconciliation by Settlement (UTR)")
            # "Matched" as a literal Remarks value is now only a rare
            # degenerate fallback (2026-08-23: Remarks is a composed phrase
            # string like "Same-period txn settled same period | Ties to
            # bank credit" - see engine.bank._classify_utr_remark) - a
            # cleanly-tied-out row is any Remarks ending in that suffix.
            matched = int(utr_bank_reco_df["Remarks"].str.endswith("Ties to bank credit").sum())
            needs_review = len(utr_bank_reco_df) - matched
            c1, c2 = st.columns(2)
            c1.metric("UTRs matched cleanly", f"{matched:,}")
            c2.metric("UTRs needing review", f"{needs_review:,}")
            st.caption(
                "One row per bank settlement (UTR), split by which period the underlying "
                "order(s) belong to and, where relevant, whether the bank credit posted within "
                "the selected period or after it - compared to the actual bank credit. Rows "
                "not ending \"Ties to bank credit\" are a starting point for manual review, not "
                "a final answer - see engine/bank.py's docstring for the classification rules used."
            )
            st.dataframe(utr_bank_reco_df, use_container_width=True, hide_index=True)


def _update_amazon_ledger(client_key, config, mtr_df, raw_settlement):
    """
    Amazon-channel equivalent of _update_dtc_ledger() above, run right
    after a successful storage.save_amazon_run(). MTR lines use "seen"
    (skip-dup) ledger entries keyed Order ID + SKU + Transaction Type, one
    ledger bucket per segment (B2C/B2A) - matches the per-segment "skip"
    check on the Upload Data page. Per the client's "Important Warning -
    Amazon MTR Report" note, this only prevents re-counting the exact same
    (order, sku, transaction_type) line twice on a wider re-upload; it does
    NOT collapse an order's Shipped and Refund lines into one - those have
    different transaction_type values and so get different keys, keeping
    both visible, exactly as the client asked. Settlement Flat File lines
    use "latest" entries (Order ID + Transaction Type), one ledger bucket
    per payment mode (COD/Online) - a settlement/adjustment can legitimately
    update in a later file.
    """
    if mtr_df is not None and not mtr_df.empty and "segment" in mtr_df.columns:
        for segment, seg_df in mtr_df.groupby("segment"):
            keys = set(dedup.build_row_key(seg_df, ["order_id", "sku", "mtr_transaction_type"]))
            storage.add_seen_keys(client_key, f"mtr__{segment}", keys)

    if raw_settlement is not None and not raw_settlement.empty:
        settlement_cols_cfg = config.get("settlement_columns", {})
        order_col = resolve_col(raw_settlement, settlement_cols_cfg.get("order_id_col"))
        tt_col = resolve_col(raw_settlement, settlement_cols_cfg.get("transaction_type_col"))
        key_cols = [c for c in [order_col, tt_col] if c]
        if key_cols and "_payment_mode" in raw_settlement.columns:
            for payment_mode, seg_df in raw_settlement.groupby("_payment_mode"):
                keys = set(dedup.build_row_key(seg_df, key_cols))
                storage.record_latest_records(client_key, f"settlement__{payment_mode}", {k: True for k in keys})


def run_marketplace_reconciliation(config, cutoff_override=None):
    """
    Executes the full Amazon/marketplace reconciliation pipeline - MTR
    load, Settlement Flat File load + expense ledger, cut-off split,
    order-level reconciliation, bank/settlement register matching,
    waterfall + tie-out - entirely fresh from whatever is CURRENTLY in
    st.session_state (mtr_files/settlement_files/bank_df), writes every
    result back to session_state, and auto-saves the run split by
    financial year, exactly as the "Run reconciliation" button in
    _render_marketplace below has always done.

    Same extraction, same reason, as run_dtc_reconciliation() above
    (2026-08-21 client "delete one raw file -> automatically recalculate"
    request) - so views/page_data_management.py's individual-raw-file-
    delete action (e.g. deleting one MTR segment or one Settlement Flat
    File payment mode) can trigger the exact same recompute the button
    does, rather than leaving the Reconciliation page's last results
    stale until manually re-run.

    cutoff_override: same optional manual cut-off date the page's own
    date_input widget collects - defaults to None (auto-infer from the
    MTR file's own last date, via infer_cutoff_date) when called from
    somewhere that has no such widget of its own (i.e. the delete
    handler), exactly matching what leaving that widget blank already
    does on this page.

    Returns {"order_count", "bank_note", "cutoff_note", "subsequent_note",
    "save_note"} - the pieces each call site's own success message is
    built from.
    """
    client_key = st.session_state.get("client_key")
    mtr_files = st.session_state.get("mtr_files") or {}
    settlement_files = st.session_state.get("settlement_files") or {}
    bank_df = st.session_state.get("bank_df")

    with st.spinner("Loading MTR report(s)..."):
        mtr_df = load_mtr_reports(mtr_files, config["mtr_columns"])

    cutoff_date = (
        pd.Timestamp(cutoff_override).normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
        if cutoff_override else infer_cutoff_date(mtr_df)
    )

    with st.spinner("Loading Settlement Flat File(s) and building the expense ledger..."):
        raw_settlement = load_settlement_files(settlement_files, config["settlement_columns"])
        settlement_summary_df = build_settlement_summary(raw_settlement, config["settlement_columns"])
        expense_ledger_df_full = build_expense_ledger(raw_settlement, config["settlement_columns"])

    with st.spinner("Splitting settlements by the reporting cut-off..."):
        within_ids, subsequent_ids = classify_settlements_by_cutoff(settlement_summary_df, cutoff_date)
        # Tag every line with its cut-off classification rather than
        # just silently slicing it away - the Expense Ledger detail
        # (session state / saved periods / Reports export) keeps ALL
        # lines from here on, "Within Cutoff" and "Subsequent" alike,
        # so an order whose only settlement fell after the cut-off is
        # still fully visible with its expenses intact, not missing
        # from the ledger the way has_settlement_row=False might
        # otherwise suggest. Only the WATERFALL/receivable totals below
        # get scoped down to the within-cutoff subset - see
        # engine/amazon_reco.py's build_waterfall docstring.
        # Vectorised (isin + a boolean map), NOT a Python lambda run
        # row-by-row via .map() - for a bulk upload's expense ledger
        # (easily 100k+ rows), a per-row Python function call adds up
        # fast; .isin() is a single C-level pass over the column. Part
        # of the "tool gets slow after uploading bulk data" fix - see
        # engine/storage.py's module docstring for the other (larger)
        # part of that fix.
        expense_ledger_df_full = expense_ledger_df_full.copy()
        within_mask = expense_ledger_df_full["settlement_id"].astype(str).isin(within_ids)
        expense_ledger_df_full["cutoff_status"] = within_mask.map({True: "Within Cutoff", False: "Subsequent"})
        expense_ledger_df = expense_ledger_df_full[within_mask]
        subsequent_df = subsequent_settlements_summary(settlement_summary_df, expense_ledger_df_full, subsequent_ids)

    with st.spinner("Building order-level reconciliation and waterfall..."):
        # has_settlement_row is computed from the FULL ledger (any
        # settlement, within-cutoff or subsequent) so a genuinely-
        # settled order never shows as "settlement pending" just
        # because its settlement landed after the cut-off - see
        # build_order_reconciliation's docstring.
        order_reco_df = build_order_reconciliation(mtr_df, expense_ledger_df, full_expense_ledger_df=expense_ledger_df_full)
        non_mtr_df = non_mtr_order_level_items(expense_ledger_df_full, set(mtr_df["order_id"]) if not mtr_df.empty else set())

    bank_ledger = None
    bank_statement_uploaded = bank_df is not None and "bank_statement" in config
    settlement_register_df = pd.DataFrame()
    received_to_date = 0.0
    received_in_transit = 0.0
    if bank_statement_uploaded:
        with st.spinner("Matching settlements to the bank statement..."):
            bank_ledger = load_bank_statement(bank_df, config["bank_statement"])
            # Match ALL settlements (not just "within cut-off") so the
            # Settlement Register stays complete/traceable - the cutoff
            # only affects which rows count toward the receivable total.
            full_register_df = settlement_register(settlement_summary_df, bank_ledger)
            settlement_register_df = full_register_df[
                full_register_df["settlement_id"].astype(str).isin(within_ids)
            ]
            # "Bank receipts must also be split by receipt date" - the
            # client's rule, independent of the settlement's own
            # deposit-date (see engine/amazon_bank.py's docstring).
            received_to_date, received_in_transit = split_received_by_cutoff(
                settlement_register_df, cutoff_date
            )

    with st.spinner("Building the waterfall and tie-out..."):
        subsequent_total = float(subsequent_df["settlement_amount"].sum()) if not subsequent_df.empty else 0.0
        waterfall_df = build_waterfall(
            order_reco_df, expense_ledger_df, mtr_df=mtr_df, received_to_date=received_to_date,
            received_in_transit=received_in_transit, subsequent_settlements_total=subsequent_total,
        )
        # Tie-out is a pure parsing/data-integrity check ("did we parse
        # every rupee in the flat file correctly") - deliberately run
        # against the FULL, unfiltered ledger/summary, not just the
        # within-cutoff subset, since it's not a period-end figure.
        tie_out_df = settlement_tie_out(expense_ledger_df_full, settlement_summary_df)

    st.session_state["amazon_cutoff_date"] = cutoff_date
    st.session_state["amazon_order_reco_df"] = order_reco_df
    st.session_state["amazon_waterfall_df"] = waterfall_df
    # The FULL, cutoff-tagged ledger - not the within-cutoff subset -
    # so the Expense Ledger detail (this page, Reports export, saved
    # periods) always shows every expense line for every order,
    # regardless of which side of the cut-off its settlement fell on.
    st.session_state["amazon_expense_ledger_df"] = expense_ledger_df_full
    st.session_state["amazon_settlement_summary_df"] = settlement_summary_df
    st.session_state["amazon_settlement_register_df"] = settlement_register_df
    st.session_state["amazon_tie_out_df"] = tie_out_df
    st.session_state["amazon_non_mtr_df"] = non_mtr_df
    st.session_state["amazon_subsequent_settlements_df"] = subsequent_df
    st.session_state["bank_statement_uploaded"] = bank_statement_uploaded

    # Auto-save, split by CALENDAR MONTH - the Amazon-path equivalent of
    # the DTC/Shopify fix above (see engine.storage.month_labels'
    # docstring for the full client-reported story). This used to split
    # only by financial year, which was already correct for the case that
    # comment originally targeted (a single upload straddling 1-April),
    # but mtr_files/settlement_files are CUMULATIVE across every upload
    # ever confirmed this client (by design), so a later run's data isn't
    # just this run's own new rows - it's everything uploaded so far.
    # Splitting only by financial year let an earlier month's orders get
    # re-saved wholesale into a LATER month's file every time more data
    # was uploaded on top, the same overlap bug as the DTC side (see
    # engine.storage.month_labels' docstring for the exact mechanics and
    # the client-reported symptom). A calendar month is always entirely
    # inside one financial year, so this split keeps the original "never
    # mix financial years in one saved period" guarantee too, not just
    # adds to it.
    #
    # Two independent things get their own month key here, same way the
    # financial-year version this replaces did:
    #   - Orders (order_reco_df) -> month of the order's OWN order_date.
    #     MTR determines which orders exist at all (this module's own
    #     rule), so the order's date is authoritative for which month an
    #     order's revenue/order-level deductions belong to.
    #   - Settlements (settlement_summary_df) -> month of the settlement's
    #     OWN deposit_date. A settlement is an atomic Amazon payout
    #     event; it isn't meaningful to split one settlement's total
    #     across two months.
    # Every expense-ledger line then follows whichever of those two it
    # belongs to: an order-level line (has an order_id Amazon's flat
    # file ties to a real order) follows that order's month; everything
    # else (settlement-level fees/reserves, or an order-id-shaped line
    # that isn't a known MTR order - MCF pass-through) follows its own
    # settlement's month, falling back to the line's own posted_date only
    # if even that's unavailable.
    #
    # Known edge case: if an order is sold in one month but its
    # settlement doesn't land until the next (a normal Amazon payout
    # lag near a month boundary), that order's own ledger lines follow
    # the order's month while the settlement's summary/tie-out follow
    # the settlement's month - so the settlement-parsing tie-out check
    # for that one settlement may show "needs review" in one or both
    # month buckets. That's the tie-out doing its job (flagging a
    # genuine cross-period timing item for a human to look at), not a
    # bug - it never affects the headline waterfall/receivable figures,
    # which are computed fresh per month-slice below. This is exactly
    # the same edge case the financial-year version already had, just at
    # finer granularity.
    month_order = storage.month_labels(order_reco_df["order_date"]) if not order_reco_df.empty else pd.Series([], dtype=object)
    month_settlement = (
        storage.month_labels(settlement_summary_df["deposit_date"])
        if settlement_summary_df is not None and not settlement_summary_df.empty
        else pd.Series([], dtype=object)
    )

    order_month_map = dict(zip(order_reco_df["order_id"].astype(str), month_order)) if not order_reco_df.empty else {}
    settlement_month_map = (
        dict(zip(settlement_summary_df["settlement_id"].astype(str), month_settlement))
        if settlement_summary_df is not None and not settlement_summary_df.empty
        else {}
    )
    ledger_month_by_order = expense_ledger_df_full["order_id"].astype(str).map(order_month_map)
    ledger_month_by_settlement = expense_ledger_df_full["settlement_id"].astype(str).map(settlement_month_map)
    ledger_month_by_posted = storage.month_labels(expense_ledger_df_full["posted_date"])
    ledger_month = ledger_month_by_order.fillna(ledger_month_by_settlement).fillna(ledger_month_by_posted).fillna("Unknown Month")

    # Chronological, not alphabetical ("April" < "June" < "May" would sort
    # wrong) - by each month's own earliest date across BOTH order_date
    # and deposit_date, same fix as engine.summary.month_summary's own
    # groupby("month") docstring.
    order_dates = pd.to_datetime(order_reco_df["order_date"], errors="coerce") if not order_reco_df.empty else pd.Series([], dtype="datetime64[ns]")
    settlement_dates = (
        pd.to_datetime(settlement_summary_df["deposit_date"], errors="coerce")
        if settlement_summary_df is not None and not settlement_summary_df.empty
        else pd.Series([], dtype="datetime64[ns]")
    )
    earliest_parts = []
    if len(order_dates):
        earliest_parts.append(order_dates.groupby(month_order).min())
    if len(settlement_dates):
        earliest_parts.append(settlement_dates.groupby(month_settlement).min())
    earliest_by_month = pd.concat(earliest_parts).groupby(level=0).min() if earliest_parts else pd.Series(dtype="datetime64[ns]")

    all_months = sorted(
        set(month_order.dropna().unique()) | set(month_settlement.dropna().unique()),
        key=lambda m: earliest_by_month.get(m) if pd.notna(earliest_by_month.get(m)) else pd.Timestamp.max,
    )

    saved_period_labels = []
    for month_label in all_months:
        month_order_mask = (month_order == month_label).to_numpy() if len(month_order) else []
        month_order_reco_df = order_reco_df[month_order_mask] if len(month_order) else order_reco_df.iloc[0:0]
        month_order_ids = set(month_order_reco_df["order_id"].astype(str))

        month_ledger_mask = (ledger_month == month_label).to_numpy()
        month_expense_ledger_df_full = expense_ledger_df_full[month_ledger_mask]
        month_expense_ledger_df = month_expense_ledger_df_full[month_expense_ledger_df_full["cutoff_status"] == "Within Cutoff"]

        month_settlement_mask = (month_settlement == month_label).to_numpy() if len(month_settlement) else []
        month_settlement_summary_df = settlement_summary_df[month_settlement_mask] if len(month_settlement) else settlement_summary_df.iloc[0:0]
        month_settlement_ids = set(month_settlement_summary_df["settlement_id"].astype(str))
        month_subsequent_ids = subsequent_ids & month_settlement_ids

        if month_order_reco_df.empty and month_settlement_summary_df.empty:
            continue

        month_settlement_register_df = (
            settlement_register_df[settlement_register_df["settlement_id"].astype(str).isin(month_settlement_ids)]
            if settlement_register_df is not None and not settlement_register_df.empty
            else settlement_register_df
        )
        month_received_to_date, month_received_in_transit = (0.0, 0.0)
        if bank_statement_uploaded and month_settlement_register_df is not None and not month_settlement_register_df.empty:
            month_received_to_date, month_received_in_transit = split_received_by_cutoff(month_settlement_register_df, cutoff_date)

        month_tie_out_df = settlement_tie_out(month_expense_ledger_df_full, month_settlement_summary_df)
        month_non_mtr_df = non_mtr_order_level_items(month_expense_ledger_df_full, month_order_ids)
        month_subsequent_df = subsequent_settlements_summary(month_settlement_summary_df, month_expense_ledger_df_full, month_subsequent_ids)
        month_subsequent_total = float(month_subsequent_df["settlement_amount"].sum()) if not month_subsequent_df.empty else 0.0

        month_mtr_df = mtr_df[mtr_df["order_id"].astype(str).isin(month_order_ids)] if mtr_df is not None and not mtr_df.empty else mtr_df

        month_waterfall_df = build_waterfall(
            month_order_reco_df, month_expense_ledger_df, mtr_df=month_mtr_df,
            received_to_date=month_received_to_date, received_in_transit=month_received_in_transit,
            subsequent_settlements_total=month_subsequent_total,
        )

        storage.save_amazon_run(
            client_key, month_label,
            order_reco_df=month_order_reco_df,
            waterfall_df=month_waterfall_df,
            expense_ledger_df=month_expense_ledger_df_full,
            settlement_summary_df=month_settlement_summary_df,
            settlement_register_df=month_settlement_register_df,
            non_mtr_df=month_non_mtr_df,
            tie_out_df=month_tie_out_df,
            subsequent_settlements_df=month_subsequent_df,
            cutoff_date=cutoff_date,
            channel_name=config.get("channel_name"),
        )
        saved_period_labels.append(month_label)

    _update_amazon_ledger(client_key, config, mtr_df, raw_settlement)

    bank_note = " Bank statement linked." if bank_statement_uploaded else ""
    cutoff_note = f" Reporting cut-off: {cutoff_date:%d-%b-%Y}." if cutoff_date is not None else ""
    subsequent_note = f" {len(subsequent_ids):,} subsequent settlement(s) excluded from this receivable." if subsequent_ids else ""
    if len(saved_period_labels) > 1:
        save_note = (
            f"This upload covered more than one calendar month, so it was saved as "
            f"**{len(saved_period_labels)} separate periods** so figures never mix across "
            f"months: {', '.join(saved_period_labels)}."
        )
    elif saved_period_labels:
        save_note = f"Saved as **{saved_period_labels[0]}** — Dashboard will include this automatically."
    else:
        save_note = "Nothing to save (no orders or settlements found)."
    return {
        "order_count": len(order_reco_df), "bank_note": bank_note, "cutoff_note": cutoff_note,
        "subsequent_note": subsequent_note, "save_note": save_note,
    }


def _render_marketplace(config):
    """
    Amazon (and future marketplace-channel) reconciliation run: MTR
    (revenue) + Settlement Flat File (expenses, melted into a granular
    line-by-line ledger - see engine/amazon_consolidator.py) + Bank
    Statement (settlement-to-bank matching by amount+date, reusing the
    same batch matcher built for Shopify's COD couriers).

    Reporting cut-off (client's own rule, 2026-08-19 - see
    engine/amazon_reco.py's module docstring for the exact wording): the
    MTR report and Settlement Flat File don't have to cover the same
    window - MTR determines which orders exist at all, but the flat file
    can run past MTR's own end date (a later settlement cycle that hadn't
    closed yet when MTR was pulled). Settlements dated after the cut-off
    are excluded from the receivable-as-of-cutoff calculation and instead
    shown as "Subsequent Settlements" - present, traceable, just not
    counted in this period's Balance Receivable. Bank receipts are split
    the same way, independently, by the receipt's own date.
    """
    mtr_files = st.session_state.get("mtr_files") or {}
    settlement_files = st.session_state.get("settlement_files") or {}

    if not mtr_files or not settlement_files:
        st.warning("No MTR report and/or Settlement Flat File uploaded yet. Go to **Upload Data** first.")
        return

    st.caption(
        "Reporting cut-off date - settlements and bank receipts dated after this are treated as "
        "\"subsequent\" (against this receivable, not yet due) rather than counted in the balance "
        "below. Leave blank to auto-use the MTR file's own last order/invoice date."
    )
    cutoff_override = st.date_input("Reporting cut-off date (optional)", value=None, key="amazon_cutoff_override")

    run_clicked = st.button("Run reconciliation", type="primary")

    if run_clicked:
        result = run_marketplace_reconciliation(config, cutoff_override)
        st.success(
            f"Done. {result['order_count']:,} MTR orders reconciled.{result['bank_note']}"
            f"{result['cutoff_note']}{result['subsequent_note']} {result['save_note']}"
        )

    waterfall_df = st.session_state.get("amazon_waterfall_df")
    if waterfall_df is None or waterfall_df.empty:
        return

    st.divider()
    cutoff_date = st.session_state.get("amazon_cutoff_date")
    cutoff_label = f" (as of {cutoff_date:%d-%b-%Y})" if cutoff_date is not None else ""
    st.subheader(f"Reconciliation waterfall{cutoff_label}")
    st.caption(
        "Sales as per MTR (Invoice Value) is used as the revenue base throughout, not Amazon's "
        "own Principal Amount field - it's the GST-invoice figure Amazon itself computes fees on. "
        "TDS/TCS is included in Order-level Deductions because it's real cash withheld, even "
        "though it's separately recoverable via the ITR. Settlements and bank receipts dated "
        "after the reporting cut-off are excluded from this waterfall and shown separately below "
        "as \"Subsequent Settlements\" - not lost, just not yet due against this period."
    )
    st.dataframe(waterfall_df, use_container_width=True, hide_index=True)

    subsequent_df = st.session_state.get("amazon_subsequent_settlements_df")
    if subsequent_df is not None and not subsequent_df.empty:
        st.divider()
        st.subheader("Subsequent Settlements (after the reporting cut-off)")
        st.caption(
            "Settlements Amazon reported/deposited after the reporting cut-off - excluded from "
            "the Balance Receivable above per the client's own rule (MTR determines the order "
            "population; settlements after that cut-off apply against this receivable in a later "
            "period, not this one). Fully traceable here, including the order-level and "
            "settlement-level deductions embedded in each."
        )
        st.dataframe(subsequent_df, use_container_width=True, hide_index=True)
        st.metric("Total subsequent settlement amount", f"Rs {subsequent_df['settlement_amount'].sum():,.2f}")

    tie_out_df = st.session_state.get("amazon_tie_out_df")
    if tie_out_df is not None and not tie_out_df.empty:
        st.divider()
        st.subheader("Settlement parsing tie-out")
        st.caption(
            "Every line this run extracted from the Settlement Flat File, summed per settlement, "
            "checked against Amazon's OWN reported total for that settlement. A clean tie-out means "
            "the parser captured every line correctly - nothing missed, nothing double-counted. "
            "This is NOT yet a match against Amazon's own fee/tax invoices (that report isn't "
            "available yet - see engine/amazon_invoice_check.py for what's needed to extend this)."
        )
        tied = int((tie_out_df["status"] == "Tied out").sum())
        c1, c2 = st.columns(2)
        c1.metric("Settlements tied out", f"{tied:,} / {len(tie_out_df):,}")
        c2.metric("Needs review", f"{len(tie_out_df) - tied:,}")
        if tied < len(tie_out_df):
            st.dataframe(tie_out_df[tie_out_df["status"] != "Tied out"], use_container_width=True, hide_index=True)

    settlement_register_df = st.session_state.get("amazon_settlement_register_df")
    if settlement_register_df is not None and not settlement_register_df.empty:
        st.divider()
        st.subheader("Settlement Register - bank matching")
        st.caption(
            "Each Amazon settlement matched to a bank credit by amount + deposit date (no per-order "
            "UTR exists for a settlement, exactly like the Shopify COD courier batches)."
        )
        st.dataframe(settlement_register_summary(settlement_register_df), use_container_width=True, hide_index=True)
        with st.expander("Full settlement-by-settlement detail"):
            st.dataframe(settlement_register_df, use_container_width=True, hide_index=True)
    elif not st.session_state.get("bank_statement_uploaded"):
        st.info("Upload a bank statement to see settlement-to-bank matching.")

    order_reco_df = st.session_state.get("amazon_order_reco_df")
    if order_reco_df is not None and not order_reco_df.empty:
        st.divider()
        st.subheader("Order-wise detail")
        pending = int((~order_reco_df["has_settlement_row"]).sum())
        c1, c2, c3 = st.columns(3)
        c1.metric("MTR orders", f"{len(order_reco_df):,}")
        c2.metric("Settlement pending", f"{pending:,}")
        c3.metric("Net payout (reconciled orders)", f"Rs {order_reco_df['net_payout'].sum():,.2f}")
        st.dataframe(order_reco_df, use_container_width=True, hide_index=True)

    expense_ledger_df = st.session_state.get("amazon_expense_ledger_df")
    if expense_ledger_df is not None and not expense_ledger_df.empty:
        st.divider()
        st.subheader("Expense & revenue ledger (full breakup)")
        st.caption(
            "Every fee, tax, promotion, and reserve movement Amazon's flat file lists, kept as its "
            "own line - never merged into a combined bucket - so every rupee can be traced back to "
            "the exact source column and row it came from (source_amount_col / row_ref). This "
            "includes lines from settlements dated AFTER the reporting cut-off too (tagged "
            "\"Subsequent\" below) - an order's expenses are never hidden here just because its "
            "settlement isn't counted in this period's receivable yet. Filter below, or download "
            "the full ledger (plus a compact horizontal/order-wise view) from the Reports page."
        )
        f1, f2 = st.columns(2)
        categories = sorted(expense_ledger_df["category"].dropna().unique().tolist())
        picked = f1.multiselect("Filter by category", categories, key="amazon_ledger_category_filter")
        shown = expense_ledger_df[expense_ledger_df["category"].isin(picked)] if picked else expense_ledger_df
        if "cutoff_status" in expense_ledger_df.columns:
            status_picked = f2.multiselect(
                "Filter by cut-off status", ["Within Cutoff", "Subsequent"], key="amazon_ledger_cutoff_filter",
            )
            if status_picked:
                shown = shown[shown["cutoff_status"].isin(status_picked)]
        st.dataframe(shown.head(2000), use_container_width=True, hide_index=True)
        if len(shown) > 2000:
            st.caption(f"Showing first 2,000 of {len(shown):,} matching rows - use the Reports page for the full download.")

    non_mtr_df = st.session_state.get("amazon_non_mtr_df")
    if non_mtr_df is not None and not non_mtr_df.empty:
        st.divider()
        st.subheader("Non-marketplace order-level items (MCF / Shopify pass-through)")
        st.caption(
            "Flat file lines with an order-id-shaped reference that isn't an actual Amazon "
            "marketplace order (mainly MCF - Amazon fulfilling Shopify/D2C orders and collecting "
            "COD cash on its behalf). Not part of the Amazon waterfall above - the prior year's own "
            "workbook booked this under the Shopify P&L - but shown here so it's never invisible."
        )
        st.dataframe(non_mtr_df, use_container_width=True, hide_index=True)
