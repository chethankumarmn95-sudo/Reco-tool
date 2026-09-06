"""
page_reports.py
----------------
One "Download Report" flow: pick Financial Year / Date range / Sales
Channel, then download a single workbook covering exactly that period -
no separate Month picker to also fill in. Replaces the old two-section
layout (a per-month quick-download list, plus a separate "combined"
section below it) that tied a single click to one hardcoded month name;
Month was later folded into a single filter row here, then dropped
entirely once Date range alone could already do the same job with more
precision. Every already-saved month is still visible - just on the Data
Management and Activity Log pages, not duplicated here.

Also includes, when the data is available: a Payment Gateway Settlement
Report sheet (Settlement Done / Pending for Settlement / Deductions, per
gateway - see engine/settlement.py), a UTR-level Bank Reconciliation sheet
(this-period vs other-period settlement split against the actual bank
credit - see engine/bank.py), and a Settlement Pending Report (order-wise
detail plus a gateway/partner-wise summary - see
engine/settlement_pending.py). All three are recomputed over whatever
combined, multi-month range is currently selected here, the same way Reco
working / Order Lookup already are - not just carried over from the single
most-recent run.
"""

import datetime as dt
import io
import pandas as pd
import streamlit as st

from engine.consolidator import summarize_receipts_by_order, receipt_detail_by_order
from engine.reco import (
    attach_settlement_pending, refine_queries_with_settlement_status, attach_receipt_status,
    apply_cod_report_gap_query, refine_split_payment_queries,
)
from engine.summary import month_summary, status_summary, open_queries, headline_totals
from engine.formatting import (
    style_workbook, add_dashboard_sheet, add_amazon_dashboard_sheet, strip_tz, ensure_valid_sheet_names,
    link_waterfall_formulas, write_workbook_sheets, EXCEL_SHEET_ROW_LIMIT, estimate_seconds_for_row_counts,
    reco_working_layout, style_reco_working_sections, add_executive_summary_sheet,
    apply_settlement_pending_summary_formulas,
)
from engine.lookup import build_exception_detail, build_exception_export_view
from engine.settlement import gateway_settlement_overall, gateway_recon_by_period
from engine.bank import (
    build_settlement_ledger, bank_reconciliation_by_utr, build_order_level_utr_detail,
    build_refund_utr_detail, classify_order_bank_status, to_naive_timestamp, to_naive_datetime_series,
    resolve_split_payment_leg_status,
)
from engine.attribution import (
    build_gokwik_payment_provider_map, build_payment_gateway_lookups, attach_payment_columns,
    resolve_cod_courier_label, build_direct_transaction_provider_map, combine_prepaid_and_cod_label,
    resolve_prepaid_evidence,
)
from engine.settlement_pending import (
    build_settlement_pending_report, settlement_pending_summary_by_gateway,
)
from engine.period import classify_order_periods
from engine.amazon_bank import settlement_register_summary
from engine.amazon_consolidator import pivot_expense_ledger_by_order, pivot_expense_ledger_by_settlement
from engine import storage
from views.filters import render_filter_controls, load_combined, load_combined_with_settlement, trim_to_date_range
from views.state_init import get_client_channels


def _build_workbook(reco_df, lookup_df, totals, channel_name=None, date_from=None, date_to=None,
                     gateway_settlement_overall_df=None, utr_bank_reco_df=None,
                     settlement_pending_df=None, settlement_pending_summary_df=None,
                     gateway_recon_by_period_df=None, delivery_partner_labels=None,
                     gateway_configs=None, progress_callback=None, period_bank_credit_total=None):
    # Client-reported 2026-08-27 ("professionally formatted... with clear
    # sections"): reorders/regroups reco_df's columns to match the client's
    # own reference workbook (SHOPIFY REPORT | DELIVERY PARTNER REPORT |
    # PAYMENT GATEWAY | BANK MATCHING) before it's written as "Reco
    # working" - see engine/formatting.py::reco_working_layout(). Internal-
    # only columns (e.g. settlement_pending_amount, which only feeds
    # headline totals/Executive Summary) are dropped from this view; every
    # other computation below (month_summary/status_summary/open_queries/
    # Exceptions/Dashboard) still uses the original reco_df, unaffected by
    # column order.
    reco_working_cols, reco_working_groups = reco_working_layout(reco_df, delivery_partner_labels)
    reco_working_df = reco_df[reco_working_cols].copy()
    # Client-reported 2026-08-30 (item 14, the residual bank-matching
    # difference): the client's own reference workbook shows Bank credit
    # = 0 for any order still settlement-pending (money not yet actually
    # credited to the bank), even when this engine's own settlement_amount
    # (receipt - deduction - refund) is already a real, non-zero figure -
    # see engine/reco.py::attach_settlement_pending's own docstring for
    # why that money isn't "received" yet. Root-caused against the
    # client's own July data: this engine's per-order "Bank credit"
    # column was never zeroed for ANY pending order at all (only the
    # aggregate Net Settlement figure netted it out - see headline_
    # totals()) - confirmed as the source of a ₹210,519.26 mismatch
    # across 305 orders when compared row-by-row against the client's
    # workbook, 303 of which are exactly this engine's own settlement-
    # pending set. Zeroed here, on this EXPORT-ONLY copy, using the same
    # settlement_pending_amount column attach_settlement_pending already
    # computed - not a new classification, just surfacing the existing
    # one at the per-order level the way the client's own workbook
    # already does. reco_df itself (used by every other computation
    # below - Month summary, Gateway Settlement, Dashboard, ...) is
    # deliberately untouched, since all of those already derive their own
    # figures independently rather than by re-summing this column.
    if "settlement_pending_amount" in reco_df.columns and "settlement_amount" in reco_working_cols:
        still_pending = reco_df["settlement_pending_amount"].fillna(0) > 0.01
        reco_working_df.loc[still_pending, "settlement_amount"] = 0.0
    # Client-reported 2026-08-30: display-only header rename to match the
    # client's own reference workbook's column titles ("Bank credit" /
    # "Query" instead of this engine's internal field names). Renamed on
    # THIS copy only, right before it becomes the "Reco working" sheet -
    # reco_working_cols (used below by add_executive_summary_sheet's
    # _reco_col_letter() to build cell-range formulas by ORIGINAL column
    # name/position) and reco_df itself are deliberately left untouched,
    # so nothing downstream needs to know about the rename.
    reco_working_df = reco_working_df.rename(columns={
        "settlement_amount": "Bank credit", "query": "Query", "receipt_status": "Recipt Remark",
    })

    # Client-reported 2026-08-30 (item 4): display-only header rename to
    # match the client's own reference workbook's column titles ("Query" /
    # "Recipt" / "Refund" / "Bank receipt" - MOD's own spelling/casing,
    # kept verbatim) - same pattern as the "Reco working" rename above.
    # "orders" / "order_value" / "exposure" are already lowercase in MOD's
    # own workbook, so those three stay as engine/summary.py::
    # open_queries() already names them.
    open_queries_export_df = open_queries(reco_df).rename(columns={
        "query": "Query", "receipt": "Recipt", "refund": "Refund", "bank_receipt": "Bank receipt",
    })

    sheet_data = {
        "Reco working": reco_working_df,
        "Month summary": month_summary(reco_df),
        "Status summary": status_summary(reco_df),
        "Open queries": open_queries_export_df,
    }

    # Client-reported 2026-08-30 (items 5/6): sheet names match the
    # client's own reference workbook exactly ("Gateway Settlement
    # overall" / "Gateway Recon Recoperiod") - see engine/settlement.py.
    if gateway_settlement_overall_df is not None and not gateway_settlement_overall_df.empty:
        sheet_data["Gateway Settlement overall"] = gateway_settlement_overall_df

    if gateway_recon_by_period_df is not None and not gateway_recon_by_period_df.empty:
        sheet_data["Gateway Recon Recoperiod"] = gateway_recon_by_period_df

    if utr_bank_reco_df is not None and not utr_bank_reco_df.empty:
        sheet_data["Bank Reco (UTR-wise)"] = utr_bank_reco_df

    if settlement_pending_summary_df is not None and not settlement_pending_summary_df.empty:
        sheet_data["Settlement Pending Summary"] = settlement_pending_summary_df

    if settlement_pending_df is not None and not settlement_pending_df.empty:
        sheet_data["Settlement Pending Detail"] = settlement_pending_df

    # Client-reported 2026-08-30 (item 10): the DOWNLOADED "Exceptions"
    # sheet's own column set now matches the client's own reference
    # workbook exactly (see engine/lookup.py::build_exception_export_view()
    # and EXCEPTION_EXPORT_COLUMNS for the full rationale) - a curated
    # subset of Reco working's own columns, not the richer lookup-merged
    # table the interactive Exceptions page still uses for its own filter
    # widgets. Same "Reconciliation Status != Reconciled" order set as
    # before (via build_exception_detail(), still computed here purely to
    # get that order_id list) - only the columns SHOWN in the workbook
    # sheet have changed.
    exception_detail_df = build_exception_detail(reco_df, lookup_df, channel_name)
    exception_order_ids = exception_detail_df.loc[
        exception_detail_df["Reconciliation Status"] != "Reconciled", "Order ID"
    ]
    sheet_data["Exceptions"] = build_exception_export_view(reco_df, exception_order_ids, channel_name)

    if lookup_df is not None:
        sheet_data["Order Lookup"] = lookup_df

    # Excel can't hold a timezone-aware datetime at all (a source file
    # whose own date column carries a UTC/IST offset - e.g. Shopify order
    # export timestamps - makes pandas infer a tz-aware dtype right from
    # the read, which then survives every merge/groupby all the way here)
    # - strip it from every sheet right before writing, see
    # engine/formatting.py's strip_tz docstring for why tz_localize(None)
    # rather than a UTC conversion.
    sheet_data = {name: strip_tz(df) for name, df in sheet_data.items()}
    sheet_data = ensure_valid_sheet_names(sheet_data)
    # Client-reported 2026-08-30 (item 12): add_executive_summary_sheet()'s
    # new Section 2/5 Python-side lookups need the ORIGINAL internal column
    # names (query/receipt_status/Payment Provider/...) - reco_df is about
    # to be reassigned to the renamed/zeroed EXPORT copy below (Query/
    # Recipt Remark/Bank credit), so the original is saved here first.
    original_reco_df = reco_df
    reco_df = sheet_data["Reco working"]
    generated_at = dt.datetime.now()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        overflow_csvs = write_workbook_sheets(writer, sheet_data, progress_callback=progress_callback)
        style_workbook(writer, sheet_data)

        # Client-reported 2026-08-27 ("clear sections... more attractive"):
        # add the 4-band group-header row above the Reco working sheet's
        # own column headers - see engine/formatting.py::
        # reco_working_layout()/style_reco_working_sections(). Skipped if
        # this run's Reco working was too large to embed and got replaced
        # by the overflow placeholder note sheet (see
        # write_workbook_sheets()'s own EXCEL_SHEET_ROW_LIMIT docstring) -
        # reco_working_groups would reference columns that placeholder
        # doesn't have.
        if len(reco_working_df) <= EXCEL_SHEET_ROW_LIMIT:
            style_reco_working_sections(writer.sheets["Reco working"], reco_working_groups)

        # Client-reported 2026-08-31 (round 4, point 7): link "Settlement
        # Pending Summary" to "Reco working" via live formulas - see
        # apply_settlement_pending_summary_formulas()'s own docstring.
        # Skipped automatically (no-op) if Reco working was too large to
        # embed this run (the EXCEL_SHEET_ROW_LIMIT overflow case above) or
        # this saved run predates gateway_configs being threaded through.
        if len(reco_working_df) <= EXCEL_SHEET_ROW_LIMIT:
            apply_settlement_pending_summary_formulas(writer, reco_working_cols, gateway_configs)

        key_metrics_cell_map = {}
        if totals:
            key_metrics_cell_map = add_dashboard_sheet(
                writer, totals, sheet_data["Month summary"], reco_df=reco_df,
                channel_name=channel_name, date_from=date_from, date_to=date_to,
                generated_at=generated_at,
            ) or {}

        # New "Executive Summary" sheet (client-reported 2026-08-27) - built
        # after every sheet it cross-references (Dashboard, Reco working,
        # Open queries, Settlement Pending Summary, Gateway Settlement,
        # Bank Reco (UTR-wise)) already exists in `writer`, so its formulas
        # resolve against real, already-populated ranges.
        if key_metrics_cell_map:
            add_executive_summary_sheet(
                writer, key_metrics_cell_map, reco_working_cols,
                open_queries_df=sheet_data.get("Open queries"),
                settlement_pending_summary_df=sheet_data.get("Settlement Pending Summary"),
                gateway_recon_df=sheet_data.get("Gateway Recon Recoperiod"),
                reco_df=original_reco_df,
                utr_bank_reco_df=sheet_data.get("Bank Reco (UTR-wise)"),
                has_bank_reco=bool(utr_bank_reco_df is not None and not utr_bank_reco_df.empty),
                channel_name=channel_name, date_from=date_from, date_to=date_to,
                generated_at=generated_at,
                period_bank_credit_total=period_bank_credit_total,
            )
    return buffer.getvalue(), overflow_csvs


def _report_filename(reco_df, date_from, date_to):
    """Names the file after the period it actually covers - never a single
    hardcoded month name, since a combined multi-month report shouldn't
    look like it only covers whichever month happened to be first."""
    if date_from and date_to:
        if date_from == date_to:
            return f"reconciliation_{date_from:%d%b%Y}.xlsx"
        return f"reconciliation_{date_from:%d%b%Y}_to_{date_to:%d%b%Y}.xlsx"
    if reco_df is not None and "created_at" in reco_df.columns and len(reco_df):
        dates = pd.to_datetime(reco_df["created_at"], errors="coerce").dropna()
        if len(dates):
            lo, hi = dates.min(), dates.max()
            if lo.date() == hi.date():
                return f"reconciliation_{lo:%d%b%Y}.xlsx"
            return f"reconciliation_{lo:%d%b%Y}_to_{hi:%d%b%Y}.xlsx"
    return "reconciliation_report.xlsx"


def _compute_settlement_sections(reco_df, consolidated_df, bank_ledger_df, gateway_configs, client_key,
                                  period_start_date=None, period_end_date=None,
                                  attribution_frames=None, attribution_sources_cfg=None):
    """
    Recomputes the Payment Gateway Settlement Report, the UTR-level bank
    reconciliation, the per-order reconciliation-category classification,
    and the Settlement Pending Report over the currently selected
    (combined, date-trimmed) reco_df/consolidated_df/bank_ledger_df - so a
    report covering, say, "Q1 FY26 (Apr-Jun)" gets figures for exactly
    that combined window, not just whichever single month was run most
    recently. "This period" for the period split = every order_id in the
    reco_df passed in here (see engine/period.py).

    bank_statement_uploaded is inferred as "at least one of the selected
    saved months had a bank statement" (a non-empty combined bank_ledger_df)
    - the closest available proxy once several months are combined; an
    individual month's own true/false flag isn't preserved in storage.

    period_start_date (added 2026-08-31, client-reported - Executive
        Summary Point 7 "Receivable Collection Period Analysis"): the
        start of the currently selected report period (the Reports
        page's own "From date", or the whole financial year's first day
        when left blank - see _render_dtc below) - passed straight
        through to engine.bank.bank_reconciliation_by_utr alongside
        period_end_date below, so it can finally tell a bank credit that
        genuinely landed WITHIN the selected window apart from one that
        landed before it even started (i.e. already reported in an
        earlier period's own run) - see that function's own "THE BUG
        THIS FIXES" docstring note for the full story.
    period_end_date (added 2026-08-23, client-reported spec): the end of
    the currently selected report period (the Reports page's own "To
    date", or the whole financial year's last day when that was left
    blank - see _render_dtc below) - passed straight through to
    engine.bank.bank_reconciliation_by_utr so it can tell apart a bank
    credit that posted within the selected window ("Bank Credit") from
    one that posted after it closed ("Settled Subsequent Period"). See
    that function's own docstring for the full rationale.

    attribution_frames/attribution_sources_cfg (added 2026-08-25, client-
    reported spec): the optional Gokwik Order Report / Gokwik Transaction
    Report uploads (see views/page_upload.py's "Payment Gateway
    Attribution" section) and their config entries - used purely to
    build the new "Payment Gateway" column on the Bank Reco (UTR-wise)
    sheet (see engine.attribution.build_payment_gateway_lookups). Both
    default to None/empty, in which case every Gokwik-sourced row simply
    shows the plain "Gokwik" label instead of the refined downstream
    processor - a fair fallback, never an error, exactly as
    engine/attribution.py's own docstring describes.
    """
    bank_statement_uploaded = bank_ledger_df is not None and not bank_ledger_df.empty
    receipt_summary = summarize_receipts_by_order(consolidated_df) if consolidated_df is not None else None

    # "This period" = every order_id in the reco_df passed in here (see
    # engine/period.py) - computed once up front so the Gateway Settlement
    # Report's own period/timing split and the UTR-level bank
    # reconciliation's period split always agree on the same
    # classification for every order.
    period_by_order_id = {}
    if receipt_summary is not None and len(receipt_summary):
        period_labels = classify_order_periods(
            receipt_summary["order_id"], reco_df[["order_id", "created_at"]], client_key,
        )
        period_by_order_id = dict(zip(receipt_summary["order_id"].astype(str), period_labels))

    recon_status_df = classify_order_bank_status(
        reco_df, consolidated_df, bank_ledger_df, gateway_configs, bank_statement_uploaded,
    )
    receipt_detail_df = receipt_detail_by_order(consolidated_df) if consolidated_df is not None else None
    # build_settlement_pending_report() call moved below (client-reported
    # 2026-09-06, round 18) - its "Payment Gateway" column now prefers
    # reco_df's own resolved "Payment Provider"/"Gateway" columns (see
    # engine/settlement_pending.py::build_settlement_pending_report()'s
    # own docstring for why), neither of which exist on reco_df yet at
    # this point in the pipeline - same reason settlement_pending_summary_df
    # was already moved below it (client-reported 2026-08-31, item 1).

    # Gokwik is a checkout aggregator, not the rail that actually moves the
    # money - see engine/attribution.py's module docstring. Computed
    # unconditionally (NOT gated on a bank statement being uploaded, unlike
    # the UTR-level bank reco below): the Reco working "Gateway" column and
    # the Payment Method/Payment Provider columns (client-reported
    # 2026-08-27) are independent of bank matching.
    #
    # Client-reported 2026-08-27: this lookup's order_id half used to be
    # silently discarded here (`_, utr_gateway_lookup = ...`), even though
    # engine/attribution.py's own docstring already promised a "Gateway"
    # column on Reco working - it was built but never wired in. Fixed here:
    # both halves are now used.
    gokwik_provider_map = build_gokwik_payment_provider_map(attribution_frames, attribution_sources_cfg)
    # 2026-09-04 (round 10): mirrors the same call in views/
    # page_reconciliation.py - a more direct, single-file alternative to
    # the two-report join above, used to REFINE its result - see
    # engine.attribution.build_direct_transaction_provider_map()'s own
    # docstring (order #28980).
    direct_txn_provider_map = build_direct_transaction_provider_map(attribution_frames, attribution_sources_cfg)
    order_id_gateway_lookup, utr_gateway_lookup = build_payment_gateway_lookups(consolidated_df, gokwik_provider_map)
    reco_df = reco_df.copy()
    reco_df["Gateway"] = reco_df["order_id"].astype(str).map(order_id_gateway_lookup.to_dict())
    # Client-reported 2026-09-05 (order #28980's Query text): mirrors the
    # same overlay in views/page_reconciliation.py - direct_txn_provider_
    # map already wins for "Payment Provider" below, but "Gateway" here
    # never got it, so refine_queries_with_settlement_status() (which keys
    # its Query text off "Gateway") kept the OLD, less-trustworthy
    # two-report join's answer ("easebuzz Setlment Pending" instead of
    # "payu Setlment Pending").
    if direct_txn_provider_map is not None and not direct_txn_provider_map.empty:
        direct_gateway_by_order = dict(zip(
            direct_txn_provider_map["order_id"].astype(str), direct_txn_provider_map["payment_provider"],
        ))
        direct_gateway_series = reco_df["order_id"].astype(str).map(direct_gateway_by_order)
        reco_df["Gateway"] = direct_gateway_series.where(direct_gateway_series.notna(), reco_df["Gateway"])
    # 2026-09-04 (round 9): mirrors the same override in views/
    # page_reconciliation.py - keeps "Gateway" (which refine_queries_with_
    # settlement_status() keys its Query text off) consistent with Payment
    # Provider for a confirmed-COD order - see engine/attribution.py::
    # resolve_cod_courier_label()'s docstring for the full story.
    # 2026-09-05 (order #26164): resolve_cod_courier_label() now needs
    # consolidated_df too - it resolves the courier label primarily from the
    # actual COD settlement/remittance rows (the same factual signal that
    # made this order COD in the first place), only falling back to the
    # older delivery_partner-based guess when that factual signal doesn't
    # resolve anything. See engine/attribution.py's docstring. Mirrors the
    # same call in views/page_reconciliation.py.
    cod_gateway_override = resolve_cod_courier_label(
        reco_df, recon_status_df, gateway_configs, consolidated_df=consolidated_df,
    )
    # 2026-09-04 (round 10): combine, don't overwrite, when this order also
    # already has a genuine prepaid-processor Gateway value (a real
    # part-prepaid/part-COD order) - see engine.attribution.combine_
    # prepaid_and_cod_label()'s own docstring (order #31956). Mirrors the
    # same call in views/page_reconciliation.py.
    # 2026-09-05: gated on genuine prepaid evidence (order #27533 and
    # siblings - a phantom Gokwik-attribution hit, no real money - must be
    # REPLACED by the COD label, not combined).
    # 2026-09-06 (round 16, order #30456): now resolve_prepaid_evidence() -
    # the union of a raw settlement-ledger row AND a confirmed-successful
    # Gokwik Transaction Report row - not has_genuine_prepaid_receipt()
    # alone, which missed a genuine part-prepaid/part-COD order whose
    # prepaid leg's own settlement file hadn't posted yet this run. See
    # that function's own docstring for the full story. Mirrors the same
    # call in views/page_reconciliation.py.
    genuine_prepaid_ids = resolve_prepaid_evidence(
        reco_df["order_id"], consolidated_df=consolidated_df, gateway_configs=gateway_configs,
        attribution_frames=attribution_frames, attribution_sources_cfg=attribution_sources_cfg,
    )
    reco_df["Gateway"] = combine_prepaid_and_cod_label(
        reco_df["Gateway"], cod_gateway_override,
        order_id_series=reco_df["order_id"], has_real_prepaid_receipt=genuine_prepaid_ids,
    )
    reco_df = attach_payment_columns(
        reco_df, attribution_frames, attribution_sources_cfg, gokwik_provider_map=gokwik_provider_map,
        direct_txn_provider_map=direct_txn_provider_map,
        # 2026-08-31 (round 8): lets a COD order with no settlement row
        # yet still resolve "<courier> COD" as its Payment Provider - see
        # attach_payment_columns()'s own docstring. Mirrors the same call
        # in views/page_reconciliation.py.
        recon_status_df=recon_status_df, gateway_configs=gateway_configs,
        consolidated_df=consolidated_df,
    )
    # Client-reported 2026-08-30: mirrors the same call in
    # views/page_reconciliation.py - see
    # engine/reco.py::refine_queries_with_settlement_status()'s own
    # docstring. Must run after recon_status_df (already computed above)
    # AND the "Gateway" column just attached.
    reco_df = refine_queries_with_settlement_status(reco_df, recon_status_df)

    # Client-reported 2026-09-06 (round 17): mirrors the same call in
    # views/page_reconciliation.py - see engine/reco.py::
    # refine_split_payment_queries()'s own docstring and engine/bank.py::
    # resolve_split_payment_leg_status()'s docstring for the full
    # root-cause story.
    leg_status_by_order = resolve_split_payment_leg_status(
        reco_df, consolidated_df, bank_ledger_df, gateway_configs,
    )
    reco_df = refine_split_payment_queries(reco_df, leg_status_by_order)

    # Client-reported 2026-09-04 (round 10, point 1, scenario 3): mirrors
    # the same call in views/page_reconciliation.py - see
    # engine/reco.py::apply_cod_report_gap_query()'s own docstring. Note:
    # this recompute path has no access to the raw (unfiltered) COD
    # gateway files a saved period was originally built from, so it can't
    # re-run attach_pending_cod_receipts() here the way the live
    # reconciliation run does - an order genuinely rescued into "pending
    # remittance" there (receipt_amount > 0) keeps that corrected figure
    # once the period is saved (engine.storage.save_run() persists
    # reco_df as-is), so this still resolves correctly for any period
    # saved AFTER this round's fix. A period saved BEFORE it may still
    # show this courier-named "not reflecting" text for an order that was
    # actually only pending remittance, not genuinely absent from the
    # report - re-running and re-saving that period picks up the
    # distinction, the same disclosed, accepted gap as every other
    # "wire into the core pipeline" fix in this engine.
    reco_df = apply_cod_report_gap_query(reco_df, gateway_configs)

    # Client-reported 2026-08-30 (item 2): mirrors the same call in
    # views/page_reconciliation.py - see engine/reco.py::
    # attach_receipt_status()'s own docstring.
    reco_df = attach_receipt_status(reco_df, recon_status_df)

    # Client-reported 2026-09-06 (round 18): moved to this point (after
    # attach_payment_columns()/the Gateway combine step above) so its
    # "Payment Gateway" column can prefer reco_df's own resolved "Payment
    # Provider"/"Gateway" - see that function's own docstring for the full
    # root-cause story (this sheet used to show "Gokwik"/"Unattributed
    # Prepaid" instead of the Reco working sheet's own "PayU"/"Easebuzz"/
    # "Delhivery COD"/"Shiprocket COD"). Mirrors the same move in
    # views/page_reconciliation.py.
    # period_end_date (2026-09-06, round 19, client-reported item 5): the
    # Reports page's own "To date" selector (or its fallback - see this
    # function's own docstring) - so "Days Pending"/"Report Period End
    # Date" reflect the SELECTED reconciliation period, not today's date.
    settlement_pending_df = build_settlement_pending_report(
        reco_df, recon_status_df, receipt_detail_df, gateway_configs, receipt_summary_df=receipt_summary,
        period_end_date=period_end_date,
    )

    # Client-reported 2026-08-31 (item 1): now built directly off reco_df's
    # own (just-finalised) Gateway/query/receipt_status columns - see
    # engine/settlement_pending.py::settlement_pending_summary_by_gateway()'s
    # own docstring - so it must run after every one of those is in place.
    settlement_pending_summary_df = settlement_pending_summary_by_gateway(
        reco_df, gateway_configs, consolidated_df=consolidated_df, bank_ledger_df=bank_ledger_df,
    )

    utr_bank_reco_df = pd.DataFrame()
    if consolidated_df is not None and not consolidated_df.empty and bank_ledger_df is not None and not bank_ledger_df.empty:
        settlement_ledger = build_settlement_ledger(consolidated_df, period_by_order_id)
        utr_bank_reco_df = bank_reconciliation_by_utr(
            settlement_ledger, bank_ledger_df, consolidated_df=consolidated_df,
            period_start_date=period_start_date, period_end_date=period_end_date,
            utr_gateway_lookup=utr_gateway_lookup,
        )
        # Client's own "Bank UTR Detail" section (2026-08-27 request) -
        # traces each order straight through to the bank credit that paid
        # it. Built from settlement_ledger/utr_bank_reco_df already
        # computed just above - never recomputed.
        # Client-reported 2026-08-30 (item 1): Refund UTR/Refund Date -
        # see engine/bank.py::build_refund_utr_detail()'s own docstring
        # for what these values represent.
        refund_utr_detail_df = build_refund_utr_detail(consolidated_df)
        utr_detail_df = build_order_level_utr_detail(
            settlement_ledger, utr_bank_reco_df, refund_utr_detail_df=refund_utr_detail_df,
        )
        if not utr_detail_df.empty:
            utr_detail_lookup = utr_detail_df.set_index("order_id")
            reco_df["order_id"] = reco_df["order_id"].astype(str)
            for col in ["Payment Date (Bank Date)", "Payment UTR", "Setlment Remarks",
                        "Refund UTR", "Refund Date"]:
                reco_df[col] = reco_df["order_id"].map(utr_detail_lookup[col].to_dict())

    # "Gateway Settlement overall" / "Gateway Recon Recoperiod" (client-
    # reported 2026-08-30, items 5/6) - replace the old single "Gateway
    # Settlement" sheet and the old "Gateway Reconciliation Health" sheet
    # with the client's own reference workbook's two-sheet design (see
    # engine/settlement.py's own module-level comment and each function's
    # docstring for the full rationale). Computed here, after utr_bank_
    # reco_df/reco_df's Payment Provider/query/receipt_status columns are
    # all already in place, since both new sheets are built from those.
    gateway_settlement_overall_df = gateway_settlement_overall(utr_bank_reco_df, reco_df)
    gateway_recon_by_period_df = gateway_recon_by_period(reco_df, gateway_settlement_overall_df, gateway_configs)

    return (reco_df, gateway_settlement_overall_df, utr_bank_reco_df, recon_status_df, settlement_pending_df,
            settlement_pending_summary_df, gateway_recon_by_period_df)


def _build_amazon_workbook(order_reco_df, waterfall_df, expense_ledger_df, settlement_summary_df,
                            settlement_register_df, non_mtr_df, tie_out_df, subsequent_settlements_df=None,
                            channel_name=None, date_from=None, date_to=None, cutoff_date=None,
                            progress_callback=None):
    """
    Amazon-channel equivalent of _build_workbook() above - one workbook,
    one sheet per output this engine produces (see engine/amazon_reco.py,
    engine/amazon_consolidator.py, engine/amazon_bank.py,
    engine/amazon_invoice_check.py), plus a front-page Dashboard sheet
    (KPIs + charts, mirroring the DTC workbook's own Dashboard sheet - see
    engine/formatting.py's add_amazon_dashboard_sheet).

    Expense Ledger is exported THREE ways rather than one giant row-per-
    line-item sheet (which runs into the lakhs of rows for a full year -
    see engine/amazon_consolidator.py's pivot_expense_ledger_by_order
    docstring):
      "Expense Ledger (Order-wise)"      - horizontal/wide, one row per
                                            order, one column per category
                                            - the primary, compact view.
      "Expense Ledger (Settlement)"      - same idea for the settlement-
                                            level (non-order) lines. Named
                                            "(Settlement)" rather than the
                                            more consistent "(Settlement-
                                            wise)" purely because the latter
                                            is 32 characters - one over
                                            Excel's hard 31-char sheet name
                                            limit (see below).
      "Expense Ledger (Full Detail)"     - the original long/row-per-line
                                            format, kept in full so every
                                            figure can still be traced back
                                            to its exact source row - never
                                            dropped, just no longer the
                                            first/only view.
    Every expense line is included in these regardless of has_settlement_row
    or reporting cut-off classification (a "cutoff_status" column marks
    each line Within Cutoff vs Subsequent when the cut-off feature tagged
    it - see views/page_reconciliation.py) - the client's own requirement
    that expenses are never hidden from this sheet just because an order
    isn't (yet) counted in the current period's receivable.
    """
    sheet_data = {
        "Waterfall": waterfall_df if waterfall_df is not None else pd.DataFrame(),
        "Order-wise Detail": order_reco_df if order_reco_df is not None else pd.DataFrame(),
    }

    if expense_ledger_df is not None and not expense_ledger_df.empty:
        order_wise = pivot_expense_ledger_by_order(expense_ledger_df)
        if not order_wise.empty:
            sheet_data["Expense Ledger (Order-wise)"] = order_wise
        settlement_wise = pivot_expense_ledger_by_settlement(expense_ledger_df)
        if not settlement_wise.empty:
            # Excel hard-caps sheet names at 31 characters - "Expense
            # Ledger (Settlement-wise)" is 32 and openpyxl only warns
            # (doesn't truncate or error), so the file saves "successfully"
            # but real Excel then treats that sheet name as invalid content
            # and throws up a "we found a problem, repair this file?"
            # prompt on open. Keep every sheet name here at or under 31.
            sheet_data["Expense Ledger (Settlement)"] = settlement_wise
        sheet_data["Expense Ledger (Full Detail)"] = expense_ledger_df

    if settlement_summary_df is not None and not settlement_summary_df.empty:
        sheet_data["Settlement Summary"] = settlement_summary_df

    if settlement_register_df is not None and not settlement_register_df.empty:
        sheet_data["Settlement Register"] = settlement_register_df
        sheet_data["Settlement Register Summary"] = settlement_register_summary(settlement_register_df)

    if tie_out_df is not None and not tie_out_df.empty:
        sheet_data["Settlement Tie-Out"] = tie_out_df

    if non_mtr_df is not None and not non_mtr_df.empty:
        sheet_data["Non-MTR Items (MCF)"] = non_mtr_df

    if subsequent_settlements_df is not None and not subsequent_settlements_df.empty:
        sheet_data["Subsequent Settlements"] = subsequent_settlements_df

    # Same tz-aware-datetime guard as _build_workbook() above - Amazon's
    # own order_date/settlement dates can carry a UTC/IST offset just as
    # easily as a DTC source file's can.
    sheet_data = {name: strip_tz(df) for name, df in sheet_data.items()}
    sheet_data = ensure_valid_sheet_names(sheet_data)
    waterfall_df = strip_tz(waterfall_df)
    order_reco_df = strip_tz(order_reco_df)
    settlement_register_df = strip_tz(settlement_register_df)
    subsequent_settlements_df = strip_tz(subsequent_settlements_df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        overflow_csvs = write_workbook_sheets(writer, sheet_data, progress_callback=progress_callback)
        style_workbook(writer, sheet_data)
        # Turns the Waterfall sheet's 4 roll-forward totals into real
        # formulas (see link_waterfall_formulas' docstring for why only
        # those 4 and not every line), and adds the "How This Is
        # Calculated" column explaining the rest - the end-user-facing
        # request that these figures be traceable, not just static
        # numbers.
        if "Waterfall" in writer.sheets and waterfall_df is not None and not waterfall_df.empty:
            link_waterfall_formulas(writer.sheets["Waterfall"], waterfall_df)
        add_amazon_dashboard_sheet(
            writer, waterfall_df, order_reco_df=order_reco_df, settlement_register_df=settlement_register_df,
            subsequent_settlements_df=subsequent_settlements_df, channel_name=channel_name,
            date_from=date_from, date_to=date_to, generated_at=dt.datetime.now(), cutoff_date=cutoff_date,
        )
    return buffer.getvalue(), overflow_csvs


def _amazon_report_filename(order_reco_df, date_from, date_to):
    """Names the file after the period it covers - same convention as
    _report_filename() above, keyed off order_reco_df's "order_date"
    column instead of reco_df's "created_at"."""
    if date_from and date_to:
        if date_from == date_to:
            return f"amazon_reconciliation_{date_from:%d%b%Y}.xlsx"
        return f"amazon_reconciliation_{date_from:%d%b%Y}_to_{date_to:%d%b%Y}.xlsx"
    if order_reco_df is not None and "order_date" in order_reco_df.columns and len(order_reco_df):
        dates = pd.to_datetime(order_reco_df["order_date"], errors="coerce").dropna()
        if len(dates):
            lo, hi = dates.min(), dates.max()
            if lo.date() == hi.date():
                return f"amazon_reconciliation_{lo:%d%b%Y}.xlsx"
            return f"amazon_reconciliation_{lo:%d%b%Y}_to_{hi:%d%b%Y}.xlsx"
    return "amazon_reconciliation_report.xlsx"


def _render_amazon(client_key, config):
    runs = storage.list_amazon_runs(client_key)
    if not runs:
        st.info(
            "No reconciliation run yet. Go to **Upload Data**, then **Reconciliation** and click "
            "Run - it saves automatically and will show up here."
        )
        return

    st.subheader("Download Report")
    st.caption(
        "Pick a Financial Year and, optionally, a date period below - leave the dates blank to "
        "cover the whole year - then download one workbook covering exactly that."
    )

    filtered_runs, (date_from, date_to) = render_filter_controls(
        runs, key_prefix=f"reports_amz_{client_key}", show_date_range=True, show_month=False
    )
    if not filtered_runs:
        st.warning("No saved periods match the current filters.")
        return

    fnames = [r["file"] for r in filtered_runs]
    combined = storage.combine_amazon_runs(client_key, fnames)
    order_reco_df = trim_to_date_range(combined["order_reco_df"], "order_date", date_from, date_to)

    if order_reco_df is None or order_reco_df.empty:
        st.warning("No orders fall inside the selected filters.")
        return

    range_note = ""
    if date_from or date_to:
        range_note = f" · Date range applied: {date_from or '…'} to {date_to or '…'}"
    st.caption(f"{len(order_reco_df):,} MTR orders in the current filtered selection.{range_note}")

    waterfall_df = combined["waterfall_df"]
    if waterfall_df is not None and not waterfall_df.empty:
        st.divider()
        st.subheader("Top-line Waterfall")
        st.dataframe(waterfall_df, use_container_width=True, hide_index=True)

    settlement_register_df = combined["settlement_register_df"]
    if settlement_register_df is not None and not settlement_register_df.empty:
        st.divider()
        st.subheader("Settlement Register Summary")
        st.caption(
            "Settlement Total vs Bank Receipts vs Variance, per payment mode - see "
            "engine/amazon_bank.py for the amount+date matching rules used."
        )
        st.dataframe(settlement_register_summary(settlement_register_df), use_container_width=True, hide_index=True)
    else:
        st.caption(
            "No settlement-to-bank matching available for this selection (no bank statement was "
            "uploaded/matched for the selected saved period(s))."
        )

    subsequent_settlements_df = combined["subsequent_settlements_df"]
    if subsequent_settlements_df is not None and not subsequent_settlements_df.empty:
        st.divider()
        st.subheader("Subsequent Settlements (after the reporting cut-off)")
        st.caption(
            "Settlements dated after each saved period's reporting cut-off - excluded from that "
            "period's Balance Receivable per the client's rule (MTR determines the order "
            "population; later settlements apply against the receivable in a subsequent period)."
        )
        st.dataframe(subsequent_settlements_df, use_container_width=True, hide_index=True)

    st.divider()

    # Namespaced by client_key so preparing a download for one channel
    # can never clobber or be confused with another channel's cached
    # report bytes when both channels' Reports tabs render in the same
    # session (see page_reports.py's render() - multi-channel tabs).
    bytes_key = f"amazon_report_bytes__{client_key}"
    filename_key = f"amazon_report_filename__{client_key}"
    fingerprint_key = f"amazon_last_prepared_fingerprint__{client_key}"
    overflow_key = f"amazon_report_overflow_csvs__{client_key}"

    current_fingerprint = (tuple(sorted(r["file"] for r in filtered_runs)), str(date_from), str(date_to))
    if (st.session_state.get(bytes_key) is not None
            and st.session_state.get(fingerprint_key) != current_fingerprint):
        st.info("Filters changed since you last prepared a report - click below to refresh the download.")

    if st.button("Prepare download", type="primary", key=f"amazon_prepare_download_{client_key}"):
        ledger_rows = len(combined["expense_ledger_df"]) if combined["expense_ledger_df"] is not None else 0
        filename = _amazon_report_filename(order_reco_df, date_from, date_to)
        data, overflow_csvs = _run_build_with_progress(
            _build_amazon_workbook,
            dict(
                order_reco_df=order_reco_df, waterfall_df=waterfall_df,
                expense_ledger_df=combined["expense_ledger_df"], settlement_summary_df=combined["settlement_summary_df"],
                settlement_register_df=settlement_register_df, non_mtr_df=combined["non_mtr_df"],
                tie_out_df=combined["tie_out_df"], subsequent_settlements_df=subsequent_settlements_df,
                channel_name=config.get("channel_name"), date_from=date_from, date_to=date_to,
                cutoff_date=combined.get("cutoff_date"),
            ),
            estimate_row_counts=[len(order_reco_df), ledger_rows],
        )
        st.session_state[bytes_key] = data
        st.session_state[filename_key] = filename
        st.session_state[fingerprint_key] = current_fingerprint
        st.session_state[overflow_key] = overflow_csvs

    ready = (
        st.session_state.get(bytes_key) is not None
        and st.session_state.get(fingerprint_key) == current_fingerprint
    )
    if ready:
        st.download_button(
            f"Download {st.session_state[filename_key]}",
            data=st.session_state[bytes_key],
            file_name=st.session_state[filename_key],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"amazon_download_btn_{client_key}",
        )
        _render_overflow_downloads(st.session_state.get(overflow_key), client_key, prefix="amazon")


def _format_time_estimate(seconds):
    """
    Turns a raw seconds estimate into the kind of plain-language message
    the client explicitly asked for ("Report generation in progress…
    This may take approximately 2-3 minutes. Please wait.") instead of a
    bare spinner with no sense of whether the app is working or stuck.

    Client-reported (2026-08-21): estimate_seconds_for_row_counts() only
    models the raw Excel/CSV write cost, not the full pipeline around it
    (reading the uploaded files, running the whole reconciliation, etc.) -
    for a modest dataset that raw write cost alone can come out to just a
    few/a dozen seconds, but real end-to-end report generation reliably
    takes on the order of a minute regardless, and longer for large
    datasets. Promising "13 seconds" (or "a few seconds") was actively
    misleading, not just imprecise - so any estimate under a minute is
    floored to the same "approximately one minute" message rather than
    surfacing that unreliable low figure. The minute-range scaling above
    60 seconds is unaffected - only the sub-minute buckets were wrong.
    """
    if seconds < 60:
        return "This may take approximately one minute. Please wait."
    minutes = seconds / 60.0
    lo = max(1, int(minutes))
    hi = lo + 1
    return f"This may take approximately {lo}-{hi} minutes. Please wait."


def _run_build_with_progress(build_fn, build_kwargs, estimate_row_counts):
    """
    Shared "show real, honest progress while the workbook is built" UI for
    both the DTC and Amazon "Prepare download" buttons - the client's
    explicit ask: a WhatsApp-upload-style indicator plus an upfront time
    estimate, instead of a plain spinner that gives no sense of whether
    the app is working or has hung.

    Two layers, both real (nothing here is a fake/decorative animation):
      1. An upfront estimated-time message, from estimate_seconds_for_
         row_counts() using the same per-row throughput this write path
         was benchmarked at (see engine/formatting.py).
      2. A progress bar + status line driven by write_workbook_sheets()'s
         own progress_callback - it reports BEFORE each sheet is written,
         so the bar/status genuinely reflects how much of the real work
         is done, not a countdown timer that could finish "early" while
         the app is still actually writing a huge sheet.

    build_fn is _build_workbook or _build_amazon_workbook; build_kwargs
    are its other arguments (progress_callback is injected here).
    """
    est_seconds = estimate_seconds_for_row_counts(estimate_row_counts)
    time_note = _format_time_estimate(est_seconds)

    progress_bar = st.progress(0.0)
    status = st.empty()
    status.info(f"⏳ Report generation in progress… {time_note}")

    def _on_sheet_progress(i, total, sheet_name, n_rows):
        pct = i / max(total, 1)
        progress_bar.progress(min(max(pct, 0.02), 0.98))
        status.info(
            f"⏳ Generating your report — preparing **{sheet_name}** "
            f"({i + 1} of {total} sheets, {n_rows:,} rows)… {time_note}"
        )

    try:
        result = build_fn(**build_kwargs, progress_callback=_on_sheet_progress)
    except Exception:
        progress_bar.empty()
        status.empty()
        raise
    progress_bar.progress(1.0)
    status.success("✅ Report generated successfully — your download is ready below.")
    return result


def _render_overflow_downloads(overflow_csvs, client_key, prefix):
    """
    Renders one extra download button per sheet that write_workbook_sheets()
    (engine/formatting.py) excluded from the main .xlsx for being too large
    to embed quickly (see EXCEL_SHEET_ROW_LIMIT's docstring) - most
    commonly the Amazon Expense Ledger's "Full Detail" line-item sheet for
    a full financial year's Settlement Flat File. Nothing here is a
    fallback/degraded view: it's the exact same full-fidelity data that
    would have been the Excel sheet, just delivered as .csv so the main
    workbook itself stays fast to generate and download.
    """
    if not overflow_csvs:
        return
    st.divider()
    st.caption(
        "One or more sheets in this report were too large to include directly in the Excel "
        "file without slowing the whole download down - see the note left in their place "
        "inside the workbook. Nothing is missing: the full data for each is available below."
    )
    for sheet_name, csv_bytes in overflow_csvs.items():
        safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in sheet_name).strip().replace(" ", "_")
        st.download_button(
            f"Download full \"{sheet_name}\" detail (CSV)",
            data=csv_bytes,
            file_name=f"{safe_name}.csv",
            mime="text/csv",
            key=f"{prefix}_overflow_download_{safe_name}_{client_key}",
        )


def _channel_has_data(channel):
    cfg = channel["config"] or {}
    key = channel["client_key"]
    if cfg.get("channel_type") == "marketplace":
        return bool(storage.list_amazon_runs(key))
    return bool(storage.list_runs(key))


def _settlement_runs_for_bank_matching(all_runs, filtered_runs):
    """
    Client-reported (2026-08-23): downloading a specific date-range report
    (e.g. "01-Apr-2026 to 30-Jun-2026") showed "Bank statement not found
    for this UTR" in the Bank Reco (UTR-wise) sheet, and the same orders
    wrongly still marked pending in Settlement Pending Detail, for UTRs
    that ARE genuinely in the bank statement and DO correctly tie out to
    this period's own gateway settlement - confirmed against the client's
    own corrected workbook, where every one of these UTRs' actual bank
    credit was dated in July, a month after the selected Apr-Jun window
    closed. That's completely normal payout lag - the exact same lag this
    engine already handles the OTHER direction via the "Settled with
    other-period transaction(s)" Remarks category (a bank credit batching
    together settlements from more than one period) - but here the July
    saved period's bank_ledger_df was never even loaded in the first
    place: views/filters.py's _apply_date_range (used to build
    filtered_runs below) excludes any saved period whose data falls
    ENTIRELY outside the picked date range, which is correct for scoping
    WHICH ORDERS this report is about, but wrong for scoping which bank
    credits are eligible to settle them - an order genuinely placed in
    April still needs its July-dated bank credit to be visible for its
    UTR to tie out, even though July itself is outside the report's own
    window.

    Returns filtered_runs' own saved periods PLUS every OTHER saved period
    for the same channel, REGARDLESS of financial year (revised
    2026-08-23, client-reported spec - see below for why the original
    same-FY-only version was itself a bug) - so consolidated_df/
    bank_ledger_df sourced from this broader set can always find a
    within-period settlement's actual bank credit, however late it
    posted, and can always find a previous-period order's own settlement
    data, however early it was recorded. reco_df/lookup_df (the report's
    own order population, and therefore what counts as "This period" in
    the UTR/settlement period split) are UNCHANGED by this - the caller
    still builds those from filtered_runs alone, which DOES stay
    confined to one financial year (views/filters.py's
    render_filter_controls has no "All financial years" option, by the
    client's own explicit, separate requirement - see that module's
    docstring); this broader set is used ONLY to source consolidated_df/
    bank_ledger_df for the settlement-matching layer, never to add
    another year's orders into the report itself. This mirrors exactly
    what the live Reconciliation page already does today (see
    views/page_reconciliation.py's run_dtc_reconciliation) - it always
    matches against the FULL currently-known consolidated/bank data,
    never a date-range- or financial-year-truncated slice of it; this
    just brings the Reports page's recompute in line with that same
    behaviour.

    Why NOT restrict to the same financial year (as this function
    originally did): this engine's own financial year is the Indian
    1-April-31-March convention (engine/storage.py's
    financial_year_label) - which means the single most common "previous
    period" case the client's own spec describes ("Order/transaction
    date: March; Report selected: April") ALWAYS crosses a financial-year
    boundary, every single year, by construction. The original same-FY
    restriction therefore silently dropped exactly the row the existing
    "Settled (other period)" column exists to show, the moment the
    previous period fell in March - confirmed by direct reproduction
    (a March-dated order settled in April vanished entirely from the
    Apr-May "Bank Reco (UTR-wise)" sheet, rather than merely being
    misclassified) before this fix. A payout genuinely never jumping
    financial years was true as a statement about the BANK CREDIT DATE
    relative to the ORDER DATE within a single settlement lag - it was
    never true as a statement about which financial-year FOLDER the
    matching data needs to be read from, since the order and its
    settlement can each legitimately fall in different financial years
    even when the lag between them is perfectly normal.
    """
    if not filtered_runs:
        return filtered_runs
    channels = {r.get("channel_name") for r in filtered_runs if r.get("channel_name")}
    already = {r["file"] for r in filtered_runs}
    extra = [
        r for r in all_runs
        if r["file"] not in already
        and (not channels or r.get("channel_name") in channels)
    ]
    return filtered_runs + extra


def _render_dtc(client_key, config):
    runs = storage.list_runs(client_key)
    if not runs:
        st.info(
            "No reconciliation saved yet. Go to **Upload Data** / **Reconciliation** first, "
            "then save it from **Data Management** to see it here."
        )
        return

    st.subheader("Download Report")
    st.caption(
        "Pick a Financial Year and, optionally, a date period below - leave the dates "
        "blank to cover the whole year - then download one workbook covering exactly "
        "that. (Looking for a list of what's already saved? That's on Data Management "
        "and Activity Log.)"
    )

    filtered_runs, (date_from, date_to) = render_filter_controls(
        runs, key_prefix=f"reports_{client_key}", show_date_range=True, show_month=False
    )

    if not filtered_runs:
        st.warning("No saved months match the current filters.")
        return

    reco_df, lookup_df = load_combined(client_key, filtered_runs)
    reco_df = trim_to_date_range(reco_df, "created_at", date_from, date_to)
    if lookup_df is not None and "Order date" in lookup_df.columns:
        lookup_df = trim_to_date_range(lookup_df, "Order date", date_from, date_to)

    # Bank/consolidated data sourced from a BROADER set of saved periods
    # than reco_df/lookup_df above - see _settlement_runs_for_bank_matching's
    # docstring for why (2026-08-23 late-bank-credit fix).
    settlement_runs = _settlement_runs_for_bank_matching(runs, filtered_runs)
    _, _, consolidated_df, bank_ledger_df = load_combined_with_settlement(client_key, settlement_runs)

    if reco_df is None or len(reco_df) == 0:
        st.warning("No orders fall inside the selected filters.")
        return

    range_note = ""
    if date_from or date_to:
        range_note = f" · Date range applied: {date_from or '…'} to {date_to or '…'}"
    st.caption(f"{len(reco_df):,} orders in the current filtered selection.{range_note}")

    # The report's own period end, for engine.bank.bank_reconciliation_by_utr's
    # "Settled Subsequent Period" / "Order ID Not Found - Settled After
    # Reco Period" splits (client-reported spec, 2026-08-23; fixed
    # 2026-08-25): the picked "To date" when there is one. When it's left
    # blank (the user filtered by Financial Year alone, without narrowing
    # to a specific date range), fall back to the LATEST order date
    # actually present in this report's own reco_df - the exact same
    # fallback _report_filename() above already uses to name the file, so
    # "the period this report covers" means the same thing everywhere in
    # it. This was previously the far later financial-year-end date
    # instead (e.g. 31-Mar of next year) - which meant that whenever a
    # report was generated by FY alone (no explicit date range), EVERY
    # bank credit up to that FAR future date counted as "within the
    # window", so a genuinely-later-period settlement (e.g. a July
    # report's UTR that actually posted in August) could never be
    # classified as "settled after" at all - reproduced exactly against
    # the client's own July data (46 UTRs with an August bank date were
    # still showing "Order ID Not Found - Settled During Reco Period").
    # Only if reco_df has no usable dates at all (a genuinely degenerate
    # case) does it fall back further to the financial year's own last
    # day, as a last-resort, never-blank cutoff.
    period_end_date = date_to
    if period_end_date is None and reco_df is not None and "created_at" in reco_df.columns and len(reco_df):
        valid_dates = pd.to_datetime(reco_df["created_at"], errors="coerce").dropna()
        if len(valid_dates):
            period_end_date = valid_dates.max()
    if period_end_date is None and filtered_runs:
        period_end_date = storage.financial_year_end_date(filtered_runs[0].get("financial_year"))

    # The report's own period START (added 2026-08-31, client-reported -
    # Executive Summary Point 7 "Receivable Collection Period Analysis"):
    # the exact same fallback logic as period_end_date above, mirrored for
    # the LOWER bound - the picked "From date", else the EARLIEST order
    # date actually in this report's own reco_df, else the financial
    # year's own first day as a last resort. Without this,
    # bank_reconciliation_by_utr's in-window test had no lower bound at
    # all, which is exactly the client-reported bug: filtering "July"
    # showed the ENTIRE June receipt total under "Previous Period Amount
    # Received This Period" instead of only the June orders whose bank
    # credit genuinely posted in July - see that function's own "THE BUG
    # THIS FIXES" docstring note.
    period_start_date = date_from
    if period_start_date is None and reco_df is not None and "created_at" in reco_df.columns and len(reco_df):
        valid_dates = pd.to_datetime(reco_df["created_at"], errors="coerce").dropna()
        if len(valid_dates):
            period_start_date = valid_dates.min()
    if period_start_date is None and filtered_runs:
        period_start_date = storage.financial_year_start_date(filtered_runs[0].get("financial_year"))

    # Executive Summary Point 7's own explicit reconciliation check
    # (client-reported 2026-08-31): "the total of the first three columns
    # should match the actual Bank Credit/Receipt amount for the selected
    # period... This should match the bank statement exactly." Computed
    # here, independently of Bank Reco (UTR-wise) entirely (a formula
    # referencing that same sheet would just restate its own Total row,
    # proving nothing) - the raw uploaded bank ledger's own credit total,
    # filtered to exactly [period_start_date, period_end_date]. A bank row
    # with no parseable date at all can't be judged against either bound -
    # included rather than excluded, the same "no evidence" convention
    # engine.bank.bank_reconciliation_by_utr itself uses.
    # 2026-08-31 fix (client-reported crash, live on her own machine):
    # "TypeError: Invalid comparison between dtype=datetime64[us] and
    # Timestamp" - a plain pd.to_datetime()/pd.Timestamp() pairing here
    # used to compare a bank-date column parsed at microsecond resolution
    # against a scalar Timestamp of a different resolution, which some
    # pandas builds refuse to compare at all. Reusing engine.bank's own
    # to_naive_datetime_series()/to_naive_timestamp() - the same hardened
    # helpers engine.bank.bank_reconciliation_by_utr() itself relies on
    # for this exact kind of date comparison - normalizes both sides to
    # the same (naive, nanosecond) resolution first, so the comparison
    # below can never hit this mismatch again. See those two functions'
    # own docstrings in engine/bank.py for the full story.
    period_bank_credit_total = None
    if bank_ledger_df is not None and not bank_ledger_df.empty and "bank_amount" in bank_ledger_df.columns:
        bank_dates_for_check = to_naive_datetime_series(bank_ledger_df.get("bank_date"))
        in_range_mask = pd.Series(True, index=bank_ledger_df.index)
        if period_start_date is not None:
            start_ts = to_naive_timestamp(period_start_date)
            in_range_mask &= bank_dates_for_check.isna() | (bank_dates_for_check >= start_ts)
        if period_end_date is not None:
            end_ts = to_naive_timestamp(period_end_date)
            in_range_mask &= bank_dates_for_check.isna() | (bank_dates_for_check <= end_ts)
        period_bank_credit_total = round(float(bank_ledger_df.loc[in_range_mask, "bank_amount"].sum()), 2)

    (reco_df, gateway_settlement_overall_df, utr_bank_reco_df, recon_status_df,
     settlement_pending_df, settlement_pending_summary_df, gateway_recon_by_period_df) = _compute_settlement_sections(
        reco_df, consolidated_df, bank_ledger_df, config.get("gateways", []), client_key,
        period_start_date=period_start_date, period_end_date=period_end_date,
        attribution_frames=st.session_state.get("attribution_frames"),
        attribution_sources_cfg=config.get("attribution_sources", []),
    )
    # Client-reported 2026-08-27, rewritten 2026-09-05: "Net Settlement"
    # must net out money not yet bank-credited, and the Dashboard/
    # Executive Summary "Settlement Pending Amount" headline figure must
    # show every still-pending order's exposure (Shiprocket COD/Payu
    # included, not just Delhivery COD) - see engine/reco.py::
    # attach_settlement_pending()'s own docstring. Reassigns reco_df so the
    # headline_totals() call below (and the "Reco working" sheet written by
    # _build_workbook()) both pick up the enriched figure/column. Must run
    # after _compute_settlement_sections() above (which already calls
    # attach_receipt_status() internally, so reco_df's receipt_status/
    # Gateway/query columns are already finalised here).
    reco_df = attach_settlement_pending(reco_df, config.get("gateways", []))

    # Client-reported 2026-08-30 (items 5/6): replaces the old "Payment
    # Gateway Settlement Report" / "Gateway Reconciliation Health"
    # previews with the client's own reference workbook's two-sheet
    # design - see engine/settlement.py.
    if gateway_settlement_overall_df is not None and not gateway_settlement_overall_df.empty:
        st.divider()
        st.subheader("Gateway Settlement overall")
        st.caption(
            "Per gateway, straight off the Bank Reco (UTR-wise) sheet: how much has actually "
            "settled to the bank this period (same month / next month), how much settled "
            "outside this period or against an order ID not found this period, and PG "
            "deductions for this period (from Reco working) - see engine/settlement.py."
        )
        st.dataframe(gateway_settlement_overall_df, use_container_width=True, hide_index=True)

    if gateway_recon_by_period_df is not None and not gateway_recon_by_period_df.empty:
        st.divider()
        st.subheader("Gateway Recon Recoperiod")
        st.caption(
            "Per gateway, for this reconciliation period: Order Value less Gateway deduction, "
            "Refund, Pending Settlement and Order Value partially not received should tie out "
            "(Check) to what's actually Received in Bank - Diff flags where it doesn't. "
            "'Part prepaid and part post paid' is a disclosed gap (always 0) - see "
            "engine/settlement.py::gateway_recon_by_period()'s own docstring."
        )
        st.dataframe(gateway_recon_by_period_df, use_container_width=True, hide_index=True)

    if settlement_pending_summary_df is not None and not settlement_pending_summary_df.empty:
        st.divider()
        st.subheader("Settlement Pending Report")
        st.caption(
            "Every order still owed to the business that hasn't been fully traced to a bank "
            "credit yet - Prepaid gateways and COD partners shown separately, plus orders "
            "flagged Exception / Manual Reconciliation Required. Full order-wise detail is "
            "in the downloaded workbook's \"Settlement Pending Detail\" sheet."
        )
        for group_name, heading in (("Prepaid", "Payment Gateway Settlement Pending"), ("COD", "COD Settlement Pending")):
            group_df = settlement_pending_summary_df[settlement_pending_summary_df["Group"] == group_name]
            if not group_df.empty:
                st.markdown(f"**{heading}**")
                st.dataframe(group_df.drop(columns=["Group"]), use_container_width=True, hide_index=True)

    if utr_bank_reco_df is not None and not utr_bank_reco_df.empty:
        st.divider()
        st.subheader("Bank Reconciliation by Settlement (UTR)")
        # "Matched" as a literal Remarks value is now only a rare
        # degenerate fallback (2026-08-23: Remarks is a composed phrase
        # string - see engine.bank._classify_utr_remark) - a cleanly-tied-
        # out row is any Remarks ending in that suffix.
        matched = int(utr_bank_reco_df["Remarks"].str.endswith("Ties to bank credit").sum())
        st.caption(
            f"{matched:,} of {len(utr_bank_reco_df):,} settlement UTRs tie out cleanly. "
            "Rows not ending \"Ties to bank credit\" are a starting point for manual review, "
            "not a final answer - see engine/bank.py's docstring for the classification rules used."
        )
        st.dataframe(utr_bank_reco_df, use_container_width=True, hide_index=True)
    elif bank_ledger_df is None or (hasattr(bank_ledger_df, "empty") and bank_ledger_df.empty):
        st.caption(
            "No bank statement was uploaded for any of the selected saved month(s), "
            "so the UTR-level bank reconciliation isn't available for this selection."
        )

    st.divider()

    # Fingerprint the current filter state so a stale, previously-prepared
    # file can never be handed out after the filters change. Namespaced by
    # client_key so preparing a download for one channel can't clobber or
    # be confused with another channel's cached bytes when both channels'
    # Reports tabs render in the same session (see render() below).
    bytes_key = f"report_bytes__{client_key}"
    filename_key = f"report_filename__{client_key}"
    fingerprint_key = f"last_prepared_fingerprint__{client_key}"
    overflow_key = f"report_overflow_csvs__{client_key}"
    current_fingerprint = (tuple(sorted(r["file"] for r in filtered_runs)), str(date_from), str(date_to))

    if (st.session_state.get(bytes_key) is not None
            and st.session_state.get(fingerprint_key) != current_fingerprint):
        st.info("Filters changed since you last prepared a report - click below to refresh the download.")

    if st.button("Prepare download", type="primary", key=f"dtc_prepare_download_{client_key}"):
        totals = headline_totals(reco_df)
        filename = _report_filename(reco_df, date_from, date_to)
        data, overflow_csvs = _run_build_with_progress(
            _build_workbook,
            dict(
                reco_df=reco_df, lookup_df=lookup_df, totals=totals,
                channel_name=config.get("channel_name"), date_from=date_from, date_to=date_to,
                gateway_settlement_overall_df=gateway_settlement_overall_df, utr_bank_reco_df=utr_bank_reco_df,
                settlement_pending_df=settlement_pending_df,
                settlement_pending_summary_df=settlement_pending_summary_df,
                gateway_recon_by_period_df=gateway_recon_by_period_df,
                delivery_partner_labels=[cfg["label"] for cfg in config.get("delivery_partners", [])],
                gateway_configs=config.get("gateways", []),
                period_bank_credit_total=period_bank_credit_total,
            ),
            estimate_row_counts=[len(reco_df), len(lookup_df) if lookup_df is not None else 0],
        )
        st.session_state[bytes_key] = data
        st.session_state[filename_key] = filename
        st.session_state[fingerprint_key] = current_fingerprint
        st.session_state[overflow_key] = overflow_csvs

    ready = (
        st.session_state.get(bytes_key) is not None
        and st.session_state.get(fingerprint_key) == current_fingerprint
    )
    if ready:
        st.download_button(
            f"Download {st.session_state[filename_key]}",
            data=st.session_state[bytes_key],
            file_name=st.session_state[filename_key],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dtc_download_btn_{client_key}",
        )
        _render_overflow_downloads(st.session_state.get(overflow_key), client_key, prefix="dtc")


def _render_channel(channel):
    cfg = channel["config"] or {}
    key = channel["client_key"]
    if cfg.get("channel_type") == "marketplace":
        _render_amazon(key, cfg)
    else:
        _render_dtc(key, cfg)


def render():
    st.title("Reports")
    client_key = st.session_state.get("client_key")
    config = st.session_state.get("config")
    if not client_key:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    # Same reasoning as the Dashboard: every channel for this client is
    # saved independently under its own client_key, so all of them should
    # keep showing up here, not just whichever one is currently "active".
    all_channels = get_client_channels((config or {}).get("client_name")) or [
        {"label": st.session_state.get("chosen_label", ""), "client_key": client_key, "config": config or {}}
    ]
    channels_with_data = [ch for ch in all_channels if _channel_has_data(ch)]

    if not channels_with_data:
        st.info(
            "No reconciliation saved yet. Go to **Upload Data** / **Reconciliation** first - "
            "Amazon saves automatically, and DTC/Shopify saves from **Data Management** - "
            "then it'll show up here."
        )
        return

    if len(channels_with_data) == 1:
        _render_channel(channels_with_data[0])
        return

    client_name = channels_with_data[0]["config"].get("client_name", "")
    st.caption(f"{client_name} — all channels with saved reconciliations")
    tabs = st.tabs([ch["config"].get("channel_name", ch["label"]) for ch in channels_with_data])
    for tab, channel in zip(tabs, channels_with_data):
        with tab:
            _render_channel(channel)
