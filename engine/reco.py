"""
reco.py
-------
Layer 3 of the pipeline - the equivalent of your "Reco working" sheet.

Takes:
  - the Shopify order-level data (one row per order, with money values)
  - delivery partner status (Shiprocket / Delhivery / etc.)
  - the per-order receipt summary from consolidator.py

...and joins them into one master reconciliation table with a Diff,
a Settlement amount, and an auto-flagged Query reason for anything that
doesn't tie out - exactly what columns D through AA do in your workbook.
"""

import pandas as pd
from .loaders import normalize_order_id, resolve_col_or_raise, resolve_col
from .bank import COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW
from .attribution import cod_component_of_gateway_label, prepaid_component_of_gateway_label

# Client-reported 2026-08-31 (points 2/3): the ONE canonical definition of
# "this order's money hasn't reached the bank yet" is
# engine.bank.classify_order_bank_status()'s "Reconciliation Category" -
# it already correctly separates "courier/gateway has money we're still
# waiting on" (these three categories) from "nothing to wait on because
# the order was never delivered/collected in the first place" (RTO/
# Cancelled/Lost with no settlement row at all -> COD_NOT_DELIVERED,
# deliberately NOT in this set). See docstring on
# refine_queries_with_settlement_status()/attach_receipt_status() below
# for why reusing this exact set (instead of re-deriving a similar-looking
# "still pending" flag from has_settlement_row/bank_matched directly, as
# both functions used to do) was the actual fix for the COD+RTO
# misclassification bug.
_RECEIPT_PENDING_CATEGORIES = {COD_SETTLEMENT_PENDING, PREPAID_SETTLEMENT_PENDING, EXCEPTION_MANUAL_REVIEW}


def build_order_master(orders_df, orders_cfg):
    """
    Your Shopify order export has one row PER LINE ITEM, so the same order
    number repeats several times. This collapses it to one row per order,
    the way your 'Reco working' sheet does with XLOOKUP/SUMIFS on the first
    matching row.

    Every field below is resolved by HEADER NAME via resolve_col (case/
    whitespace/punctuation-insensitive, same as the upload-time validation
    check) rather than assumed to match the config's literal string exactly
    - a column renamed slightly between export versions (e.g. "financial
    status" vs "Financial Status") used to pass upload validation (which
    already used resolve_col) but then crash right here with a raw pandas
    KeyError, because this function indexed the dataframe with the config's
    literal spec string instead of the column resolve_col actually found.

    Only order_id is truly required (nothing can be joined without it) and
    total (the core money figure everything else measures against) - every
    other field here is best-effort: if the column can't be found at all
    (renamed beyond recognition, or genuinely absent in a trimmed-down
    export), that one field comes back blank/zero instead of failing the
    whole run, per the "don't fail just because a non-essential column is
    missing" requirement.
    """
    df = orders_df.copy()
    order_id_col = resolve_col_or_raise(df, orders_cfg["order_id_col"], orders_cfg.get("label", "Orders"))
    df["order_id"] = normalize_order_id(df[order_id_col])

    created_at_col = resolve_col(df, orders_cfg.get("created_at_col")) if orders_cfg.get("created_at_col") else None

    month_col = resolve_col(df, orders_cfg.get("month_col")) if orders_cfg.get("month_col") else None
    if not month_col:
        # "Month" usually isn't a real Shopify export column - it's something
        # added manually in Excel. Derive it from the order date instead so
        # this doesn't break on a raw, unmodified export.
        date_source = df[created_at_col] if created_at_col else pd.Series(pd.NaT, index=df.index)
        df["_derived_month"] = pd.to_datetime(date_source, errors="coerce").dt.strftime("%B")
        month_col = "_derived_month"

    # (output field name, config key, whether missing should still block the
    # run) - total is the only money field treated as load-bearing enough to
    # warn about via a blank/zero fallback rather than silently vanishing;
    # everything else just degrades to blank/zero.
    optional_fields = [
        ("created_at", "created_at_col"),
        ("financial_status", "financial_status_col"),
        ("fulfillment_status", "fulfillment_status_col"),
        ("subtotal", "subtotal_col"),
        ("shipping", "shipping_col"),
        ("taxes", "taxes_col"),
        ("total", "total_col"),
    ]

    agg_kwargs = dict(month=(month_col, "first"))
    missing_fields = []
    for out_name, cfg_key in optional_fields:
        spec = orders_cfg.get(cfg_key)
        resolved = resolve_col(df, spec) if spec else None
        if resolved:
            agg_kwargs[out_name] = (resolved, "first")
        else:
            missing_fields.append(out_name)

    # Payment Method (e.g. Shopify's "Payment Method" column: "Cash on
    # Delivery (COD)", "Razorpay", "gokwik_pay", ...) - optional, only used
    # to attribute not-yet-settled orders to a gateway for the Payment
    # Gateway Settlement Report's "Pending for Settlement" figure (see
    # engine/settlement.py). Missing/renamed column just means that
    # attribution falls back to "Unattributed" rather than the whole run failing.
    payment_method_col = resolve_col(df, orders_cfg.get("payment_method_col")) if orders_cfg.get("payment_method_col") else None
    if payment_method_col:
        agg_kwargs["payment_method"] = (payment_method_col, "first")

    agg = df.groupby("order_id").agg(**agg_kwargs).reset_index()

    if "payment_method" not in agg.columns:
        agg["payment_method"] = None

    # Backfill any field whose source column couldn't be found at all, so
    # every column this function has ever returned is still present and
    # every downstream computation that reads it keeps working - just with
    # a blank/zero for that one field instead of a KeyError.
    for out_name, _ in optional_fields:
        if out_name not in agg.columns:
            agg[out_name] = 0 if out_name in ("subtotal", "shipping", "taxes", "total") else None

    for c in ["subtotal", "shipping", "taxes", "total"]:
        agg[c] = pd.to_numeric(agg[c], errors="coerce").fillna(0)

    if missing_fields:
        agg.attrs["missing_order_fields"] = missing_fields

    return agg


def classify_status(raw_status):
    """
    Maps a courier/marketplace's raw status text to one of our standard
    buckets: Delivered, RTO, In Transit, Cancelled, Undelivered, Lost,
    Refunded, or Other.

    This is keyword-based rather than an exact-match list, because every
    courier spells things differently (Delhivery's "RETURNED_TO_ORIGIN" vs
    Shiprocket's "RTO DELIVERED" vs Unicommerce's "RTO_IN_TRANSIT" all mean
    the same thing). An exact-match list breaks the moment a courier uses
    a status text you didn't anticipate - which is exactly what happened
    with order #16411 ("RETURNED_TO_ORIGIN" wasn't in the old RTO list).

    Returns None only for genuinely blank/missing status values - a real
    status of any kind should always map to a real bucket, not "Undefined".
    """
    if raw_status is None:
        return None
    s = str(raw_status).strip().upper()
    if not s or s in ("NAN", "NONE", ""):
        return None

    if "RTO" in s or "RETURN" in s:
        return "RTO"
    if "CANCEL" in s:
        return "Cancelled"
    if "REFUND" in s:
        return "Refunded"
    if "LOST" in s:
        return "Lost"
    if "OUT_FOR_DELIVERY" in s or "OUT FOR DELIVERY" in s or s == "OFD":
        return "In Transit"
    if "UNDELIVER" in s or "NDR" in s or "PICKUP EXCEPTION" in s or "PICKUP_FAILED" in s or "FAILED" in s:
        return "Undelivered"
    if "DELIVER" in s:
        return "Delivered"
    if "TRANSIT" in s or "SHIP" in s or "PICK" in s or "NEW ORDER" in s or "PROCESSING" in s or "READY" in s:
        return "In Transit"
    return "Other"


# How definitive/final each classify_status() bucket is, used ONLY to pick
# a winner when more than one delivery partner has a classifiable status
# for the SAME order (see attach_delivery_status()'s resolve_row() below -
# client-reported 2026-09-04, order #31475: Shiprocket cancelled a leg,
# Delhivery genuinely delivered a re-shipment of the same order, and the
# tool needs to show Delhivery/Delivered, not whichever courier happens to
# be listed first in configs/*.json). Higher = wins. "Cancelled" is
# deliberately the LOWEST real bucket: it means nothing happened on THAT
# courier's leg, which is exactly the kind of outcome a genuine result
# from a different courier (delivered, returned, lost, even still in
# transit) should be allowed to supersede. Two partners that land in the
# SAME bucket (both "Delivered", both "Cancelled", ...) still tie-break to
# the original config priority order, unchanged from before this fix.
_STATUS_FINALITY = {
    "Delivered": 6,
    "RTO": 5,
    "Lost": 5,
    "Refunded": 5,
    "Undelivered": 3,
    "In Transit": 2,
    "Other": 2,
    "Cancelled": 1,
}


def attach_delivery_status(order_master, delivery_frames, delivery_configs):
    """
    Checks EVERY delivery partner file for each order (not just the first
    one that has it), and records:
      - each partner's raw status, in its own column (e.g. shiprocket_raw_status)
        - kept regardless of which partner ends up deciding the final status,
        so you can always cross-check against Unicommerce or any other source.
      - which partner's status was used to decide the final status
      - the final status, classified into a standard bucket
      - delivered_date / rto_date, when that partner's file has those columns

    Priority = order of the list in the config, but critically: if the
    first partner that has the order gives a status we can't classify
    (rare, but possible with a truly novel status text), we keep checking
    the remaining partners instead of giving up and showing "Status
    Undefined" - that silent giving-up was the actual bug behind order
    #16411 showing Undefined despite Delhivery clearly showing RTO.

    Client-reported 2026-09-05 (order #32318): a REAL courier's own
    classifiable status always wins over an OMS/WMS platform's (currently
    just Unicommerce, identified by its config setting courier_label_col
    - see resolve_row() below), regardless of which bucket looks more
    "final" by _STATUS_FINALITY. Unicommerce only ever tracks what the
    real courier (and any manual ops adjustment) told it and can lag
    behind, so it's consulted only when NO real courier gave a
    classifiable status for the order at all - never used to override
    one. Order #32318 itself: Delhivery says Lost, Unicommerce's own
    tracking still shows Delivered - final_delivery_status must still
    resolve to Lost.
    """
    df = order_master.copy()
    df["delivery_partner"] = None
    df["final_delivery_status"] = "Status Undefined"
    df["delivered_date"] = None
    df["rto_date"] = None

    partner_status_lookups = {}
    partner_courier_label_lookups = {}
    for cfg in delivery_configs:
        label = cfg["label"]
        if label not in delivery_frames:
            continue
        partner_df = delivery_frames[label].copy()
        order_id_col = resolve_col_or_raise(partner_df, cfg["order_id_col"], label)
        partner_df["order_id"] = normalize_order_id(partner_df[order_id_col])
        # status_col is this file's entire reason for existing (a delivery
        # partner report with no status column is useless), so it's still
        # required - but resolved by header name like everything else here,
        # not assumed to match the config's literal string exactly (a file
        # that already passed upload validation via resolve_col could
        # otherwise still crash right here on a slightly-renamed header).
        status_col = resolve_col_or_raise(partner_df, cfg["status_col"], label)
        status_lookup = partner_df.groupby("order_id")[status_col].first()

        # Some sources leave the primary status blank in legitimate cases
        # (e.g. Unicommerce's "Shipping Tracking Status" is blank when an
        # order was cancelled before it ever shipped) but have a secondary
        # column that IS populated for those rows. Fill gaps from it.
        fallback_col = resolve_col(partner_df, cfg.get("status_col_fallback")) if cfg.get("status_col_fallback") else None
        if fallback_col:
            fallback_lookup = partner_df.groupby("order_id")[fallback_col].first()
            status_lookup = status_lookup.where(status_lookup.notna(), fallback_lookup)

        partner_status_lookups[label] = status_lookup

        # Always capture this partner's raw status against every order it
        # has, regardless of who ends up "winning" - needed for audit /
        # cross-checking and the exception report's per-source columns.
        raw_col = f"{label.lower().replace(' ', '_')}_raw_status"
        df[raw_col] = df["order_id"].map(status_lookup)

        # Some sources aren't a courier themselves - Unicommerce is a WMS/OMS
        # platform that ships through whichever courier the order was
        # actually handed to, and names that real courier in its own
        # "Shipping provider" column (client-reported 2026-08-27: the tool
        # was writing the literal label "Unicommerce" as the delivery
        # partner instead of the courier Unicommerce itself names). When a
        # config entry sets courier_label_col, capture that per-order text
        # here so resolve_row() below can substitute it for this source's
        # own generic label.
        courier_label_col = resolve_col(partner_df, cfg.get("courier_label_col")) if cfg.get("courier_label_col") else None
        if courier_label_col:
            partner_courier_label_lookups[label] = partner_df.groupby("order_id")[courier_label_col].first()

        delivered_col = resolve_col(partner_df, cfg.get("delivered_date_col")) if cfg.get("delivered_date_col") else None
        if delivered_col:
            date_lookup = partner_df.groupby("order_id")[delivered_col].first()
            matched = df["order_id"].isin(date_lookup.index)
            # .to_dict() first, not .map()'d straight off the Series - see
            # engine/summary.py's month_summary for the same fix (and the
            # full explanation): mapping a datetime-typed Series that's
            # EMPTY (this partner's file resolved to zero order_id groups -
            # e.g. an uploaded delivery report with a header but no data
            # rows) crashes on some pandas versions/dtype backends
            # ("TypeError: Cannot cast DatetimeArray to dtype float64"),
            # regardless of whether the .map() target is also empty. A
            # plain dict sidesteps that pandas dtype-inference path
            # entirely, empty or not.
            df.loc[matched, "delivered_date"] = df.loc[matched, "order_id"].map(date_lookup.to_dict())

        rto_date_col = resolve_col(partner_df, cfg.get("rto_date_col")) if cfg.get("rto_date_col") else None
        if rto_date_col:
            rto_lookup = partner_df.groupby("order_id")[rto_date_col].first()
            matched = df["order_id"].isin(rto_lookup.index)
            df.loc[matched, "rto_date"] = df.loc[matched, "order_id"].map(rto_lookup.to_dict())

    # Now decide the final status. Client-reported 2026-09-04 (order
    # #31475): an order can genuinely be handed to MORE THAN ONE courier
    # over its life - picked up by Shiprocket, cancelled there, then
    # re-shipped and actually delivered by Delhivery. The old rule below
    # simply walked delivery_configs in config-array order and took the
    # FIRST partner with any classifiable status at all - so Shiprocket's
    # stale "CANCELED" (still sitting in its own raw file, correctly, as a
    # record of that abandoned leg) outranked Delhivery's later genuine
    # "DELIVERED" purely because Shiprocket happens to be listed first in
    # configs/*.json, even though Delhivery's status is the one that
    # actually describes what happened to the order. Confirmed against the
    # client's own July data: every one of the 8 orders where two couriers
    # both have a row (#28296/31475/31669/31803/32006/32192 = one courier
    # Cancelled + another Delivered/RTO; #26444/26829 = one courier's order
    # barely created ("NEW ORDER"/In Transit) while another already shows
    # RTO_DELIVERED) should show the courier with the more DEFINITIVE
    # outcome, not the earlier-priority one.
    #
    # Fix: when more than one configured partner has a classifiable status
    # for the same order, rank each classified bucket by how definitive/
    # final an outcome it represents (_STATUS_FINALITY below) and pick the
    # partner with the single most definitive one - falling back to the
    # original config priority order only to break a genuine tie (e.g. two
    # couriers both say "Delivered" - keep whichever is listed first,
    # unchanged from before). "Cancelled" ranks LOWEST of every real
    # bucket here on purpose: a cancelled leg on one courier is exactly the
    # kind of outcome a genuine delivery/RTO/lost result on a DIFFERENT
    # courier should be allowed to supersede, whereas two couriers that
    # both say "Cancelled" (or anything else that agrees) still tie-break
    # to the config's own priority order same as always. The vast majority
    # of orders (7276 of 7289 on the client's own July data) have only one
    # courier with a row at all, so this is a no-op for them - identical
    # output to before.
    def resolve_row(order_id):
        candidates = []  # (priority_index, label, classified) for every
        # REAL COURIER (or manual-report) partner that gave this order a
        # classifiable status - almost always just one.
        oms_candidates = []  # same, but for an OMS/WMS platform config
        # (cfg["courier_label_col"] set - currently just Unicommerce).
        for priority_index, cfg in enumerate(delivery_configs):
            label = cfg["label"]
            lookup = partner_status_lookups.get(label)
            if lookup is None or order_id not in lookup.index:
                continue
            classified = classify_status(lookup[order_id])
            if classified is None:
                continue
            (oms_candidates if cfg.get("courier_label_col") else candidates).append(
                (priority_index, label, classified)
            )

        # Client-reported 2026-09-05 (order #32318): Unicommerce is an
        # order-management/WMS platform, not a courier - it just tracks
        # whatever the real courier (and any manual ops adjustment) told
        # it, and can lag behind. It must never outrank an actual
        # courier's own classifiable status, no matter how "final" its
        # own reported status looks by _STATUS_FINALITY (e.g. Unicommerce
        # showing "Delivered" while Delhivery's own feed already says
        # "Lost" must still resolve to Lost - Delhivery is the real
        # courier here and Unicommerce is only ever a fallback). So a real
        # courier's classifiable status is used whenever ANY exists; an
        # OMS/WMS platform's own status is only even considered when NO
        # real courier gave a classifiable status for this order at all
        # ("if status not clear then rely on Unicommerce", per the
        # client's own framing). The finality-based tie-break above
        # (order #31475) still applies WITHIN each of these two groups -
        # e.g. two real couriers disagreeing is unaffected by this change.
        candidate_pool = candidates if candidates else oms_candidates

        if candidate_pool:
            best_priority, best_label, best_classified = max(
                candidate_pool,
                key=lambda c: (_STATUS_FINALITY.get(c[2], 0), -c[0]),
            )
            display_label = best_label
            # Prefer this source's own real-courier column (e.g.
            # Unicommerce's "Shipping provider") over its generic label,
            # falling back to the label when the column is missing/blank
            # for this order so behavior degrades safely.
            courier_lookup = partner_courier_label_lookups.get(best_label)
            if courier_lookup is not None and order_id in courier_lookup.index:
                courier_name = courier_lookup[order_id]
                if pd.notna(courier_name) and str(courier_name).strip():
                    display_label = str(courier_name).strip()
            # Client-reported 2026-09-04 (order #30091): Unicommerce is an
            # order-management/inventory system, not a courier - it should
            # never itself be shown as the delivery_partner. Normally its
            # own "Shipping provider" column (courier_lookup above) names
            # the REAL courier and that substitution already handles it;
            # this only fires on the residual case where Unicommerce is the
            # order's only resolvable source AND its own "Shipping
            # provider" is blank for this order too (Unicommerce genuinely
            # doesn't know who shipped it either) - confirmed against the
            # client's own reference workbook, which shows the literal
            # "partner undifined" (her own spelling) for exactly this case
            # rather than the software's name.
            if display_label == "Unicommerce":
                display_label = "partner undifined"
            return display_label, best_classified

        # No partner gave a classifiable status - but if ANY partner at
        # least had the order, say so rather than a bare "Undefined".
        for cfg in delivery_configs:
            label = cfg["label"]
            lookup = partner_status_lookups.get(label)
            if lookup is not None and order_id in lookup.index:
                return ("partner undifined" if label == "Unicommerce" else label), "Status Undefined"
        return None, "Status Undefined"

    resolved = df["order_id"].apply(resolve_row)
    df["delivery_partner"] = resolved.apply(lambda t: t[0])
    df["final_delivery_status"] = resolved.apply(lambda t: t[1])

    return df


def attach_receipts_and_diff(order_master, receipt_summary):
    """
    Joins in the receipt/deduction/refund totals per order and computes:
      Diff       = Total order value - Receipt received
      Settlement = Receipt received - Deduction
    Exactly the X3 and Z3 formulas from your 'Reco working' sheet.
    """
    df = order_master.merge(receipt_summary, on="order_id", how="left")
    for c in ["receipt_amount", "total_deduction", "refund_amount"]:
        df[c] = df[c].fillna(0)

    df["diff"] = df["total"] - df["receipt_amount"]
    df["settlement_amount"] = df["receipt_amount"] - df["total_deduction"] - df["refund_amount"]
    return df


def attach_pending_cod_receipts(reco_df, pending_cod_df):
    """
    Client-reported 2026-09-04 (round 10, point 2): "Shiprocket COD Amount
    Not Reflecting in receipt_amount" - see engine.consolidator.build_
    pending_cod_receipts()'s own docstring for the full root-cause story
    (normalize_gateway_df's settled_status_col filter drops a not-yet-
    remitted Shiprocket/Prozo COD row entirely before it ever reaches
    receipt_summary, so attach_receipts_and_diff() above never sees it -
    receipt_amount stays 0 even though the courier's own COD report
    plainly lists the amount).

    pending_cod_df is that same function's own output - the not-yet-
    settled rows, kept ALONGSIDE (not instead of) the existing settled-
    only consolidated_df/receipt_summary. This rescues receipt_amount
    (and, so the "diff"/"settlement_amount" formulas stay internally
    consistent, total_deduction alongside it) for an order that:
      - is Delivered by the SAME courier the pending row came from (a COD
        gateway's own label minus its " COD" suffix - e.g. a "Shiprocket
        COD" pending row only ever fills in for an order whose delivery_
        partner is "Shiprocket" - never a different courier's order that
        happens to share an order_id collision), and
      - has receipt_amount == 0 so far (nothing from ANY other, already-
        settled source has already been recorded for it) - deliberately
        conservative: an order already showing a genuine partial receipt
        from some other source is left exactly as the existing pipeline
        already treats it, rather than risk double-counting a scenario
        this fix wasn't asked to handle.

    Deliberately does NOT touch engine.bank.classify_order_bank_status()'s
    own has_settlement_row/Reconciliation Category at all - that function
    keeps computing both from the ORIGINAL, unchanged, settled-only
    consolidated_df, so an order this function rescues still correctly
    lands in COD_SETTLEMENT_PENDING exactly as before (the "has this
    money reached OUR bank" question is unaffected); only the Reco
    working sheet's own receipt_amount/diff/settlement_amount - and,
    downstream, the Recipt Remark / Query text, both of which now key off
    receipt_amount for exactly this reason - see attach_receipt_status()
    and refine_queries_with_settlement_status() below - change.

    Call this right after run_shopify_pipeline(), before
    classify_order_bank_status() - so recon_status_df/settlement_pending_
    df/every later step in the pipeline all see the corrected receipt_
    amount consistently, and (once the period is saved via engine.storage.
    save_run()) so the fix is permanent for that saved period, not just
    the current screen - mirroring how every other "wire into the core
    pipeline, not just the exported workbook" fix in this engine works.

    pending_cod_df may legitimately be None/empty (no COD gateway in this
    client's config declares settled_status_col, or none of those files
    were uploaded this run, or none of their rows are pending) - reco_df
    is returned completely unchanged in that case.
    """
    df = reco_df.copy()
    if pending_cod_df is None or pending_cod_df.empty:
        return df
    if "delivery_partner" not in df.columns or "final_delivery_status" not in df.columns:
        return df

    df["order_id"] = df["order_id"].astype(str)
    pending = pending_cod_df.copy()
    pending["order_id"] = pending["order_id"].astype(str)
    pending["_courier"] = pending["source"].astype(str).str.replace(r"\s*COD$", "", regex=True).str.strip()

    amount_by_key = pending.groupby(["order_id", "_courier"])["amount"].sum().to_dict()
    deduction_by_key = pending.groupby(["order_id", "_courier"])["deduction"].sum().to_dict()

    receipt = df["receipt_amount"].fillna(0.0) if "receipt_amount" in df.columns else pd.Series(0.0, index=df.index)
    eligible = (df["final_delivery_status"] == "Delivered") & (receipt.abs() <= 0.004)
    if not eligible.any():
        return df

    keys = list(zip(df["order_id"], df["delivery_partner"].astype(str)))
    matched_amount = pd.Series([amount_by_key.get(k) for k in keys], index=df.index)
    matched_deduction = pd.Series([deduction_by_key.get(k) for k in keys], index=df.index)
    apply_mask = eligible & matched_amount.notna() & (matched_amount.fillna(0.0).abs() > 0.004)
    if not apply_mask.any():
        return df

    if "total_deduction" not in df.columns:
        df["total_deduction"] = 0.0
    df.loc[apply_mask, "receipt_amount"] = matched_amount[apply_mask]
    df.loc[apply_mask, "total_deduction"] = (
        df.loc[apply_mask, "total_deduction"].fillna(0.0) + matched_deduction[apply_mask].fillna(0.0)
    )
    if "refund_amount" not in df.columns:
        df["refund_amount"] = 0.0
    df["diff"] = df["total"] - df["receipt_amount"]
    df["settlement_amount"] = df["receipt_amount"] - df["total_deduction"] - df["refund_amount"]
    return df


def attach_settlement_pending(reco_df, gateway_configs):
    """
    Attaches each order's outstanding Settlement Pending amount onto
    reco_df as "settlement_pending_amount", so engine/summary.py's
    headline_totals() can report it as its own "Settlement pending"
    headline figure - the same figure the Dashboard tile and Executive
    Summary "Headline Numbers" section both read off of.

    Rewritten 2026-09-05 (client-reported): the Dashboard/Executive
    Summary "Settlement Pending Amount" only showed Delhivery COD's
    pending money - Shiprocket COD's ₹1,17,670.47 and Payu's pending
    amount were both missing. The PREVIOUS implementation summed engine/
    settlement_pending.py::build_settlement_pending_report()'s "Settlement
    Amount", filtered to that same report's "Gateway Amount" > 0.01 -
    where "Gateway Amount" is read off recon_status_df (engine.bank.
    classify_order_bank_status()'s output, computed from the ORIGINAL,
    unrescued consolidated_df). Two different kinds of still-pending money
    fell through that filter for two different reasons:
      - Shiprocket COD (or any settled_status_col-configured courier)
        orders engine.reco.attach_pending_cod_receipts() rescues (round
        10) - recon_status_df's own receipt_amount never learns about that
        rescue, so "Gateway Amount" stayed 0 for them there even though
        reco_df's own receipt_amount is correctly populated.
      - Payu (or any prepaid gateway) orders with NOTHING collected yet at
        all - excluded on purpose under the OLD design, which existed only
        to net Net Settlement (see the removed docstring's own history:
        money never added to receipt_amount shouldn't be double-subtracted
        from it). That purpose no longer needs this exclusion - see below.

    Now built from engine/settlement_pending.py::pending_amount_by_order(),
    the SAME per-order helper settlement_pending_summary_by_gateway() (the
    by-gateway Settlement Pending Summary sheet) already uses - so the
    Dashboard/Executive Summary headline figure and that by-gateway sheet
    can never disagree again, and BOTH now correctly include every still-
    pending order regardless of gateway or whether anything's been
    collected yet (Payu/Gokwik orders with nothing collected show their
    full order Total as outstanding, exactly like the Settlement Pending
    Summary sheet already did for them).

    Why this no longer needs the old "only if Gateway Amount > 0" carve-
    out for Net Settlement purposes: engine/reco.py::attach_receipt_status()
    (2026-09-05, client-reported separately - the "Bank credit" leak fix)
    now directly zeroes "settlement_amount" for every order still in a
    pending Reconciliation Category, BEFORE this function ever runs. So
    headline_totals()'s "pending_deduction" (settlement_pending_amount
    capped at each order's own non-negative settlement_amount) is already
    0 for every such order regardless of how large settlement_pending_
    amount is - settlement_amount, not this column, is what actually
    prevents Net Settlement from double-counting still-pending money now.
    This column can therefore safely show the FULL outstanding exposure
    (matching the Settlement Pending Summary sheet) without corrupting Net
    Settlement - see headline_totals()'s own docstring for that reasoning
    in full.

    Must run after attach_receipt_status() (needs reco_df's own finalised
    "receipt_status"/"Gateway"/"query" columns - the same signals that
    function and settlement_pending_summary_by_gateway() already key off)
    - both views/page_reconciliation.py and views/page_reports.py call
    this right after attach_receipt_status(), not before.
    """
    from .settlement_pending import pending_amount_by_order

    df = reco_df.copy()
    df["settlement_pending_amount"] = pending_amount_by_order(df, gateway_configs).fillna(0.0)
    return df


def flag_queries(df):
    """
    Auto-generates the same kind of query buckets you track manually in
    column AA - e.g. "Delivered but amount not received". This is meant as
    a first-pass flag for your team to review, not a final answer.

    Important: RTO/Cancelled orders only get flagged as a refund concern
    when money was ACTUALLY received (receipt_amount > 0) and not yet
    refunded. For COD orders, nothing is collected until delivery, so an
    RTO simply means "never collected" - that's normal, not an exception.
    Flagging every COD RTO as "refund pending" (comparing total order value
    against a receipt that was always going to be zero) was the bug here.

    Client-reported 2026-08-30 (phrasing/coverage match against their own
    corrected workbook): two changes on top of the categories above -
      - A DELIVERED order that has a refund recorded against it at all
        (refund_amount > 1) is now its own flag ("Delivery status
        Delivered - Why Refunded?") regardless of diff - a delivered order
        being refunded is unusual enough on its own to warrant a look,
        whether or not the remaining money still ties out.
      - The RTO/Cancelled "amount received but not refunded" wording now
        matches the client's own house phrasing ("Delhivery status is
        {status} what is Refund status?" - "Delhivery" is the client's own
        spelling, not a typo introduced here) rather than this engine's
        original generic wording, purely a rename - the underlying
        received-not-refunded condition is unchanged.
    Only the wording/coverage of THIS function's own diff-based rules
    changed. The separate "money's been collected by the gateway/courier
    but hasn't reached the bank yet" queries (e.g. "Payu Setlment
    Pending", "Shiprocket COD setlment pending") are NOT decided here -
    see refine_queries_with_settlement_status() below, which runs later
    (once bank-matching/reconciliation-category data exists) and can
    overwrite an "Okk" this function assigned once diff alone said the
    order tied out.

    Client-reported 2026-08-31 (point 2), corrected on direct cell-by-cell
    comparison against the client's own July Reco working data (988+
    examples), on top of everything above:
      - "Delivery status Delivered - Why Refunded?" used to fire on ANY
        refund at all on a Delivered order. The client's own data shows
        that's wrong when only PART of what was received got refunded
        (e.g. total 2147, receipt 2147, refund 549 - a normal partial
        adjustment/discount refund on an otherwise fully-paid order,
        20 such examples, all left "Okk") - only a refund that consumes
        the ENTIRE amount actually received (refund >= receipt, within
        rounding) is the genuine "paid then fully reversed on a Delivered
        order" anomaly worth flagging (3 confirmed examples). Also fixed
        the flagged text itself to match the client's own exact wording -
        "Delivery status Delivered Why Refunded", no dash, no "?".
      - The RTO/Cancelled/Lost/Refunded/Status Undefined "received but not
        refunded -> what's the refund status?" flag used to apply to all
        five statuses. The client's data shows it genuinely only applies
        to RTO and Lost (21 + more confirmed RTO examples) - a Cancelled
        order that was fully paid and never refunded is, in the client's
        own workbook, left as a plain "Okk" (4 confirmed examples;
        Status Undefined's own occasional use of this wording is a
        separate, already-disclosed judgment call - see
        refine_queries_with_settlement_status()'s docstring). Narrowed
        accordingly - Cancelled/Refunded/Status Undefined orders that are
        received-and-not-refunded now read "Okk" rather than being flagged.
      - New: any of RTO/Cancelled/Lost/Refunded/Status Undefined with a
        PARTIAL receipt that was itself fully refunded (0 < receipt <
        total, refund >= receipt) reads "Why Partially received and same
        amount refunded" - a real, reproducible pattern confirmed across
        all of those statuses (4 examples: RTO, Cancelled, and 2x Status
        Undefined), previously not generated by this function at all.
      - "Status Undefined" no longer gets its own blanket "Delivery status
        Required" - the client's own data shows that phrase belongs to
        "In Transit" instead (see below).
      - "In Transit" now reads "Delivery status Required" when something
        has actually been received (payment came in but the shipment's
        outcome is still unclear - worth a look), "Okk" when nothing has
        (still moving, nothing pending yet - the original, correct
        behaviour for the common case).

    Client-reported 2026-08-31 (round 4, point 2) - OVERRIDES the previous
    "Status Undefined follows Cancelled" note above: a Status Undefined
    order with nothing received no longer reads "Okk" - the client gave an
    explicit, confirmed rule that it should read
    "COD Delivery status Undifined  Amount not Received" (the client's own
    spelling/spacing, kept verbatim), since an undetermined delivery
    status is a genuine open question for their team, unlike RTO/Lost
    where "nothing collected" is the expected, normal state. Status
    Undefined is therefore its own branch now, not grouped with RTO/
    Cancelled/Lost/Refunded - the partial-receipt-fully-refunded pattern
    still applies to it exactly as before; only the "nothing received"
    case's wording changes. See attach_receipt_status() below for the
    matching Recipt Remark change ("Not received").

    Client-reported 2026-09-05 (point 1): a Status Undefined order that
    HAS received something (in full, or partially without a full refund -
    the two cases already carved out above still take priority) used to
    fall through to "Okk" - wrong, since an undetermined delivery status
    with money already in hand is itself an open question, not a clean
    tie-out. Now reads "Amount received Delivery status undefined"
    (example order #26735). Note this is Layer-3's own diff-only wording;
    when the order also resolves to a specific pending payment gateway
    (e.g. Payu) via bank-matching data, refine_queries_with_settlement_
    status() below can further replace the "nothing received" branch's
    text with a gateway-specific one - see that function's own docstring
    (point 2).

    Client-reported 2026-09-05 (point 3): a Lost shipment used to be
    treated the same as an RTO - flagged only when something was actually
    received and not refunded, otherwise a plain "Okk" (nothing collected
    being the normal, expected COD outcome). The client's own explicit
    rule: a Lost shipment is ALWAYS an open question about credit note
    status, regardless of whether anything was ever collected - so Lost
    now always reads "Shipment LOST what is Credit note Status?" (example
    order #28994), replacing the previous conditional RTO-shared wording
    entirely. RTO's own behaviour (and Cancelled/Refunded) is unchanged.

    Disclosed gap, not guessed around: a Delivered order whose
    receipt_amount sits at roughly half of Total with nothing further
    missing (order value genuinely split part-prepaid/part-COD, ~36
    examples in the client's July data, mostly exactly 50% but ranging
    13%-75%) still reads "Partial Payment Received" here rather than the
    client's own "Okk" - there is no column in reco_df today that marks an
    order as a genuine split payment (see engine/settlement.py's own
    "Part prepaid and part post paid" column, already disclosed there as
    always 0 for the same reason), and the receipt ratio alone isn't a
    safe way to tell a real split payment from a real shortfall apart from
    guessing a threshold - left as the pre-existing (already
    client-confirmed) diff-based behaviour rather than invented.
    """
    def classify_row(row):
        status = row["final_delivery_status"]
        diff = row["diff"]
        total = row["total"]
        receipt = row["receipt_amount"]
        refund = row["refund_amount"]
        received = receipt > 1
        refunded = refund > 1
        partial_receipt = received and receipt < total - 1

        if status == "Delivered":
            if refund > 1 and refund >= receipt - 1:
                # The ENTIRE amount actually received (not just "some"
                # refund) came back - a Delivered order fully paid-and-
                # reversed (or a partial receipt refunded in full) is the
                # real anomaly; a partial refund on top of money that's
                # still substantially retained is normal and not flagged
                # (see docstring above).
                return "Delivery status Delivered Why Refunded"
            if diff > 1:
                # Some money is missing - but "some" could mean "none at all"
                # or "part of it". These are very different situations and
                # need different follow-up, so don't collapse them into one
                # generic message (that was the bug behind order #18359).
                if receipt > 1:
                    # Flat, amount-free text on purpose (client-reported
                    # 2026-08-27): embedding the rupee amounts here made every
                    # partial-payment order's remark unique, so
                    # engine/summary.py's open_queries() rollup - which groups
                    # by this exact text - could never collapse them into one
                    # summary line; it produced one row per order instead.
                    # Keeping it flat lets multiple partial-payment
                    # transactions/orders consolidate into a single "Partial
                    # Payment Received" line as intended.
                    return "Partial Payment Received"
                return "Delivered but amount not received - reason?"
            if diff < -1:
                excess = receipt - total
                return f"Excess Payment Received - ₹{excess:,.2f} extra"
            return "Okk"
        if status == "Status Undefined":
            # Client-reported 2026-08-31 (round 4, point 2), overriding the
            # previously-disclosed "follows Cancelled" judgment call above:
            # the client gave an explicit, confirmed rule for this status -
            # a Status Undefined order with NOTHING received should read
            # exactly "COD Delivery status Undifined  Amount not Received"
            # (the client's own spelling/spacing, kept verbatim so this
            # matches their sheet on a text diff - not a typo introduced
            # here), not "Okk". This is a real open query for their team
            # (the delivery status itself couldn't be determined at all),
            # unlike RTO/Lost/Cancelled where "nothing collected" is the
            # normal, expected state - so Status Undefined is pulled out of
            # the shared block below into its own branch. The partial-
            # receipt-fully-refunded pattern still applies here exactly as
            # it does for RTO/Cancelled/Lost/Refunded (unchanged, confirmed
            # against real examples of this exact combination).
            if partial_receipt and refund >= receipt - 1:
                return "Why Partially received and same amount refunded"
            if not received:
                return "COD Delivery status Undifined  Amount not Received"
            # Client-reported 2026-09-05 (point 1): previously fell through
            # to "Okk" whenever something was actually received (the
            # partial-fully-refunded and nothing-received cases above were
            # already handled) - an order that HAS a receipt but whose
            # delivery status still couldn't be determined is itself worth
            # a look, not a clean "Okk". Covers both a full receipt and a
            # partial receipt that was NOT fully refunded (order 26735).
            return "Amount received Delivery status undefined"
        if status in ("RTO", "Cancelled", "Lost", "Refunded"):
            if partial_receipt and refund >= receipt - 1:
                # Whatever partial amount came in went straight back out -
                # confirmed across RTO/Cancelled/Status Undefined orders in
                # the client's own data, not status-specific.
                return "Why Partially received and same amount refunded"
            if status == "Lost":
                # Client-reported 2026-09-05 (point 3): previously only
                # flagged a Lost shipment when something was actually
                # received and not refunded, otherwise falling through to
                # "Okk" (treated the same as an RTO - "nothing collected is
                # normal for COD"). The client's own explicit rule: a Lost
                # shipment is ALWAYS an open question about credit note
                # status, regardless of whether anything was ever
                # collected - so this is now unconditional and no longer
                # shares the RTO received-and-not-refunded branch below.
                return "Shipment LOST what is Credit note Status?"
            if status == "RTO" and received and not refunded:
                # For COD, nothing is collected until delivery, so RTO
                # simply means "never collected" - normal, not an exception
                # - UNLESS something WAS actually received and hasn't been
                # refunded, which is the real anomaly worth flagging.
                # Cancelled/Refunded do NOT get this same flag even when
                # received-and-not-refunded (confirmed against the client's
                # own data - see docstring above).
                return f"Delhivery status is {status} what is Refund status?"
            return "Okk"
        if status == "Undelivered":
            return "Undelivered - delivery attempt failed, follow up required"
        if status == "In Transit":
            return "Delivery status Required" if received else "Okk"
        if status == "Other" and diff > 1:
            return "Status unclear - review manually"
        return "Okk"

    df["query"] = df.apply(classify_row, axis=1)
    return df


# Payment-provider/COD-courier labels whose own settlement-pending state
# (engine.bank.classify_order_bank_status()'s has_settlement_row/
# bank_matched, consumed via recon_status_df) can override an "Okk" from
# flag_queries() above into a specific "still with the
# gateway/courier, not yet bank-credited" query - see
# refine_queries_with_settlement_status()'s own docstring. Keyed by the
# exact "Gateway" column label (engine.attribution.build_payment_gateway_
# lookups' own COD label convention, e.g. "Shiprocket COD") to the phrase
# used for a FULLY-reflecting pending order of that courier - confirmed
# against the client's own corrected July workbook for Delhivery COD /
# Shiprocket COD; Prozo COD's own phrase is inferred by the same pattern
# (no Prozo orders were actually sitting in this state in the client's
# reference data to confirm the exact wording against).
_COD_SETTLEMENT_PENDING_PHRASES = {
    "Delhivery COD": "Delhivery COD Setlment pending",
    "Shiprocket COD": "Shiprocket COD setlment pending",
    "Prozo COD": "Prozo COD setlment pending",
}


def refine_queries_with_settlement_status(reco_df, recon_status_df):
    """
    Client-reported 2026-08-30: a large block of orders (~300+ on the
    client's own July data) that flag_queries() above correctly marks
    "Okk" purely by diff (receipt_amount already equals total - nothing
    LOOKS missing from the Reco working sheet's own numbers) are, in the
    client's own corrected workbook, still treated as open queries -
    because the money hasn't actually reached the BANK yet, even though
    the courier/gateway has already reported it collected. That's exactly
    what engine.bank.classify_order_bank_status() already tracks
    (has_settlement_row + bank_matched, surfaced here via recon_status_df)
    but flag_queries() above has no visibility into, since it runs in
    Layer 3 - before bank matching exists at all (see run_shopify_pipeline
    below). This is therefore a SEPARATE, later pass: both
    views/page_reconciliation.py and views/page_reports.py call it right
    after recon_status_df/the Gateway column both exist (the same point
    attach_settlement_pending() already runs), and it may REPLACE an
    "Okk" flag_queries() assigned with a specific pending label - never
    the reverse; an order flag_queries() already flagged for a real reason
    (partial payment, refund, RTO refund status, ...) keeps that flag
    untouched, since that's a more specific, already-correct concern.

    Only touches orders classify_order_bank_status() puts in a still-
    pending category (has_settlement_row True but not yet bank_matched -
    "collected, not yet credited" - OR has_settlement_row False for a
    prepaid gateway that's never even reported collecting it) AND whose
    current query is still "Okk". The phrase used depends on the
    order's resolved "Gateway" column (see engine.attribution -
    must already be attached before this runs):

    Client-reported 2026-08-31 (point 2): this used to re-derive "still
    pending" itself from has_settlement_row/bank_matched
    ((has_row & ~matched) | ~has_row), which wrongly caught every COD
    RTO/Cancelled/Lost order too - those never had a settlement row
    either (nothing was ever collected because the order came back), but
    for a completely different reason than "collected and awaiting bank
    credit". engine.bank.classify_order_bank_status() already tells the
    two apart (RTO/Cancelled with no settlement row -> COD_NOT_DELIVERED,
    excluded on purpose - see its own module docstring), so this now
    reuses that exact classification (_RECEIPT_PENDING_CATEGORIES) instead
    of recomputing a look-alike condition - confirmed against the
    client's own July data: 293 COD RTO orders that were incorrectly
    getting reclassified into "COD Delivered Amount not Received" (no
    resolvable Gateway on a returned order) now correctly stay "Okk".
      - a recognised COD courier (see _COD_SETTLEMENT_PENDING_PHRASES) -
        "<Courier> [Ss]etlment pending", suffixed " not reflecting" if
        nothing has reached Reco working's own receipt_amount yet
        (has_settlement_row False) or " partially reflecting" if only
        part of the order Total is reflected there yet (0 < receipt <
        total) - confirmed against the client's own workbook for both
        Delhivery COD and Shiprocket COD.
      - any other resolved gateway (e.g. "Payu") - "<Gateway> Setlment
        Pending", matching the client's own "Payu Setlment Pending" -
        no reflecting/not-reflecting variants, since every confirmed
        example of this case had nothing collected at all yet.
      - no resolvable Gateway at all - "COD Delivered Amount not
        Received", the same catch-all flag_queries() already uses
        elsewhere for "we can't even say who's holding this money".

    Disclosed gap, not guessed around: the client's own workbook also
    reclassifies a handful of orders (well under 5% of all open queries)
    in ways this function does NOT attempt to reproduce - e.g. a specific
    "Status Undefined" order relabelled by a specific delivery status
    ("...is Cancelled what is Refund status") that final_delivery_status
    itself does not record anywhere in reco_df. Those look like
    judgment calls made by directly reading a delivery partner's raw
    status text order-by-order, not a rule derivable from reco_df's own
    columns - safer to leave as whatever flag_queries() already produced
    than to guess a rule from a handful of examples and misclassify a
    different order that happens to match the same columns later. A
    related, smaller trade-off from the 2026-08-31 _RECEIPT_PENDING_
    CATEGORIES fix above: a handful of Lost/Status Undefined orders that
    never had a settlement row (nothing was ever collected) used to fall
    into this function's catch-all too, purely as a side-effect of the
    same over-broad condition that wrongly caught RTO orders - now that
    the condition is fixed, those few orders correctly stay "Okk" instead
    of the catch-all, which is more consistent with how RTO/Cancelled are
    treated even though it no longer reproduces 2-3 of the client's own
    manually-decided rows.

    Client-reported 2026-08-31 (point 2), second half: flag_queries()
    above can hand this function "Delivered but amount not received -
    reason?" instead of "Okk" for an order that turns out to be settlement
    -pending - it happens whenever receipt_amount is still 0 at the Layer-
    3 stage (before this function's bank-matching data exists), which is
    exactly the normal state of a prepaid order the gateway hasn't
    reported collecting yet (confirmed against the client's own data:
    328 of the 335 "Payu Setlment Pending" orders arrived here already
    labelled "Delivered but amount not received - reason?", not "Okk",
    and were never being touched). Both labels are candidates for the
    same reason: each is Layer 3's best diff-only guess for "money isn't
    where it should be", and this function's whole job is to replace that
    guess with the real reason once it's known - a real, already-specific
    flag_queries() concern (a refund, an RTO refund-status question, a
    genuine partial payment) is still never touched.

    Client-reported 2026-09-05 (point 2): a Status Undefined order with
    nothing received gets a THIRD candidate label from flag_queries() -
    "COD Delivery status Undifined  Amount not Received" - which this
    function previously never touched at all (it wasn't in the generic-
    labels tuple), even when the order resolves to a specific pending
    gateway. Confirmed against the client's own report: order #26366
    (Payment Provider "Payu", Status Undefined, nothing received) kept the
    COD-flavoured wording verbatim even though "Payu" isn't COD at all -
    the client's own corrected text is "Delivery status Undifined Payu
    Amount not setled". Fixed by adding this label as a fourth candidate,
    with its OWN phrasing template (not the "<Gateway> Setlment Pending"
    one used for the Delivered-style labels above) - see _label() below:
    a resolved COD-courier Gateway leaves the original COD wording
    untouched (a genuine COD order with an undetermined delivery status
    and nothing collected is, by design, excluded from "still pending"
    anyway - see the _RECEIPT_PENDING_CATEGORIES note above - so this
    branch is mostly reached by non-COD gateways in practice), a resolved
    non-COD gateway swaps in "Delivery status Undifined {Gateway} Amount
    not setled", and an unresolved Gateway leaves the original text as-is
    (nothing more specific to say).
    """
    if reco_df is None or reco_df.empty:
        return reco_df
    df = reco_df.copy()
    if recon_status_df is None or recon_status_df.empty:
        return df

    status_cols = ["order_id", "Reconciliation Category"]
    missing = [c for c in status_cols if c not in recon_status_df.columns]
    if missing:
        return df

    status = recon_status_df[status_cols].copy()
    status["order_id"] = status["order_id"].astype(str)
    status = status.drop_duplicates(subset="order_id", keep="last").set_index("order_id")

    df["order_id"] = df["order_id"].astype(str)
    category = df["order_id"].map(status["Reconciliation Category"].to_dict())

    # "Still pending" = classify_order_bank_status() already put this order
    # in one of the three "money hasn't reached the bank yet" categories -
    # see _RECEIPT_PENDING_CATEGORIES above. Only an order flag_queries()
    # left "Okk" or the generic diff-only "Delivered but amount not
    # received - reason?" gets touched - a real, already-flagged concern
    # is left alone.
    still_pending = category.isin(_RECEIPT_PENDING_CATEGORIES)
    # "COD Delivery status Undifined  Amount not Received" (client-reported
    # 2026-09-05, point 2) is a fourth candidate label - see this
    # function's own docstring above - handled with its own phrasing
    # template in _label() below rather than the generic one.
    _STATUS_UNDEFINED_NOT_RECEIVED = "COD Delivery status Undifined  Amount not Received"
    generic_labels = (
        "Okk", "Delivered but amount not received - reason?", _STATUS_UNDEFINED_NOT_RECEIVED,
    )
    candidates = df["query"].isin(generic_labels) & still_pending.fillna(False)
    if not candidates.any():
        return df

    gateway = df["Gateway"] if "Gateway" in df.columns else pd.Series(None, index=df.index)
    receipt = df["receipt_amount"].fillna(0.0)
    total = df["total"].fillna(0.0)
    orig_query = df["query"]

    def _label(idx):
        gw = gateway.loc[idx]
        gw = str(gw).strip() if pd.notna(gw) else ""
        # 2026-09-06 (round 16, order #30456): engine.attribution.
        # combine_prepaid_and_cod_label() can now hand "Gateway" a
        # COMBINED label - "<courier> COD, <prepaid provider>" (e.g.
        # "Delhivery COD, PayU") - for a genuine part-prepaid/part-COD
        # order. Before this fix, the _COD_SETTLEMENT_PENDING_PHRASES
        # lookup below only ever matched a bare, single label, so a
        # combined string fell straight to the generic "{gw} Setlment
        # Pending" branch and produced a garbled "Delhivery COD, PayU
        # Setlment Pending" Query text - the client's own follow-up ask
        # ("update the Query and Receipt Remark for this Order ID based
        # on the corrected Payment Provider"). cod_component_of_gateway_
        # label() recovers just the COD leg (always first, per round 16's
        # ordering convention) - the prepaid leg was already collected at
        # checkout, so the genuinely OUTSTANDING/pending money this Query
        # text describes is always the COD leg, never the already-settled
        # prepaid one.
        gw_cod = cod_component_of_gateway_label(gw)
        rcpt, tot = receipt.loc[idx], total.loc[idx]

        if orig_query.loc[idx] == _STATUS_UNDEFINED_NOT_RECEIVED:
            # Status Undefined's own placeholder - a different phrasing
            # template from the Delivered-style labels below (keeps the
            # "Delivery status Undifined" framing, only the amount/gateway
            # half changes).
            if gw_cod in _COD_SETTLEMENT_PENDING_PHRASES:
                # A genuine COD courier resolves here in practice - leave
                # the original COD wording untouched.
                return _STATUS_UNDEFINED_NOT_RECEIVED
            if gw and gw not in ("#N/A", "NA", "nan"):
                return f"Delivery status Undifined {gw} Amount not setled"
            return _STATUS_UNDEFINED_NOT_RECEIVED

        if gw_cod in _COD_SETTLEMENT_PENDING_PHRASES:
            base = _COD_SETTLEMENT_PENDING_PHRASES[gw_cod]
            if rcpt <= 1:
                return f"{base} not reflecting"
            if tot > 1 and rcpt < tot - 1:
                return f"{base} partially reflecting"
            return base
        if gw and gw not in ("#N/A", "NA", "nan"):
            return f"{gw} Setlment Pending"
        return "COD Delivered Amount not Received"

    df.loc[candidates, "query"] = [_label(i) for i in df.index[candidates]]
    return df


def refine_split_payment_queries(reco_df, leg_status_by_order):
    """
    Client-reported 2026-09-06 (round 17) - order #30456's own direct
    follow-up to round 16's Payment Provider fix: "yes now order id 30456
    updated correctly but query update missing now showing query Partial
    Payment Received but this case delhivery COD setled but payu
    setlment pending query to be asked 'Delhivery COD Setled, payu
    Setlment Pending'".

    refine_queries_with_settlement_status() above can't do this: it only
    ever replaces one of a fixed set of GENERIC diff-only labels ("Okk",
    "Delivered but amount not received - reason?",
    _STATUS_UNDEFINED_NOT_RECEIVED) with a single gateway's pending
    phrase. "Partial Payment Received" (flag_queries()'s own label for
    0 < receipt_amount < total) is deliberately not one of them - a
    genuine partial payment (real money genuinely short) and a genuine
    split payment (all the money is in, just via two rails with two
    different settlement timelines) look identical by receipt_amount vs
    total alone. The combined "Gateway" label (engine.attribution.
    combine_prepaid_and_cod_label()'s "<courier> COD, <prepaid provider>"
    output) is the only reliable signal telling them apart, so this is a
    separate, later pass - called only for orders whose Gateway/Payment
    Provider is actually a combined label, never touching a genuine
    single-gateway partial payment.

    leg_status_by_order: engine.bank.resolve_split_payment_leg_status()'s
    output (order_id -> {"cod_matched": bool/None, "prepaid_matched":
    bool/None}) - the row-level, per-leg bank-matching signal that
    classify_order_bank_status()'s order-level aggregate can't see (see
    that function's own docstring for the full root-cause story). A leg
    missing from this dict, or explicitly None, is treated the same as
    "not yet matched" for wording purposes - conservative by design, since
    this function should never claim a leg is "Setled" without a positive
    signal that it is.

    Only touches an order whose query is STILL exactly "Partial Payment
    Received" (flag_queries()'s own label - see that function's
    docstring) AND whose Gateway/Payment Provider resolves to a genuine
    combined label (both a COD and a prepaid component present). Any
    other query text, or a bare single-gateway label, is left completely
    untouched - a real, already-specific concern from flag_queries() or
    refine_queries_with_settlement_status() is never overwritten.
    """
    if reco_df is None or reco_df.empty or "query" not in reco_df.columns:
        return reco_df
    df = reco_df.copy()

    gateway_col = "Gateway" if "Gateway" in df.columns else (
        "Payment Provider" if "Payment Provider" in df.columns else None
    )
    if gateway_col is None:
        return df

    is_partial = df["query"] == "Partial Payment Received"
    if not is_partial.any():
        return df

    leg_status_by_order = leg_status_by_order or {}
    gateway = df[gateway_col]
    order_id = df["order_id"].astype(str)

    def _label(idx):
        gw = gateway.loc[idx]
        gw = str(gw).strip() if pd.notna(gw) else ""
        prepaid_label = prepaid_component_of_gateway_label(gw)
        if prepaid_label is None:
            # Not a genuine combined label - a real, still-short partial
            # payment on a single gateway - leave flag_queries()'s own
            # text untouched.
            return None
        cod_label = cod_component_of_gateway_label(gw)
        legs = leg_status_by_order.get(order_id.loc[idx], {})
        cod_word = "Setled" if legs.get("cod_matched") else "Setlment Pending"
        prepaid_word = "Setled" if legs.get("prepaid_matched") else "Setlment Pending"
        return f"{cod_label} {cod_word}, {prepaid_label} {prepaid_word}"

    for idx in df.index[is_partial]:
        new_label = _label(idx)
        if new_label is not None:
            df.at[idx, "query"] = new_label

    return df


def apply_cod_report_gap_query(reco_df, gateway_configs):
    """
    Client-reported 2026-09-04 (round 10, point 1, scenario 3): "COD
    Delivered but Amount Not Found in COD Report" - an order delivered by
    a COD-configured courier (Shiprocket/Delhivery/Prozo) that genuinely
    has NO row at all in that courier's own COD report - not even a
    not-yet-remitted one (engine.reco.attach_pending_cod_receipts(),
    called earlier in the pipeline, already rescues that "found but
    pending remittance" case into a receipt_amount > 0 state, which is
    why this function only ever sees the genuinely-absent case by the
    time it runs) - should read a specific, courier-named Query:
        "{Courier} COD Delivered but amount not reflecting in {Courier} COD Report"
    (e.g. "Delhivery COD Delivered but amount not reflecting in Delhivery
    COD Report") rather than the generic "{Courier} COD setlment pending
    not reflecting" / "COD Delivered Amount not Received" text
    refine_queries_with_settlement_status() above would otherwise leave in
    place for it.

    Deliberately conservative, same pattern as every other query-
    refinement pass in this module: only overrides an order that is
    STILL one of the generic "nothing received yet, still pending" labels
    at this point - a real, already-specific concern (a refund, a partial
    payment, an RTO refund-status question) is left completely untouched.

    Call this right after refine_queries_with_settlement_status() above
    (same pipeline point, both view call sites) - order relative to
    attach_receipt_status() doesn't matter, since that function only ever
    touches "receipt_status", never "query".
    """
    df = reco_df.copy()
    if df is None or df.empty:
        return df
    if "delivery_partner" not in df.columns or "final_delivery_status" not in df.columns \
            or "query" not in df.columns:
        return df

    cod_couriers = {
        str(cfg.get("label", "")).strip()[:-len(" COD")].strip()
        for cfg in (gateway_configs or [])
        if str(cfg.get("payment_mode", "")).strip().lower() == "cod"
        and str(cfg.get("label", "")).strip().upper().endswith(" COD")
    }
    if not cod_couriers:
        return df

    generic_labels = set(_COD_SETTLEMENT_PENDING_PHRASES.values())
    generic_labels |= {f"{p} not reflecting" for p in _COD_SETTLEMENT_PENDING_PHRASES.values()}
    generic_labels |= {
        "COD Delivered Amount not Received",
        "Delivered but amount not received - reason?",
    }

    receipt = df["receipt_amount"].fillna(0.0) if "receipt_amount" in df.columns else pd.Series(0.0, index=df.index)
    is_cod_courier_delivered = (
        (df["final_delivery_status"] == "Delivered")
        & df["delivery_partner"].astype(str).str.strip().isin(cod_couriers)
    )
    mask = is_cod_courier_delivered & (receipt.abs() <= 0.004) & df["query"].isin(generic_labels)
    if not mask.any():
        return df

    df.loc[mask, "query"] = df.loc[mask, "delivery_partner"].astype(str).str.strip().apply(
        lambda partner: f"{partner} COD Delivered but amount not reflecting in {partner} COD Report"
    )
    return df


def attach_receipt_status(reco_df, recon_status_df):
    """
    New "receipt_status" column, matching the client's own reference
    workbook's "Recipt Remark" column on the Reco working sheet
    (client-reported 2026-08-30 as a missing column - "Receipt Status...
    has not been added"). Placed on the Reco working sheet between "diff"/
    "Bank credit" and "Query" (see RECO_WORKING_GROUPS in
    engine/formatting.py) and renamed to "Recipt Remark" at export time
    (the same rename step in views/page_reports.py::_build_workbook() that
    already handles "settlement_amount"->"Bank credit" and "query"->
    "Query" - reco_working_cols/the internal reco_df keep the snake_case
    name "receipt_status").

    Client-reported 2026-09-05: also corrects "settlement_amount" ("Bank
    credit") for the same still-settlement-pending orders this function
    already identifies for its own "Received Bank settlement pending"
    label - see the block at the end of this function, right before the
    return, for the full rule and root-cause story.

    Rule derived by reverse-engineering the client's own filled-in values
    against reco_df's other columns (988 orders' worth of examples on the
    client's July data, cross-tabbed by final_delivery_status / refunded /
    received / fully-paid): whether the order is still settlement-pending
    dominates every other signal - an order can show receipt_amount == total
    (the courier/gateway already reports it fully collected) and still be
    "Not Received" here, because this column answers "has the BANK actually
    credited it", not "did the courier say they collected it". That is
    exactly the has_settlement_row/bank_matched signal
    refine_queries_with_settlement_status() above already uses (same
    recon_status_df, same "still pending" definition) - call this at the
    same point in the pipeline, after recon_status_df exists.

    Once settlement-pending orders are set aside, the remaining ~91% of
    orders (verified against the client's own workbook - the only
    mismatches without the pending override were exactly the
    still-pending ones) follow a simple three-way split on
    receipt_amount vs total, refined by delivery status and whether a
    refund was recorded:
      - nothing received (receipt_amount <= 1): "Not Received <status>"
        for a courier-reported status the client's own labels
        distinguish (RTO, Lost), "Not received" for Status Undefined
        (see the 2026-08-31 round-4 note below - no longer folded into
        Cancelled), "Not Received Cancelled" for Cancelled, and a generic
        "Not Received <status>" fallback for any other/unrecognised
        status.
      - something but not everything received (0 < receipt_amount <
        total - 1): "Partially received and Refunded" if a refund was
        also recorded, else "Partialy Received" (client's own spelling,
        kept verbatim rather than corrected, so this genuinely matches
        their sheet on a text diff).
      - fully received (receipt_amount >= total - 1) with a refund also
        recorded (client-reported 2026-08-31, round 4, point 3 - see that
        note below): "Refunded" when the refund consumes essentially the
        ENTIRE amount received (refund >= receipt - 1, the same
        magnitude threshold flag_queries() uses for its own "Why
        Refunded" flag), regardless of delivery status; "Partially
        Refunded" when only part of a fully-received order was later
        given back. Fully received with no refund at all is simply
        "Received" in every status.

    Client-reported 2026-08-31 (round 4, point 2 and point 3) - two
    explicit corrections that OVERRIDE judgment calls this function used
    to make, superseding the two disclosed gaps that previously stood
    here:
      - Status Undefined no longer shares Cancelled's "Not Received
        Cancelled" wording (previously disclosed as a deliberate
        simplification - this function didn't separately track which
        Status Undefined orders were actually cancelled vs something
        else). The client gave an explicit, confirmed rule instead: a
        Status Undefined order with nothing received reads "Not
        received" on its own (see flag_queries() above for the matching
        Query-side change).
      - A full refund on a fully-received Delivered order used to always
        read "Received" (the client's own workbook was observed treating
        a SMALL stray refund as noise in 12-of-14 sampled cases). Order
        28188 (refund_amount == receipt_amount == total, a complete
        reversal, not a stray adjustment) confirmed that rule was too
        broad - fully reversing what was received is a real "Refunded"
        regardless of status, while a genuinely partial give-back is its
        own new "Partially Refunded" rather than being silently absorbed
        into "Received".

    Client-reported 2026-08-31 (point 2): the "still settlement-pending"
    override above used to be re-derived from has_settlement_row/
    bank_matched directly ((has_row & ~matched) | ~has_row), which wrongly
    forced every COD RTO/Cancelled/Lost order into the generic "Not
    Received" here too - those never have a settlement row (nothing was
    ever collected, the order came back), which is a completely different
    situation from "collected but not yet bank-matched". Confirmed against
    the client's own July data: 293 COD RTO orders should read "Not
    Received RTO", not "Not Received". Fixed by reusing
    engine.bank.classify_order_bank_status()'s own "Reconciliation
    Category" (_RECEIPT_PENDING_CATEGORIES, same set
    refine_queries_with_settlement_status() above now uses) instead of
    recomputing a look-alike condition - that function already correctly
    excludes RTO/Cancelled-with-no-settlement-row into its own
    COD_NOT_DELIVERED bucket, so this now agrees with it by construction
    rather than by coincidence.
    """
    if reco_df is None or reco_df.empty:
        return reco_df
    df = reco_df.copy()

    in_pending = pd.Series(False, index=df.index)
    if recon_status_df is not None and not recon_status_df.empty:
        status_cols = ["order_id", "Reconciliation Category"]
        if all(c in recon_status_df.columns for c in status_cols):
            status = recon_status_df[status_cols].copy()
            status["order_id"] = status["order_id"].astype(str)
            status = status.drop_duplicates(subset="order_id", keep="last").set_index("order_id")
            oid = df["order_id"].astype(str)
            category = oid.map(status["Reconciliation Category"].to_dict())
            in_pending = category.isin(_RECEIPT_PENDING_CATEGORIES)
            in_pending = in_pending.fillna(False)

    receipt = df["receipt_amount"].fillna(0.0) if "receipt_amount" in df.columns else pd.Series(0.0, index=df.index)
    refund = df["refund_amount"].fillna(0.0) if "refund_amount" in df.columns else pd.Series(0.0, index=df.index)
    total = df["total"].fillna(0.0) if "total" in df.columns else pd.Series(0.0, index=df.index)
    status_col = df["final_delivery_status"] if "final_delivery_status" in df.columns else pd.Series("", index=df.index)

    refunded = refund > 1
    received = receipt > 1
    full_pay = receipt >= (total - 1)

    def _label(i):
        if in_pending.loc[i]:
            # 2026-09-04 (round 10, point 2): a COD order engine.reco.
            # attach_pending_cod_receipts() already rescued - i.e. the
            # courier's own COD report already confirms the amount, only
            # the BANK hasn't credited it yet - reads as genuinely
            # received-but-bank-pending, not "Not Received" (reserved for
            # an order nothing has been reported collecting at all yet -
            # e.g. a Payu/Gokwik order the gateway hasn't reported
            # collecting, or a COD order genuinely absent from its own COD
            # report even after that rescue - see
            # apply_cod_report_gap_query() above for that "not reflecting"
            # case, which is untouched here since receipt_amount is still
            # 0 for it).
            if received.loc[i]:
                return "Received Bank settlement pending"
            return "Not Received"
        st = status_col.loc[i]
        st = str(st).strip() if pd.notna(st) else ""
        r_recv, r_full, r_ref = received.loc[i], full_pay.loc[i], refunded.loc[i]
        if not r_recv:
            if st == "RTO":
                return "Not Received RTO"
            if st == "Lost":
                return "Not Received LOST"
            # Client-reported 2026-08-31 (round 4, point 2), overriding the
            # previously-disclosed "Status Undefined reads the same as
            # Cancelled" gap above with an explicit, confirmed rule: a
            # Status Undefined order with nothing received reads "Not
            # received" (the client's own exact casing) on its own, no
            # longer folded into Cancelled's "Not Received Cancelled".
            if st == "Status Undefined":
                return "Not received"
            if st == "Cancelled":
                return "Not Received Cancelled"
            return f"Not Received {st}".strip() if st else "Not Received"
        if not r_full:
            return "Partially received and Refunded" if r_ref else "Partialy Received"
        if r_ref:
            # Client-reported 2026-08-31 (round 4, point 3), overriding the
            # previous status-based split ("Received" for Delivered, always
            # "Refunded" otherwise): order 28188 (Delivered, fully received,
            # then refunded in FULL - refund_amount == receipt_amount ==
            # total) should read "Refunded", not "Received" - the earlier
            # rule treated every refund on a fully-received Delivered order
            # as noise, which is wrong once the refund consumes the ENTIRE
            # amount actually received, exactly the same magnitude-based
            # threshold flag_queries() already uses for its own "Delivery
            # status Delivered Why Refunded" flag on Delivered orders (see
            # that function above). Generalised here across every status,
            # not just Delivered: a refund that reverses essentially all of
            # what was received is "Refunded"; a refund that only takes
            # back PART of a fully-received order (a discount/adjustment,
            # some money still genuinely retained) is now its own distinct
            # "Partially Refunded" - the client explicitly asked for a
            # partial refund to surface as its own status rather than be
            # silently absorbed into "Received".
            return "Refunded" if refund.loc[i] >= receipt.loc[i] - 1 else "Partially Refunded"
        return "Received"

    df["receipt_status"] = [_label(i) for i in df.index]

    # Client-reported 2026-09-05: the Shiprocket-COD-pending receipt_amount
    # rescue above (attach_pending_cod_receipts(), 2026-09-04/round 10)
    # correctly fixed receipt_amount, but exposed a second bug - the SAME
    # rescued amount was also leaking into "settlement_amount" ("Bank
    # credit" at export), even though the bank hasn't actually credited it
    # yet. Root cause: settlement_amount is a pure arithmetic derivative of
    # receipt_amount (attach_receipts_and_diff()/attach_pending_cod_
    # receipts() above: receipt_amount - total_deduction - refund_amount)
    # with no awareness of whether the money has actually reached the bank
    # - that was safe before the rescue only because normalize_gateway_df()
    # 's settled_status_col filter meant receipt_amount was NEVER populated
    # until a gateway/courier report already showed the money as remitted,
    # so "has a receipt" and "bank-credited-ish" happened to coincide. The
    # rescue deliberately broke that coincidence for genuinely-still-
    # pending COD orders (so receipt_amount/Recipt Remark could correctly
    # show "received, bank pending" - see the branch just above), without
    # ever correcting settlement_amount for the same orders.
    #
    # Client's own explicit rule (2026-09-05):
    #   1. COD amount reflecting in the COD report AND credited to bank -
    #      both receipt_amount and Bank credit.
    #   2. COD amount reflecting in the COD report but settlement/bank-
    #      credit pending - receipt_amount only; Bank credit must stay 0.
    #   3. COD amount not reflecting in the COD report at all - neither
    #      column (already true unconditionally - receipt_amount was never
    #      populated for these).
    #
    # `in_pending` above (computed from engine.bank.classify_order_bank_
    # status()'s own "Reconciliation Category" - the exact same signal two
    # branches above already use to choose "Received Bank settlement
    # pending" over "Received") is precisely "has the BANK actually
    # credited this order" == False - not a new heuristic, and not
    # specific to the Shiprocket-COD rescue: it's the general not-yet-
    # bank-matched bucket (COD_SETTLEMENT_PENDING / PREPAID_SETTLEMENT_
    # PENDING / EXCEPTION_MANUAL_REVIEW), so this also correctly keeps Bank
    # credit at 0 for e.g. a Prepaid order whose settlement row exists but
    # hasn't been bank-matched yet - a case the pre-existing settlement_
    # pending_amount-driven zeroing in views/page_reports.py::
    # _build_workbook() already handled correctly (has_settlement_row True
    # there), so this doesn't change behaviour for it, just fixes it one
    # layer earlier and additionally covers the has_settlement_row-False
    # rescued case that older mechanism couldn't see (recon_status_df's own
    # has_settlement_row/receipt_amount never learn about the rescue - see
    # attach_pending_cod_receipts()'s own docstring). That downstream
    # zeroing step, and engine/reco.py::attach_settlement_pending()/
    # engine/summary.py::headline_totals()'s own "Settlement pending"
    # netting, are left exactly as they are - now simply redundant-but-
    # harmless for every order this function already zeroes here (Bank
    # credit is already 0, so "zero it again" / "subtract 0 more" are both
    # no-ops). Nothing about the previously-verified bank-matching/grace-
    # period/exception classification logic itself changes - only this
    # already-derived "Bank credit" figure.
    if "settlement_amount" in df.columns:
        df["settlement_amount"] = df["settlement_amount"].where(~in_pending, 0.0)

    return df


def refresh_delivery_status(reco_df, delivery_frames, delivery_configs):
    """
    Patches an ALREADY-BUILT reco_df (typically one loaded back from a
    previously-SAVED period) with updated delivery status for whichever
    orders actually appear in the freshly-uploaded delivery_frames -
    leaving every other order's already-recorded status completely
    untouched - per the client's rule that a later delivery-partner report
    showing an updated status (e.g. Undelivered -> Delivered) must be
    used instead of ignored as a duplicate.

    Deliberately NOT the same as calling attach_delivery_status(reco_df,
    delivery_frames, ...) directly: that function resets EVERY order to
    "Status Undefined" before re-deriving status from delivery_frames,
    which is correct for a brand-new run (delivery_frames covers the
    whole order set) but wrong here, where delivery_frames is often just
    the newest incremental file covering a handful of orders - applying
    it directly would wipe out the correct, already-recorded status of
    every order NOT in this particular file.

    Only requires reco_df to have an "order_id" column - the reco_df
    saved by engine.storage.save_run() always does, since it's built on
    top of build_order_master()'s output.
    """
    shell = pd.DataFrame({"order_id": reco_df["order_id"]})
    refreshed = attach_delivery_status(shell, delivery_frames, delivery_configs)

    # Only orders that actually matched something in delivery_frames come
    # back with a real status - everything else is still the placeholder
    # "Status Undefined" attach_delivery_status starts every order at,
    # which must NOT be applied (that would overwrite a perfectly good
    # existing status with "unknown" just because this particular file
    # doesn't happen to mention that order).
    touched_mask = refreshed["final_delivery_status"] != "Status Undefined"
    if not touched_mask.any():
        return reco_df, 0

    patch_cols = ["delivery_partner", "final_delivery_status", "delivered_date", "rto_date"]
    patch_cols += [c for c in refreshed.columns if c.endswith("_raw_status")]

    touched_order_ids = set(refreshed.loc[touched_mask, "order_id"])
    out = reco_df.copy()
    row_mask = out["order_id"].isin(touched_order_ids)

    lookup_df = refreshed.set_index("order_id")
    for col in patch_cols:
        if col not in out.columns:
            out[col] = None
        if col in lookup_df.columns:
            out.loc[row_mask, col] = out.loc[row_mask, "order_id"].map(lookup_df[col])

    out = flag_queries(out)
    return out, int(row_mask.sum())


def run_shopify_pipeline(orders_df, delivery_frames, receipt_summary, config):
    """
    Runs all of Layer 3 in order. This is the single function app.py calls
    for the financial Reco working table.
    """
    order_master = build_order_master(orders_df, config["orders"])
    with_delivery = attach_delivery_status(order_master, delivery_frames, config["delivery_partners"])
    with_receipts = attach_receipts_and_diff(with_delivery, receipt_summary)
    final = flag_queries(with_receipts)
    return final
