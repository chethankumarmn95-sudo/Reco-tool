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

    # Now decide the final status: walk partners in priority order, and for
    # each order, use the first partner whose raw status actually classifies
    # to a known bucket - falling through to the next partner if it doesn't.
    def resolve_row(order_id):
        for cfg in delivery_configs:
            label = cfg["label"]
            lookup = partner_status_lookups.get(label)
            if lookup is None or order_id not in lookup.index:
                continue
            raw = lookup[order_id]
            classified = classify_status(raw)
            if classified is not None:
                display_label = label
                # Prefer this source's own real-courier column (e.g.
                # Unicommerce's "Shipping provider") over its generic label,
                # falling back to the label when the column is missing/blank
                # for this order so behavior degrades safely.
                courier_lookup = partner_courier_label_lookups.get(label)
                if courier_lookup is not None and order_id in courier_lookup.index:
                    courier_name = courier_lookup[order_id]
                    if pd.notna(courier_name) and str(courier_name).strip():
                        display_label = str(courier_name).strip()
                return display_label, classified
        # No partner gave a classifiable status - but if ANY partner at
        # least had the order, say so rather than a bare "Undefined".
        for cfg in delivery_configs:
            label = cfg["label"]
            lookup = partner_status_lookups.get(label)
            if lookup is not None and order_id in lookup.index:
                return label, "Status Undefined"
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


def attach_settlement_pending(reco_df, settlement_pending_df):
    """
    Merges each order's outstanding Settlement Pending amount (engine/
    settlement_pending.py::build_settlement_pending_report()'s own
    "Settlement Amount" column - already the correct net-of-deduction/
    refund figure for orders with a settlement row, or the full order
    Total for orders with none yet - see that function's docstring) onto
    reco_df as "settlement_pending_amount", so engine/summary.py's
    headline_totals() can subtract it from the raw settlement_amount total
    to get a "Net settlement" that reflects only money that has actually
    reached the bank (client-reported 2026-08-27: Net Settlement should
    ultimately tie to the Bank Credit amount, and didn't, because it was
    counting still-pending money as already settled - see
    headline_totals()'s own docstring for the full story).

    Client-reported again 2026-08-30, against a fresh July-only run: Net
    Settlement was STILL wrong (too low by ~₹3.2 lakh on the client's own
    data) even after the fix above, because this originally summed
    build_settlement_pending_report()'s "Settlement Amount" column
    UNFILTERED - which deliberately includes two different kinds of
    "pending" money (see that function's own has_settlement_row branch):
    (a) orders the gateway/courier has already reported money received
    for but that hasn't reached the bank yet (Gateway Amount > 0 - this
    money IS already inside reco_df's own receipt_amount, and therefore
    inside "Receipt before deduction" above, so it's exactly what needs
    subtracting to avoid double-counting it as settled), and (b) orders
    nothing has been collected for AT ALL yet (Gateway Amount == 0, has_
    settlement_row False - e.g. a Payu/Gokwik order still awaiting its
    first settlement file) - whose "Settlement Amount" is that function's
    fallback to the order's full gross Total, precisely because there's no
    receipt yet to net against. That figure was NEVER added into receipt_
    amount/"Receipt before deduction" in the first place, so subtracting
    it here double-subtracted money that was never counted as received -
    confirmed against the client's own corrected workbook, where COD
    partners awaiting bank credit (money already collected, e.g. Delhivery
    COD/Shiprocket COD settlement-pending amounts) reduce Net Settlement,
    but Payu/Gokwik orders with nothing collected yet do not. Fix: only
    orders with a settlement row already (Gateway Amount > 0 - i.e.
    genuinely "collected but not yet bank-credited") are summed into
    settlement_pending_amount; "nothing collected yet" orders correctly
    stay visible in the Settlement Pending Detail/Summary sheets (used for
    exception-chasing, not touched by this filter) but no longer reduce
    Net Settlement.

    Deliberately NOT part of run_shopify_pipeline() below: settlement
    pending can only be computed once bank matching / reconciliation
    categories exist (engine.bank.classify_order_bank_status(), consumed by
    build_settlement_pending_report()), which happens later in
    views/page_reconciliation.py and views/page_reports.py than Layer 3
    runs - both call this explicitly, right after they compute
    settlement_pending_df.

    Only orders build_settlement_pending_report() actually lists (COD/
    Prepaid orders still Settlement Pending or flagged an Exception - see
    that function's PENDING_CATEGORIES) AND that already have a Gateway
    Amount (money genuinely collected, awaiting bank credit) get a
    non-zero value; every other order's settlement_pending_amount is 0 -
    either its settlement_amount is already fully realized, or nothing has
    been collected for it yet, so there's nothing sitting inside receipt_
    amount that needs netting back out.

    settlement_pending_df may legitimately be None/empty (e.g. no bank
    statement uploaded this run, or nothing is currently pending) - reco_df
    still gets the column, just filled with 0.0 everywhere, so downstream
    code (headline_totals(), the Reco working export) can always rely on
    the column being present rather than checking for it every time.
    """
    df = reco_df.copy()
    df["settlement_pending_amount"] = 0.0
    if settlement_pending_df is None or settlement_pending_df.empty:
        return df

    pending = settlement_pending_df[["Order ID", "Settlement Amount", "Gateway Amount"]].copy()
    pending["Order ID"] = pending["Order ID"].astype(str)
    # Only money already collected by the gateway/courier (and therefore
    # already inside reco_df's receipt_amount / "Receipt before deduction")
    # but not yet bank-credited counts against Net Settlement - see the
    # docstring above. A blank/zero Gateway Amount means nothing has been
    # collected yet; that order's full-Total "Settlement Amount" fallback
    # must NOT reduce Net Settlement, or money never counted as received
    # gets subtracted from it anyway.
    pending["Gateway Amount"] = pending["Gateway Amount"].fillna(0.0)
    pending = pending[pending["Gateway Amount"] > 0.01]
    pending_lookup = pending.groupby("Order ID")["Settlement Amount"].sum()

    df["order_id"] = df["order_id"].astype(str)
    matched = df["order_id"].isin(pending_lookup.index)
    # .to_dict() first, not .map()'d straight off the Series - same fix as
    # attach_delivery_status()'s date lookups above and
    # engine/summary.py's month_summary(): sidesteps a pandas dtype-
    # inference crash when the mapper Series is empty.
    df.loc[matched, "settlement_pending_amount"] = df.loc[matched, "order_id"].map(pending_lookup.to_dict())
    df["settlement_pending_amount"] = df["settlement_pending_amount"].fillna(0.0)
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
            return "Okk"
        if status in ("RTO", "Cancelled", "Lost", "Refunded"):
            if partial_receipt and refund >= receipt - 1:
                # Whatever partial amount came in went straight back out -
                # confirmed across RTO/Cancelled/Status Undefined orders in
                # the client's own data, not status-specific.
                return "Why Partially received and same amount refunded"
            if status in ("RTO", "Lost") and received and not refunded:
                # For COD, nothing is collected until delivery, so RTO/Lost
                # simply means "never collected" - normal, not an exception
                # - UNLESS something WAS actually received and hasn't been
                # refunded, which is the real anomaly worth flagging.
                # Cancelled/Refunded do NOT get this same flag even when
                # received-and-not-refunded (confirmed against the client's
                # own data - see docstring above).
                if status == "Lost":
                    return "Shipment LOST - Refund/Credit note status required?"
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
    generic_labels = ("Okk", "Delivered but amount not received - reason?")
    candidates = df["query"].isin(generic_labels) & still_pending.fillna(False)
    if not candidates.any():
        return df

    gateway = df["Gateway"] if "Gateway" in df.columns else pd.Series(None, index=df.index)
    receipt = df["receipt_amount"].fillna(0.0)
    total = df["total"].fillna(0.0)

    def _label(idx):
        gw = gateway.loc[idx]
        gw = str(gw).strip() if pd.notna(gw) else ""
        rcpt, tot = receipt.loc[idx], total.loc[idx]
        if gw in _COD_SETTLEMENT_PENDING_PHRASES:
            base = _COD_SETTLEMENT_PENDING_PHRASES[gw]
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
