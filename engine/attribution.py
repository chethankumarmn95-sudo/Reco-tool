"""
attribution.py
---------------
Optional "which specific rail actually moved this money" layer, on top
of engine.consolidator's own per-gateway settlement figures - purely
informational (never changes any settlement $ math), feeding the Bank
Reco (UTR-wise) sheet's "Payment Gateway" column and the Reco working
sheet's own "Gateway" column (2026-08-25 client request).

Two distinct signals feed this:
  1. Which CONFIGURED gateway/COD source actually reported this
     UTR/order at all (Delhivery COD, Shiprocket COD, Prozo COD,
     Razorpay, Gokwik, ...) - always available for free from
     engine.consolidator's own consolidated_df["source"] column, no
     extra upload needed. Correct and unambiguous on its own for every
     source EXCEPT Gokwik.
  2. Gokwik itself is a checkout AGGREGATOR, not the rail that actually
     moves the money - it routes each transaction through one of
     several downstream payment processors (easebuzz, payu, ...), and
     the client wants THAT specific processor shown, not the bare label
     "Gokwik". Gokwik's own settlement file (the existing "Gokwik"
     gateway upload, unchanged by any of this) has no column naming the
     downstream processor at all - that only shows up in two SEPARATE
     reports Gokwik provides: an "Order Report" (Shopify order <->
     Gokwik's own Payment ID) and a "Transaction Report" (Payment ID
     <-> downstream Payment Provider). Joining those two - see
     build_gokwik_payment_provider_map() below - gives an order_id ->
     downstream-provider mapping that refines "Gokwik" into "easebuzz"/
     "payu"/etc wherever a Shopify order_id is resolvable.

Both new reports are entirely OPTIONAL (mandatory: false in config) -
without them, Gokwik-sourced money simply shows as "Gokwik" (the
gateway's own label), still correct, just less specific. Known
residual gap, disclosed rather than guessed around: a Gokwik-sourced
settlement-ledger row whose order_id could never be resolved at all
(engine.bank.NOT_FOUND) has no order to join the Order Report against
either, so it keeps the plain "Gokwik" label even when both reports are
uploaded - there is no order-independent bank-matchable field in
either Gokwik report today to refine it further (neither report
carries a UTR or a bank reference of any kind).
"""

import pandas as pd

from .loaders import normalize_order_id, resolve_col, resolve_col_or_raise


def attach_payment_columns(reco_df, attribution_frames, attribution_sources_cfg, gokwik_provider_map=None,
                            direct_txn_provider_map=None, recon_status_df=None, gateway_configs=None,
                            consolidated_df=None):
    """
    Adds "Payment Method" and "Payment Provider" columns to reco_df
    (client requirement, 2026-08-27 - see Reco working sheet spec), sourced
    from the same optional Gokwik Order Report / Gokwik Transaction Report
    uploads this module already reads (views/page_upload.py's "Payment
    Gateway Attribution" section):

      - Payment Provider reuses build_gokwik_payment_provider_map()'s own
        order_id -> downstream-processor mapping as-is (gokwik_provider_map
        param, already built once per run and shared with
        build_payment_gateway_lookups() below - not recomputed here).
      - Payment Method is a NEW per-order lookup off the Gokwik Order
        Report's own payment_method_col (config key added alongside this
        function - e.g. "Payment Method"/"Payment Mode"), resolved via
        resolve_col (never resolve_col_or_raise: this source is
        `mandatory: false`, so a missing report or a renamed-beyond-
        recognition column must degrade to blank, not error the whole run).

    Both columns fall back to whatever reco_df already carries (the Shopify
    order report's own "payment_method" field, and, when present, the
    "Gateway" column - see build_payment_gateway_lookups()) whenever Gokwik
    attribution can't resolve a given order - "less specific, not wrong",
    the same fallback philosophy this whole module already uses (see the
    module docstring above).

    Client-reported 2026-09-04 (round 9): a Report re-generated for an
    already-saved period showed "Payment Method"/"Payment Provider" back
    to the generic Gokwik/blank fallback for THOUSANDS of orders that had
    already resolved correctly (to "easebuzz"/"payu"/"UPIIntent"/...) in
    an earlier export of the exact same July period - traced to
    attribution_frames being whatever the Gokwik Order/Transaction Report
    upload slots CURRENTLY hold in this live session (see views/
    page_upload.py), not something saved alongside the period the way
    consolidated_df/bank_ledger_df already are (engine.storage.save_run()).
    Re-running and re-saving a period for an unrelated fix (e.g. adding
    the Shiprocket COD file) without ALSO having the Gokwik reports live
    that same session silently overwrites the period's own saved reco_df
    with blank/generic attribution - permanently, since save_run() persists
    exactly the reco_df it's handed. Fixed here by adding the reco_df's
    OWN pre-existing "Payment Method"/"Payment Provider" values (whatever
    a previous, correctly-attributed run already baked in) as a fallback
    layer BETWEEN the live Gokwik lookup and the least-specific raw-
    Shopify-field fallback - so regenerating a Report can only ever add
    information this session's live uploads provide, never erase
    information a previous run already resolved. A brand-new live run
    (views/page_reconciliation.py) has no such column yet at this point in
    the pipeline, so this fallback is simply inert (all blank) there -
    identical behaviour to before this fix for that path.

    Client-reported 2026-08-31 (round 8): "Payment Provider" was blank for
    every COD order the "Gateway" column itself has no value for - i.e.
    every COD order that has no consolidated_df row yet (no COD courier's
    settlement/remittance file has reported this order_id at all, so
    build_payment_gateway_lookups() has nothing to key off), even when the
    order's own delivery_partner is already known (a real courier already
    delivered it - see engine.reco.attach_delivery_status). The client's
    own reference workbook still shows "Delhivery COD"/"Shiprocket COD"
    here (confirmed against orders 31677/32318 in her July data - both
    Delhivery-delivered, receipt_amount still 0, Query still the generic
    "COD Delivered Amount not Received", yet Payment Provider correctly
    reads "Delhivery COD") - i.e. Payment Provider for a COD order can be
    derived from WHICH COURIER delivered it, independent of whether that
    courier's settlement file has reported it yet. Only the Payment
    Provider *label* is affected by this fix - the Query/receipt_status
    columns are untouched (confirmed against the same two orders: MOD
    keeps them exactly as flag_queries()/refine_queries_with_settlement_
    status() already produce).

    Client-reported 2026-09-04 (round 10, point 1): two more changes,
    both in the "Payment Provider" block specifically -
      1. direct_txn_provider_map (see build_direct_transaction_provider_
         map() below) - built off the Gokwik Transaction Report's own
         Platform-order-number column, when present - now wins over
         gokwik_provider_map wherever both resolve. Order #28980 showed
         "easebuzz" (the two-report Order-Report+Payment-ID join's
         answer) when the Transaction Report itself plainly lists Payment
         Provider "payu" against that order's own Payment ID/Platform
         order number - the direct, single-file lookup is strictly more
         trustworthy than a two-file join that has already needed several
         rounds of fixes for exactly this failure mode (stale/duplicate
         Payment IDs on either side - see _collapse_unambiguous/
         build_gokwik_payment_provider_map). Optional and additive: an
         export without a recognisable order-number column on the
         Transaction Report (direct_txn_provider_map empty/None) leaves
         this whole block exactly as before.
      2. A confirmed-COD order (see resolve_cod_courier_label below) that
         ALSO already has a genuine prepaid-processor Payment Provider
         resolved (from either map above) is no longer simply overwritten
         by the COD label - see combine_prepaid_and_cod_label() below.
         Order #31956 (Gokwik PPCOD - a real part-prepaid/part-COD order:
         ₹95.67 paid via easebuzz at checkout, the ₹890 balance collected
         COD on delivery through Delhivery) now shows "easebuzz, Delhivery
         COD" - both legs - rather than only the COD leg as before.

    New optional params, both required together for the COD-courier
    override below to run (silently skipped, not an error, if either is
    omitted - same degrade-gracefully philosophy as the rest of this
    module):
      recon_status_df: engine.bank.classify_order_bank_status()'s own
        output - specifically its "payment_type" column ("COD"/"Prepaid"/
        "Unknown"), the same already-validated COD/Prepaid signal
        refine_queries_with_settlement_status() and the Settlement Pending
        Summary already rely on (see classify_payment_type()'s own
        docstring for why a courier-COD settlement row always outranks a
        blank/ambiguous Payment Method field for this call's real data).
      gateway_configs: config["gateways"] - used only to confirm
        f"{delivery_partner} COD" is actually a configured COD gateway
        label (payment_mode == "COD") before using it, so an unrecognised
        delivery_partner (e.g. "Self Delivery / Manual Report", which has
        no matching COD gateway entry) never produces a made-up label.

    consolidated_df (2026-09-05, optional): engine.consolidator's own
      settlement ledger - passed through to has_genuine_prepaid_receipt()
      so the COD-courier override's combine-vs-replace decision (see
      combine_prepaid_and_cod_label()'s own docstring, order #27533) can
      tell a genuine part-prepaid leg from a phantom Gokwik-attribution
      hit with no real money behind it. Omitting it degrades safely to
      combine_prepaid_and_cod_label()'s own pre-2026-09-05 behaviour
      (treat any prepaid-looking label as genuine) - same graceful-
      degrade philosophy as every other optional param here. Also (added
      same day, order #26164) passed to resolve_cod_courier_label() itself
      so IT can resolve the COD courier label from the actual settlement
      source rather than only guessing from delivery_partner - see that
      function's own docstring.

    Returns a NEW dataframe (reco_df.copy()) with the two columns added -
    never mutates the input in place.
    """
    df = reco_df.copy()
    df["order_id"] = df["order_id"].astype(str)

    # --- Payment Method: Gokwik Order Report's own payment_method_col ----
    method_by_order = {}
    if attribution_frames and attribution_sources_cfg:
        for cfg in attribution_sources_cfg:
            if cfg.get("role") != "order_report" or not cfg.get("payment_method_col"):
                continue
            order_df = attribution_frames.get(cfg.get("label"))
            if order_df is None or order_df.empty:
                continue
            order_id_col = resolve_col(order_df, cfg.get("order_id_col"))
            method_col = resolve_col(order_df, cfg.get("payment_method_col"))
            if not order_id_col or not method_col:
                continue  # degrade to blank for this source, never raise
            tmp = pd.DataFrame({
                "order_id": normalize_order_id(order_df[order_id_col]),
                "_method": order_df[method_col].astype(str).str.strip(),
            })
            tmp = tmp[(tmp["order_id"].str.len() > 0) & (tmp["_method"].str.len() > 0)]
            method_by_order.update(dict(zip(tmp["order_id"], tmp["_method"])))

    # Fallback chain, highest to lowest specificity: this session's live
    # Gokwik Order Report lookup -> whatever "Payment Method" reco_df
    # already carried coming IN (a previous, correctly-attributed run's
    # own saved value - see the 2026-09-04 fix note above; absent on a
    # brand-new live run, so a no-op there) -> the Shopify order report's
    # own raw "payment_method" field (least specific of all).
    prior_method = df["Payment Method"] if "Payment Method" in df.columns else pd.Series(None, index=df.index)
    fallback_method = df["payment_method"] if "payment_method" in df.columns else pd.Series(None, index=df.index)
    df["Payment Method"] = df["order_id"].map(method_by_order)
    df["Payment Method"] = df["Payment Method"].where(df["Payment Method"].notna(), prior_method)
    df["Payment Method"] = df["Payment Method"].where(df["Payment Method"].notna(), fallback_method)

    # --- Payment Provider: Gokwik order_id -> downstream provider map -----
    # direct_txn_provider_map (see build_direct_transaction_provider_map()
    # below) is applied AFTER gokwik_provider_map so it wins wherever both
    # resolve - the direct, single-file Transaction-Report lookup is more
    # trustworthy than the two-report join (2026-09-04, round 10 - see the
    # docstring above).
    provider_by_order = {}
    if gokwik_provider_map is not None and not gokwik_provider_map.empty:
        provider_by_order = dict(zip(
            gokwik_provider_map["order_id"].astype(str), gokwik_provider_map["payment_provider"],
        ))
    if direct_txn_provider_map is not None and not direct_txn_provider_map.empty:
        provider_by_order.update(dict(zip(
            direct_txn_provider_map["order_id"].astype(str), direct_txn_provider_map["payment_provider"],
        )))

    # Same fallback chain as Payment Method above, plus the "Gateway"
    # column (consolidated_df-sourced, reliable/reliable regardless of this
    # session's live attribution uploads - see build_payment_gateway_
    # lookups()) ahead of the Shopify-raw fallback.
    prior_provider = df["Payment Provider"] if "Payment Provider" in df.columns else pd.Series(None, index=df.index)
    fallback_provider = df["Gateway"] if "Gateway" in df.columns else pd.Series(None, index=df.index)
    df["Payment Provider"] = df["order_id"].map(provider_by_order)
    df["Payment Provider"] = df["Payment Provider"].where(df["Payment Provider"].notna(), prior_provider)
    df["Payment Provider"] = df["Payment Provider"].where(df["Payment Provider"].notna(), fallback_provider)

    # --- Payment Provider: confirmed-COD orders ALWAYS show "<courier>
    # COD" (2026-08-31 round 8 fix, widened 2026-09-04 round 9) ----------
    # Round 8: this used to be a last-resort fallback, only applied where
    # Payment Provider was still blank after the two lookups above.
    # Round 9 (client-reported, orders #27533/28062/28211/28219/28569/
    # #28582): that wasn't enough - all six are genuinely COD orders
    # (payment_type == "COD", correctly Shiprocket-delivered, nothing ever
    # collected through any payment gateway - receipt_amount 0) that STILL
    # showed a specific-looking but WRONG "easebuzz" Payment Provider,
    # because the Gokwik Order/Transaction Report happened to carry a row
    # for that same order_id too - most plausibly an abandoned/failed
    # Gokwik checkout attempt before the order was actually placed and
    # fulfilled COD instead. A real courier-confirmed COD delivery is
    # stronger evidence of how an order was actually paid for than an
    # unrelated payment-gateway log entry, so for a confirmed-COD order
    # this now OVERRIDES the Gokwik-sourced value outright rather than
    # only filling in a blank - reusing the exact same payment_type/
    # gateway_configs signal as before (see the docstring above), just
    # applied with higher priority. An order recon_status_df hasn't
    # classified as COD (Prepaid/Unknown, or recon_status_df/
    # gateway_configs simply weren't passed) is completely untouched by
    # this block - only ever narrows towards a MORE specific, more
    # confirmed answer for an order already known to be COD, never guesses
    # for one that isn't.
    # 2026-09-04 (round 10): a confirmed-COD order that ALSO already has a
    # genuine prepaid-processor Payment Provider (order #31956 - a real
    # part-prepaid/part-COD "Gokwik PPCOD" order) now shows BOTH legs
    # combined ("easebuzz, Delhivery COD") rather than the COD label
    # silently replacing the prepaid one - see combine_prepaid_and_cod_
    # label()'s own docstring.
    # 2026-09-05: that combine step regressed round 9's own fix for order
    # #27533 and its siblings (a phantom Gokwik-attribution hit, no real
    # money) - now gated on has_genuine_prepaid_receipt() so only a
    # REAL prepaid leg gets combined; a phantom one still gets the plain
    # "<courier> COD" replace (round 9's original, correct behaviour).
    # Only applied when consolidated_df is actually supplied - an EMPTY
    # verified-set (consolidated_df omitted) must NOT be passed to
    # combine_prepaid_and_cod_label() as if it were a real "nothing here
    # is genuine" answer, or it would block every genuine combine too
    # (caught by this module's own isolated unit tests, which exercise
    # attach_payment_columns() without a consolidated_df at all) - pass
    # None instead, which combine_prepaid_and_cod_label() already treats
    # as "can't verify, fall back to the pre-2026-09-05 behaviour".
    # 2026-09-06 (round 16, order #30456): genuine_prepaid_ids is now the
    # UNION of has_genuine_prepaid_receipt() (a raw settlement-ledger row)
    # AND has_confirmed_prepaid_transaction() (a confirmed-successful row
    # in the Gokwik Transaction Report itself) - see resolve_prepaid_
    # evidence()'s own docstring for why the settlement-ledger check alone
    # missed a genuine part-prepaid/part-COD order whose prepaid leg's own
    # settlement file hadn't posted yet this run.
    cod_override = resolve_cod_courier_label(df, recon_status_df, gateway_configs, consolidated_df=consolidated_df)
    genuine_prepaid_ids = resolve_prepaid_evidence(
        df["order_id"], consolidated_df=consolidated_df, gateway_configs=gateway_configs,
        attribution_frames=attribution_frames, attribution_sources_cfg=attribution_sources_cfg,
    )
    df["Payment Provider"] = combine_prepaid_and_cod_label(
        df["Payment Provider"], cod_override,
        order_id_series=df["order_id"], has_real_prepaid_receipt=genuine_prepaid_ids,
    )

    return df


def resolve_cod_courier_label(df, recon_status_df, gateway_configs, consolidated_df=None):
    """
    Shared "<courier> COD" resolver (2026-09-04) - the exact same rule
    attach_payment_columns() above applies to the "Payment Provider"
    column, factored out here so views/page_reconciliation.py and views/
    page_reports.py can apply the identical override to the "Gateway"
    column too (see their own call sites). Both columns need to agree -
    refine_queries_with_settlement_status() keys its query text off
    "Gateway", so leaving Gateway on a stale/wrong "easebuzz" while
    Payment Provider correctly says "Shiprocket COD" is exactly what
    produced the client-reported "easebuzz Setlment Pending" Query text on
    six genuinely-COD, Shiprocket-delivered orders (#27533 and siblings)
    whose Payment Provider was already showing the right courier.

    Client-reported 2026-09-05 (order #26164): the courier label used to
    be derived ENTIRELY from delivery_partner (engine.reco.attach_
    delivery_status()'s own resolution, sourced from each courier's
    DELIVERY-STATUS tracking report) + " COD" - silently assuming the
    courier tracking this order's delivery status is the SAME courier
    whose own COD SETTLEMENT/remittance file actually collected the
    money. Those are two independent uploads/signals that aren't
    guaranteed to agree (a courier can settle an order's COD payment
    without that same courier's own status feed happening to cover this
    particular order, or vice versa). Order #26164 - Payment Method COD,
    genuinely matched in Delhivery's own COD settlement report - kept a
    phantom "easebuzz" Payment Provider (an unrelated Gokwik attribution-
    report hit) because delivery_partner for this order didn't resolve to
    exactly "Delhivery COD"'s own courier name that run, so this function
    returned NaN and the COD override never fired at all - the exact
    "should not take assumption basis" gap the client called out.

    Fixed: now resolves the courier label from the FACTUAL settlement
    source first - which COD-configured gateway's own settlement/
    remittance file (consolidated_df, the same ledger engine.bank.
    classify_payment_type()'s has_cod_settlement_row signal already reads)
    actually has a non-refund row for this order_id - never a guess from a
    different report. Only falls back to the old delivery_partner-based
    guess for a row the factual signal above couldn't resolve (consolidated
    _df omitted, or this COD order's settlement file genuinely hasn't been
    uploaded/matched yet this run) - identical behaviour to before for
    every scenario that already worked correctly.

    Client-reported 2026-09-06 (round 16, order #26105): an RTO'd order -
    never delivered, so no COD settlement/remittance row is ever expected
    (nothing was ever collected) - had delivery_partner "Delhivery D2C"
    but Payment Provider stayed completely BLANK, when it should
    auto-resolve to "Delhivery COD". Root cause: the delivery_partner
    fallback used to require an EXACT match - f"{delivery_partner} COD"
    had to equal a configured COD gateway label character-for-character.
    A delivery-status tracking report can legitimately name a courier
    slightly differently from that same courier's own COD gateway config
    label (e.g. "Delhivery D2C" vs the configured "Delhivery COD" - a
    service-line variant of the same courier, not a different one), so
    the exact-match concatenation silently failed and this function
    returned NaN. Fixed via _match_cod_courier_label() below: the
    delivery_partner text is matched by COURIER-FAMILY PREFIX against
    each configured COD gateway label (that label with its own trailing
    " COD" stripped) - "Delhivery D2C" starts with "Delhivery", the
    prefix of configured label "Delhivery COD", so it now resolves
    correctly. An exact match (the pre-existing behaviour) is still tried
    first and always wins when it applies; a genuinely unrecognised
    delivery_partner (no configured COD label's prefix matches at all,
    e.g. "Self Delivery / Manual Report" or "partner undifined") still
    correctly returns NaN rather than guessing.

    Returns a pandas Series (df's own index) of the candidate label
    ("Delhivery COD", "Shiprocket COD", ...) wherever recon_status_df has
    classified the order as "COD" AND a label could be resolved (from the
    factual settlement source, or the delivery_partner fallback) that is
    an actually-configured COD gateway label - NaN everywhere else
    (Prepaid/Unknown orders, or nothing resolvable at all, e.g. "Self
    Delivery / Manual Report" or "partner undifined" with no matching COD
    settlement row either). Callers decide how to combine this with their
    own existing value - both current call sites use it as an override
    for a confirmed-COD order, never for anything recon_status_df hasn't
    classified as COD.
    """
    empty = pd.Series(None, index=df.index, dtype=object)
    if not gateway_configs:
        return empty
    cod_gateway_labels = {
        str(cfg.get("label", "")).strip()
        for cfg in gateway_configs
        if str(cfg.get("payment_mode", "")).strip().lower() == "cod"
    }
    if not cod_gateway_labels or recon_status_df is None or recon_status_df.empty \
            or "payment_type" not in recon_status_df.columns:
        return empty
    payment_type_by_order = dict(zip(
        recon_status_df["order_id"].astype(str), recon_status_df["payment_type"],
    ))
    is_cod = df["order_id"].astype(str).map(payment_type_by_order).eq("COD")
    if not is_cod.any():
        return empty

    order_id_text = df["order_id"].astype(str)

    # Primary signal: the actual COD gateway's own settlement/remittance
    # file that reports this order - a fact, not a guess.
    from_settlement = pd.Series(None, index=df.index, dtype=object)
    if (consolidated_df is not None and not consolidated_df.empty
            and "order_id" in consolidated_df.columns and "source" in consolidated_df.columns):
        cod_rows = consolidated_df[consolidated_df["source"].isin(cod_gateway_labels)]
        if "is_refund" in cod_rows.columns:
            cod_rows = cod_rows[~cod_rows["is_refund"]]
        if not cod_rows.empty:
            # An order could in principle have non-refund rows from more
            # than one COD gateway (a genuine courier hand-off) - keep the
            # alphabetically-first as a simple, deterministic tie-break;
            # any one real COD source is a strictly better answer than a
            # guess from a different report either way.
            settlement_label_by_order = (
                cod_rows.groupby(cod_rows["order_id"].astype(str))["source"]
                .agg(lambda s: sorted(set(s))[0]).to_dict()
            )
            from_settlement = order_id_text.map(settlement_label_by_order)

    # Fallback: the pre-existing delivery_partner-based guess, used only
    # where the factual settlement signal above didn't resolve anything.
    # 2026-09-06 (round 16, order #26105): matched by courier-family
    # PREFIX now, not an exact f"{delivery_partner} COD" match - see
    # _match_cod_courier_label()'s own docstring and this function's
    # docstring above.
    from_delivery = pd.Series(None, index=df.index, dtype=object)
    if "delivery_partner" in df.columns:
        from_delivery = df["delivery_partner"].astype(str).str.strip().apply(
            lambda text: _match_cod_courier_label(text, cod_gateway_labels)
        )

    candidate = from_settlement.where(from_settlement.notna(), from_delivery)
    candidate = candidate.where(candidate.isin(cod_gateway_labels))
    return candidate.where(is_cod)


def _match_cod_courier_label(delivery_partner_text, cod_gateway_labels):
    """
    2026-09-06 (round 16, order #26105) - matches a delivery-status
    tracking report's own courier name against a configured COD gateway
    label by COURIER-FAMILY PREFIX, not by exact f"{delivery_partner} COD"
    concatenation - see resolve_cod_courier_label()'s own docstring for
    the order #26105 story (delivery_partner "Delhivery D2C" needed to
    resolve to the configured COD gateway label "Delhivery COD").

    Tries an exact match first (delivery_partner + " COD" == some
    configured label) - the original, still-correct behaviour for the
    overwhelming majority of couriers whose delivery-status report name
    and COD-gateway config label already agree exactly. Only when that
    fails does it fall back to a case-insensitive PREFIX match: a
    configured COD label with its own trailing " COD" stripped (its
    "courier family name", e.g. "Delhivery" from "Delhivery COD") checked
    against the START of delivery_partner_text. When more than one
    configured COD label's prefix matches (unlikely, but handled rather
    than guessed at), the LONGEST/most-specific prefix wins.

    Returns the matched configured COD gateway label, or None when
    nothing - exact or prefix - matches (e.g. "Self Delivery / Manual
    Report", "partner undifined") - callers already treat None as "not
    resolvable", never a guess.
    """
    text = str(delivery_partner_text or "").strip()
    if not text:
        return None
    exact = f"{text} COD"
    if exact in cod_gateway_labels:
        return exact
    text_upper = text.upper()
    best_label, best_prefix_len = None, -1
    for label in cod_gateway_labels:
        if not label.upper().endswith(" COD"):
            continue
        prefix = label[: -len(" COD")].strip()
        if prefix and text_upper.startswith(prefix.upper()) and len(prefix) > best_prefix_len:
            best_label, best_prefix_len = label, len(prefix)
    return best_label


def has_genuine_prepaid_receipt(order_ids, consolidated_df, gateway_configs):
    """
    Client-reported 2026-09-05 (order #27533 and siblings - the SAME six
    orders round 9 already fixed once, see resolve_cod_courier_label()'s
    own docstring): "Payment Provider is Shiprocket COD but tool report
    showing 'easebuzz, Shiprocket COD'". Root cause: combine_prepaid_and_
    cod_label() (round 10) treats ANY non-blank, non-"Gokwik", non-"<
    courier> COD"-shaped existing label as a genuine prepaid leg worth
    keeping alongside the COD override - but the Gokwik Order Report /
    Transaction Report attribution this module reads is, per its OWN
    module docstring, "purely informational - never changes any
    settlement $ math". An order can show up in those two reports (e.g.
    an abandoned/failed Gokwik checkout attempt before the customer paid
    COD instead - the exact root cause round 9 already diagnosed for
    these same six orders) with NO real money ever having moved through
    that processor at all. Round 10's combine logic couldn't tell that
    apart from order #31956's genuine part-prepaid/part-COD split (a REAL
    ₹95.67 collected via easebuzz), so it silently regressed round 9's
    fix for every order whose only "prepaid" evidence is a phantom
    attribution-report entry.

    This function is the disambiguator: was any money for this order_id
    EVER actually reported through a genuinely prepaid-configured gateway
    (gateway_configs entries whose own payment_mode is "Prepaid" - e.g.
    Razorpay, or the raw "Gokwik" aggregator label consolidated_df itself
    uses BEFORE the attribution overlay refines it to "easebuzz"/"payu" -
    see build_payment_gateway_lookups()'s own docstring for that
    raw-vs-refined distinction)? consolidated_df is the actual settlement
    ledger - a row there means a real transaction was reported by that
    gateway's own raw file, independent of whatever the (separate,
    optional, informational-only) Gokwik Order/Transaction Report
    attribution overlay guesses. Order #31956 has a real Gokwik-sourced
    row here; order #27533 and its round-9 siblings do not (confirmed:
    round 9's own docstring already states these six have "nothing ever
    collected through any payment gateway - receipt_amount 0").

    Returns a plain Python set of order_id strings that DO have such a
    row - empty set (never an error) when consolidated_df or
    gateway_configs is missing/empty, so a caller that can't supply this
    (or an isolated unit test) degrades to "nothing verified as genuine",
    matching combine_prepaid_and_cod_label()'s own conservative default
    when its has_real_prepaid_receipt gate isn't supplied at all.
    """
    if consolidated_df is None or consolidated_df.empty or not gateway_configs:
        return set()
    if "source" not in consolidated_df.columns or "order_id" not in consolidated_df.columns:
        return set()
    prepaid_labels = {
        str(cfg.get("label", "")).strip()
        for cfg in gateway_configs
        if str(cfg.get("payment_mode", "")).strip().lower() == "prepaid"
    }
    if not prepaid_labels:
        return set()
    hits = consolidated_df[consolidated_df["source"].isin(prepaid_labels)]
    return set(hits["order_id"].astype(str))


def has_confirmed_prepaid_transaction(order_ids, attribution_frames, attribution_sources_cfg):
    """
    2026-09-06 (round 16, order #30456): a SECOND, independent source of
    "genuine, not phantom" prepaid evidence for combine_prepaid_and_cod_
    label()'s has_real_prepaid_receipt gate - see has_genuine_prepaid_
    receipt() above for the first (a raw settlement-ledger row from a
    Prepaid-configured gateway). Both feed resolve_prepaid_evidence()
    below, which every real caller now uses instead of calling either one
    alone.

    Client-reported: order #30456 - a genuine part-prepaid/part-COD order
    (customer paid PART of the order via a Gokwik-routed PayU transaction
    at checkout, the balance COD on delivery through Delhivery - the
    client's own words: "customer choosed prepaid and paid part amount it
    apearing in transaction report during delivery time customer paid
    balance amount that is the reason its coming both the report" -
    exactly like order #31956's round-10 scenario) kept showing PLAIN
    "Delhivery COD" instead of the combined "Delhivery COD, PayU" the
    client's own workbook expects. Root cause: has_genuine_prepaid_
    receipt() alone requires an actual row in consolidated_df (the RAW
    settlement ledger, built from each gateway's own settlement/remittance
    file) sourced from a Prepaid-configured gateway - correct evidence
    when it exists, but a real checkout-time payment can show up in the
    (separate, purely informational per this module's own docstring)
    Gokwik Order Report / Transaction Report attribution well before that
    gateway's OWN settlement file has actually posted the money for this
    specific order/period - a timing gap, not a phantom transaction.

    This function reads the SAME two attribution reports resolve_cod_
    courier_label()'s sibling maps (build_gokwik_payment_provider_map(),
    build_direct_transaction_provider_map()) already parse, but asks a
    narrower question those maps don't expose: for this order_id, is
    there a transaction row whose own STATUS is genuinely successful (not
    merely "some row for this order_id exists in the report at all")?
    That status signal is exactly what already distinguishes a genuine
    payment from the phantom/abandoned-checkout hits round 9's own fix
    targeted (order #27533 and siblings - "most plausibly an abandoned/
    failed Gokwik checkout attempt" - see resolve_cod_courier_label()'s
    docstring): an abandoned checkout's own Transaction Report row would
    show a failed/pending status, not "success" - so this function
    correctly excludes those orders too, without needing the raw
    settlement ledger to have caught up at all.

    Mirrors both join paths attach_payment_columns() already relies on:
      1. The direct, single-file order-number-column join (see
         build_direct_transaction_provider_map()'s own docstring) -
         wherever a transaction_report source config carries an
         order_number_col.
      2. The two-report Order Report <-> Transaction Report Payment-ID
         join (see build_gokwik_payment_provider_map()'s own docstring) -
         for a provider where both halves are present.
    A row counts as "confirmed successful" if its own status_col/
    success_values match (same status_col "Status" / success_values
    ["success"] defaults build_direct_transaction_provider_map() already
    uses); a transaction_report source with NO recognisable status column
    at all degrades to "presence counts as confirmed" (same graceful-
    degrade philosophy this module already applies elsewhere - a raw
    Transaction Report export with no status column at all typically only
    ever lists completed transactions in the first place, so treating
    every row as confirmed there is the safer default, not a guess).

    Returns a plain Python set of order_id strings - empty (never an
    error, never None) when attribution_frames/attribution_sources_cfg is
    missing or empty, or no transaction_report source is configured at
    all - callers union this with has_genuine_prepaid_receipt()'s own
    result (see resolve_prepaid_evidence()), so an empty set here simply
    contributes nothing rather than blocking anything.
    """
    if not attribution_frames or not attribution_sources_cfg:
        return set()

    def _success_mask(txn_df, cfg):
        status_col = resolve_col(txn_df, cfg.get("status_col") or "Status")
        if not status_col:
            return pd.Series(True, index=txn_df.index)
        success_values = {str(v).strip().lower() for v in cfg.get("success_values", ["success"])}
        status_text = txn_df[status_col].astype(str).str.strip().str.lower()
        return status_text.isin(success_values)

    confirmed = set()

    # --- Path 1: direct order-number-column join (see build_direct_
    # transaction_provider_map's own docstring) --------------------------
    for cfg in attribution_sources_cfg:
        if cfg.get("role") != "transaction_report" or not cfg.get("order_number_col"):
            continue
        txn_df = attribution_frames.get(cfg.get("label"))
        if txn_df is None or txn_df.empty:
            continue
        order_number_col = resolve_col(txn_df, cfg.get("order_number_col"))
        if not order_number_col:
            continue
        tmp_order_id = normalize_order_id(txn_df[order_number_col])
        success = _success_mask(txn_df, cfg)
        hit_mask = (tmp_order_id.str.len() > 0) & success
        confirmed.update(tmp_order_id[hit_mask])

    # --- Path 2: two-report Order Report <-> Transaction Report Payment-ID
    # join (see build_gokwik_payment_provider_map's own docstring) -------
    by_provider = {}
    for cfg in attribution_sources_cfg:
        label = cfg.get("label")
        df = attribution_frames.get(label)
        if df is None or df.empty:
            continue
        by_provider.setdefault(cfg.get("provider"), {})[cfg.get("role")] = (df, cfg)

    for provider, roles in by_provider.items():
        if "order_report" not in roles or "transaction_report" not in roles:
            continue
        order_df, order_cfg = roles["order_report"]
        txn_df, txn_cfg = roles["transaction_report"]
        order_id_col = resolve_col(order_df, order_cfg.get("order_id_col"))
        order_payment_id_col = resolve_col(order_df, order_cfg.get("payment_id_col"))
        txn_payment_id_col = resolve_col(txn_df, txn_cfg.get("payment_id_col"))
        if not order_id_col or not order_payment_id_col or not txn_payment_id_col:
            continue

        order_side = pd.DataFrame({
            "order_id": normalize_order_id(order_df[order_id_col]),
            "_payment_id": _normalize_join_id(order_df[order_payment_id_col]),
        })
        success = _success_mask(txn_df, txn_cfg)
        txn_side = pd.DataFrame({
            "_payment_id": _normalize_join_id(txn_df[txn_payment_id_col]),
            "_success": success.values,
        })
        _degenerate_ids = {"", "NAN", "NONE", "0", "-", "NA", "N/A"}
        order_side = order_side[~order_side["_payment_id"].str.upper().isin(_degenerate_ids)]
        txn_side = txn_side[~txn_side["_payment_id"].str.upper().isin(_degenerate_ids)]
        # Any successful row for a Payment ID is enough - no need for
        # _collapse_unambiguous's conflict-avoidance here, since this
        # function only ever answers a yes/no "was real money confirmed"
        # question, never names a specific provider (that's still
        # gokwik_provider_map/direct_txn_provider_map's own job).
        success_payment_ids = set(txn_side.loc[txn_side["_success"], "_payment_id"])
        joined = order_side[order_side["_payment_id"].isin(success_payment_ids)]
        joined = joined[joined["order_id"].str.len() > 0]
        confirmed.update(joined["order_id"])

    return confirmed


def resolve_prepaid_evidence(order_ids, consolidated_df=None, gateway_configs=None,
                              attribution_frames=None, attribution_sources_cfg=None):
    """
    2026-09-06 (round 16) - the single entry point every real caller now
    uses instead of has_genuine_prepaid_receipt() alone, unioning it with
    has_confirmed_prepaid_transaction() (see that function's own
    docstring for the order #30456 story on why one source alone missed a
    genuine part-prepaid/part-COD order).

    Returns None (not an empty set) only when NEITHER underlying source
    was actually supplied at all (consolidated_df/gateway_configs both
    absent AND attribution_frames/attribution_sources_cfg both absent) -
    combine_prepaid_and_cod_label() treats a bare None as "can't verify
    anything, fall back to the pre-2026-09-05 permissive behaviour" (see
    its own docstring) - the same graceful-degrade this module's isolated
    unit tests already rely on. The moment EITHER source is actually
    supplied, this returns a real (possibly empty) set, so a genuinely
    unverified order is correctly excluded rather than assumed genuine.
    """
    have_ledger_source = consolidated_df is not None and gateway_configs
    have_attribution_source = attribution_frames and attribution_sources_cfg
    if not have_ledger_source and not have_attribution_source:
        return None
    evidence = set()
    if have_ledger_source:
        evidence |= has_genuine_prepaid_receipt(order_ids, consolidated_df, gateway_configs)
    if have_attribution_source:
        evidence |= has_confirmed_prepaid_transaction(order_ids, attribution_frames, attribution_sources_cfg)
    return evidence


def combine_prepaid_and_cod_label(existing_series, cod_override, order_id_series=None, has_real_prepaid_receipt=None):
    """
    Shared "<courier> COD, prepaid_provider" combiner (2026-09-04, round
    10; label order reversed 2026-09-06, round 16) - used identically for
    the "Payment Provider" column (attach_payment_columns above) and the
    "Gateway" column (both view call sites, views/page_reconciliation.py
    and views/page_reports.py), so query refinement (which keys off
    "Gateway" - see engine.reco.refine_queries_with_settlement_status) and
    the Payment Provider display never disagree about a mixed order.

    Client-reported order #31956 (round 10): a "Gokwik PPCOD" order (a
    real part-prepaid/part-COD checkout - ₹95.67 paid via easebuzz at
    checkout, the ₹890 balance collected COD on delivery through
    Delhivery) should show BOTH legs combined. Before that fix, resolve_
    cod_courier_label()'s COD override (round 8/9) simply overwrote
    whatever prepaid value was already there - correct for a pure-COD
    order (nothing to combine with), but silently dropped the genuine
    prepaid leg of a mixed order.

    Client-reported again 2026-09-05 (order #27533 and round 9's other
    five siblings - a REGRESSION of round 9's own fix): combining is only
    correct when the existing "prepaid-looking" label is backed by REAL
    money (see resolve_prepaid_evidence() above) - not merely present
    because the Gokwik attribution overlay (purely informational) happens
    to name a processor for this order_id. Two optional params, both
    required together: order_id_series (the same order_id column the
    caller's existing_series/cod_override are aligned against - same
    index, e.g. reco_df["order_id"]) and has_real_prepaid_receipt (a
    set/collection of order_id strings with genuine prepaid evidence - see
    resolve_prepaid_evidence() above). When both are supplied, a row only
    combines if BOTH the label looks prepaid-like AND its order_id is in
    that set; otherwise it falls through to the plain "<courier> COD"
    replace path (round 9's original, correct behaviour for a phantom
    attribution hit). When left as None (default - an isolated caller with
    nothing to check against, e.g. a unit test exercising the combine
    mechanics directly), behaves exactly as before that fix: any
    non-blank, non-Gokwik, non-COD-shaped label is treated as genuine - so
    this stays purely additive, never a required argument.

    Client-reported 2026-09-06 (round 16, order #30456): the combined
    label's own ORDER is now "<courier> COD, prepaid_provider" - e.g.
    "Delhivery COD, PayU" - not "prepaid_provider, <courier> COD" as
    round 10 originally produced ("easebuzz, Delhivery COD" for order
    #31956). This is a display-order change for EVERY combined order, not
    just #30456 - flagged explicitly since it reverses round 10's own
    convention per the client's latest explicit instruction ("the Payment
    Provider should be 'Delhivery COD, PayU'").

    Only combines where existing_series already holds a genuine PREPAID
    processor name for that row - not blank, not the bare "Gokwik"
    aggregator placeholder (that's a "we don't know which processor"
    fallback, not a real answer worth keeping alongside the COD label),
    and not itself already a "<courier> COD"-shaped label (nothing new to
    add - covers this function being applied twice, or an order whose
    only resolved value so far is already a COD label). Every other COD
    order - the overwhelming majority, no prepaid leg at all - keeps the
    exact plain "<courier> COD" behaviour every existing caller already
    had, unchanged.
    """
    result = existing_series.copy()
    has_cod = cod_override.notna()
    if not has_cod.any():
        return result
    existing_text = existing_series.astype(str).str.strip()
    is_prepaid_like = (
        existing_series.notna()
        & (existing_text.str.len() > 0)
        & (existing_text.str.lower() != "gokwik")
        & (existing_text.str.lower() != "nan")
        & (~existing_text.str.upper().str.endswith(" COD"))
    )
    if has_real_prepaid_receipt is not None and order_id_series is not None:
        # order_id_series must share existing_series's own index (both are
        # always columns off the SAME reco_df at every real call site) so
        # this aligns by position, not by any order_id value coincidence.
        verified = order_id_series.astype(str).isin(has_real_prepaid_receipt)
        is_prepaid_like = is_prepaid_like & verified
    cod_text = cod_override.astype(str).str.strip()
    same_value = is_prepaid_like & has_cod & (existing_text == cod_text)
    combine = has_cod & is_prepaid_like & ~same_value
    # 2026-09-06 (round 16): "<courier> COD" now comes FIRST, the prepaid
    # provider SECOND - see docstring above (order #30456).
    result.loc[combine] = cod_text[combine] + ", " + existing_text[combine]
    plain_cod = has_cod & ~combine
    result.loc[plain_cod] = cod_override[plain_cod]
    return result


def cod_component_of_gateway_label(label):
    """
    2026-09-06 (round 16, order #30456) - shared helper for every OTHER
    module that reads the "Gateway"/"Payment Provider" column and expects
    a single, bare gateway/courier label (engine.reco.refine_queries_with_
    settlement_status(), engine.settlement_pending.settlement_pending_
    summary_by_gateway()): combine_prepaid_and_cod_label() above can now
    hand that column a COMBINED "<courier> COD, <prepaid provider>" label
    for a genuine part-prepaid/part-COD order (e.g. "Delhivery COD,
    PayU") - and always puts the COD-configured label FIRST (round 16's
    own ordering convention).

    For every purpose those other modules care about - which courier is
    this order's money still waiting on, which Payment Gateway row should
    it group under - the answer is always the COD leg: the prepaid leg
    was already collected at checkout, so it is the COD leg that is
    genuinely outstanding/pending. Splitting on ", " and keeping only the
    first component recovers that COD label; a bare (non-combined) label
    is returned unchanged, so every caller can apply this unconditionally
    without needing to first check whether the value is combined.
    """
    # 2026-09-06 (round 21, client-reported - phantom "nan" row in Gateway
    # Settlement overall): `str(label or "")` looks like a safe blank-out
    # for a missing label, but a pandas NaN FLOAT (as opposed to Python
    # None - what an unattributed order's "Payment Provider" actually
    # holds after a merge) is truthy, so `nan or ""` evaluates to the NaN
    # itself, and `str(nan)` is the literal text "nan" - not blank. Every
    # caller of this function then groups/displays that as a real-looking
    # gateway named "nan" (engine.settlement.gateway_settlement_overall()'s
    # own deduction_by_gateway groupby, in particular). Same latent bug,
    # same fix, as engine/settlement.py::expected_gateway_for_order()'s own
    # 2026-09-06 (round 18) fix - use pd.isna() instead of Python
    # truthiness.
    text = "" if pd.isna(label) else str(label).strip()
    return text.split(", ", 1)[0].strip() if ", " in text else text


def prepaid_component_of_gateway_label(label):
    """
    Companion to cod_component_of_gateway_label() just above, added
    2026-09-06 (round 17, order #30456's own direct follow-up) for
    engine.reco.refine_split_payment_queries() - unlike every other
    caller (which only ever wants the COD leg), that function needs BOTH
    legs of a combined "<courier> COD, <prepaid provider>" label, since
    its whole job is to spell out each leg's own settlement status
    ("Delhivery COD Setled, payu Setlment Pending").

    Returns None for a bare (non-combined) label - callers use that to
    detect "this isn't actually a split-payment order" without a separate
    ", " check of their own, the same way a bare label passing through
    cod_component_of_gateway_label() unchanged signals "nothing to split".
    """
    # See cod_component_of_gateway_label()'s own note just above - same
    # pd.isna() fix, same NaN-truthiness bug.
    text = "" if pd.isna(label) else str(label).strip()
    if ", " not in text:
        return None
    return text.split(", ", 1)[1].strip()


def build_direct_transaction_provider_map(attribution_frames, attribution_sources_cfg):
    """
    2026-09-04 (round 10) - a more direct, single-file alternative to
    build_gokwik_payment_provider_map()'s own two-report (Order Report +
    Transaction Report, joined via Payment ID) approach, used to REFINE
    (not replace - see attach_payment_columns() above) its result.

    Client-reported: order #28980 showed "easebuzz" (the two-report
    join's answer) when the freshly re-uploaded Gokwik Transaction Report
    itself plainly lists Payment Provider "payu" against Payment ID
    KWIKA5YCNLHE7375404MP - which that same report row already tags with
    Platform order number "#28980". Many Gokwik Transaction Report
    exports carry this Shopify order-number column directly, making the
    fragile two-file Payment-ID join (already the root cause of several
    earlier rounds' "wrong provider" reports - stale/duplicate Payment
    IDs on either side, see _collapse_unambiguous's own docstring)
    entirely unnecessary wherever it's present.

    Optional (the transaction_report config entry's own "order_number_col"
    key) - a Transaction Report export without this column, or with it
    configured but not actually present, simply yields an empty map here,
    and attach_payment_columns() falls back to gokwik_provider_map alone,
    exactly as before this fix existed.

    Returns order_id | payment_provider - one row per Shopify order the
    Transaction Report's own order-number column resolves to a single,
    unambiguous Payment Provider for. When the report also carries a
    recognisable status column (the "status_col"/"success_values" config
    keys - optional, defaults to a plain "Status" column and "success"),
    a SUCCESSFUL transaction row is preferred over an abandoned/failed one
    for the same order wherever both exist, so a failed first attempt
    can't shadow the real, completed payment. An order whose transaction
    rows genuinely disagree on provider even after that preference (e.g.
    two separate successful attempts through different processors) is
    dropped from the map rather than guessed at - the exact same
    "don't guess a conflict" rule _collapse_unambiguous already applies
    elsewhere in this module.
    """
    cols = ["order_id", "payment_provider"]
    if not attribution_frames or not attribution_sources_cfg:
        return pd.DataFrame(columns=cols)

    frames = []
    for cfg in attribution_sources_cfg:
        if cfg.get("role") != "transaction_report" or not cfg.get("order_number_col"):
            continue
        txn_df = attribution_frames.get(cfg.get("label"))
        if txn_df is None or txn_df.empty:
            continue
        order_number_col = resolve_col(txn_df, cfg.get("order_number_col"))
        provider_col = resolve_col(txn_df, cfg.get("payment_provider_col"))
        if not order_number_col or not provider_col:
            continue

        tmp = pd.DataFrame({
            "order_id": normalize_order_id(txn_df[order_number_col]),
            "payment_provider": txn_df[provider_col].astype(str).str.strip(),
        })
        tmp = tmp[(tmp["order_id"].str.len() > 0) & (tmp["payment_provider"].str.len() > 0)]

        status_col = resolve_col(txn_df, cfg.get("status_col") or "Status")
        if status_col and len(tmp):
            success_values = {str(v).strip().lower() for v in cfg.get("success_values", ["success"])}
            status_text = txn_df.loc[tmp.index, status_col].astype(str).str.strip().str.lower()
            tmp = tmp.assign(_is_success=status_text.isin(success_values))
            has_success = tmp.groupby("order_id")["_is_success"].transform("any")
            # Keep every successful row, plus any order with NO successful
            # row at all (keep its non-successful rows rather than drop the
            # order entirely - still better than no signal, and a genuine
            # multi-value conflict is caught by _collapse_unambiguous below
            # exactly as it would be otherwise).
            tmp = tmp[tmp["_is_success"] | ~has_success].drop(columns="_is_success")

        if len(tmp):
            frames.append(tmp[cols])

    if not frames:
        return pd.DataFrame(columns=cols)
    combined = pd.concat(frames, ignore_index=True)
    out, _conflicts = _collapse_unambiguous(combined, "order_id", "payment_provider")
    return out.reset_index(drop=True)


def _normalize_join_id(series):
    """
    Same defensive cleanup engine.loaders.normalize_order_id already
    applies to every order_id join key in this app, extended here to the
    Payment ID columns this module joins on - which never had it. Excel/
    pandas will silently read a numeric-looking ID column (Payment ID,
    UTR, order number...) as a float64 rather than text whenever the
    source file didn't format that column as Text, appending a trailing
    ".0" that a same-value text column on the OTHER side of a join won't
    have - without this strip, that alone is enough to make two sides
    that actually agree fail to match. Applied even-handedly to both the
    Gokwik Order Report and Gokwik Transaction Report's Payment ID
    columns (2026-08-31 fix, see build_gokwik_payment_provider_map).
    """
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def _collapse_unambiguous(df, key_col, value_col):
    """
    Collapses df to one row per key_col - but ONLY for a key where every
    row agrees on value_col. A key with more than one distinct value_col
    (e.g. one Payment ID the Transaction Report lists against two
    different downstream providers, or one Shopify order_id whose Order
    Report rows resolve to two different Payment IDs that in turn map to
    two different providers - most plausibly two separate payment
    attempts) is a genuine data conflict this join has no reliable way to
    arbitrate: neither Gokwik report carries an amount, a timestamp, or a
    success/failure flag to say which attempt actually counts.

    The function this replaced (drop_duplicates(..., keep="last")) picked
    a side anyway - silently, based on nothing more than which row
    happened to land last in the uploaded file. That is the most likely
    root cause of the 2026-08-31 client report that orders 22671 and
    24497 show "Easebuzz" in the tool when the Gokwik Transaction Report
    actually says "PayU" for them: two of these joins almost certainly
    had a genuine collision (of either kind above, or a false one caused
    by the missing ".0" normalization this same fix adds - see
    _normalize_join_id), and file-row-order happened to pick the wrong
    side for exactly those two.

    This function instead DROPS a conflicting key from the map entirely
    rather than guessing which value is correct. Callers already treat
    "not in the map" as "no refinement available" and fall back to the
    generic "Gokwik" label (see attach_payment_columns / build_payment_
    gateway_lookups) - less specific, never wrong, the same fallback
    philosophy this module's own docstring already commits to for every
    other case where a specific answer can't be confirmed. This can only
    ever make the map SMALLER (safer) than before, never introduce a new
    wrong label - real June/July data was not available to re-run this
    against directly (only the two order IDs were reported, not the raw
    files), so this is a code-level fix verified with a synthetic
    regression case reproducing the same "two wrong, rest correct"
    symptom pattern, not a confirmed trace of orders 22671/24497
    specifically - flagged here rather than overstated.

    Returns (resolved_df, conflicts_df) - conflicts_df exists purely for
    diagnostics/logging by a caller that wants it; nothing currently
    consumes it.
    """
    if df.empty:
        return df.iloc[0:0].reset_index(drop=True), df.iloc[0:0].reset_index(drop=True)
    grouped = df.groupby(key_col)[value_col].agg(lambda s: sorted(set(s)))
    is_unique = grouped.apply(len) == 1
    resolved = grouped[is_unique].apply(lambda vals: vals[0]).rename(value_col).reset_index()
    conflicts = grouped[~is_unique].rename(f"conflicting_{value_col}").reset_index()
    return resolved, conflicts


def build_gokwik_payment_provider_map(attribution_frames, attribution_sources_cfg):
    """
    attribution_frames: {label: df} - whatever's been uploaded under the
    "Payment Gateway Attribution" upload section (see
    views/page_upload.py), keyed by each attribution_sources config
    entry's own "label".
    attribution_sources_cfg: config["attribution_sources"] - each entry
    tagged with a "provider" (e.g. "Gokwik") and a "role"
    ("order_report" | "transaction_report").

    Returns order_id | payment_provider - one row per Shopify order
    whose Gokwik Order Report entry resolves to a Payment ID the
    Transaction Report also has a Payment Provider for. Empty (not an
    error) if either report for a given provider hasn't been uploaded
    yet, or attribution_sources isn't configured at all - callers treat
    that as "no refinement available", never as a failure.

    2026-08-31 fix (client-reported: orders 22671/24497 showed "Easebuzz"
    when the Gokwik Transaction Report says "PayU" for them) - two
    changes, both described in full on the helpers above:
      1. Payment ID join keys are now normalized the same defensive way
         every order_id join key already is (_normalize_join_id), so a
         Payment ID that is genuinely identical on both sides can't fail
         to match just because one file stored it as a number and the
         other as text.
      2. A Payment ID or order_id that genuinely resolves to more than
         one distinct provider is no longer silently resolved by "last
         row wins" (_collapse_unambiguous) - it's dropped from the map,
         falling back to the generic "Gokwik" label rather than risking
         another wrong specific-provider label.
    """
    cols = ["order_id", "payment_provider"]
    if not attribution_frames or not attribution_sources_cfg:
        return pd.DataFrame(columns=cols)

    by_provider = {}
    for cfg in attribution_sources_cfg:
        label = cfg.get("label")
        df = attribution_frames.get(label)
        if df is None or df.empty:
            continue
        by_provider.setdefault(cfg.get("provider"), {})[cfg.get("role")] = (df, cfg)

    frames = []
    for provider, roles in by_provider.items():
        if "order_report" not in roles or "transaction_report" not in roles:
            continue  # need BOTH halves of the join to attribute anything
        order_df, order_cfg = roles["order_report"]
        txn_df, txn_cfg = roles["transaction_report"]

        order_id_col = resolve_col_or_raise(order_df, order_cfg["order_id_col"], order_cfg["label"])
        order_payment_id_col = resolve_col_or_raise(order_df, order_cfg["payment_id_col"], order_cfg["label"])
        txn_payment_id_col = resolve_col_or_raise(txn_df, txn_cfg["payment_id_col"], txn_cfg["label"])
        txn_provider_col = resolve_col_or_raise(txn_df, txn_cfg["payment_provider_col"], txn_cfg["label"])

        order_side = pd.DataFrame({
            "order_id": normalize_order_id(order_df[order_id_col]),
            "_payment_id": _normalize_join_id(order_df[order_payment_id_col]),
        })
        txn_side = pd.DataFrame({
            "_payment_id": _normalize_join_id(txn_df[txn_payment_id_col]),
            "payment_provider": txn_df[txn_provider_col].astype(str).str.strip(),
        })
        txn_side = txn_side[txn_side["payment_provider"].str.len() > 0]
        # 2026-08-31 (round 6) fix, client-reported: order 28980 (Payment
        # Method "Gokwik Cards", but receipt_amount == 0 - no settlement
        # row from ANY gateway file yet, i.e. this checkout never actually
        # completed/posted) showed a confident but wrong "easebuzz" label,
        # with no conflict for _collapse_unambiguous above to catch (only
        # ONE provider was on offer for whatever Payment ID this order
        # joined on - so nothing looked ambiguous). An order whose
        # checkout never completed has no real Payment ID to report yet;
        # if its Order Report row's Payment ID cell is blank/placeholder
        # (blank, "nan", "none", "0", "-", ...) - which normalize_order_id-
        # style ".0"/whitespace cleanup can't distinguish from a genuine
        # short numeric ID - joining on that empty/placeholder string would
        # silently attribute it to whatever OTHER unrelated transaction
        # happens to share that same blank/placeholder value in the
        # Transaction Report, which is a real transaction but not this
        # order's. A blank or placeholder Payment ID can never correctly
        # identify one specific transaction on either side of this join,
        # so both sides drop those rows before the merge, rather than
        # letting them collide on an accidental empty string.
        _degenerate_ids = {"", "NAN", "NONE", "0", "-", "NA", "N/A"}
        order_side = order_side[~order_side["_payment_id"].str.upper().isin(_degenerate_ids)]
        txn_side = txn_side[~txn_side["_payment_id"].str.upper().isin(_degenerate_ids)]
        # A Payment ID can legitimately repeat across a few transaction
        # rows (e.g. a webhook retry) with every row agreeing on the
        # provider - that collapses to one row for free. Only a REAL
        # conflict (two different providers for the same Payment ID) is
        # dropped rather than guessed at - see _collapse_unambiguous.
        txn_side, _txn_conflicts = _collapse_unambiguous(txn_side, "_payment_id", "payment_provider")

        joined = order_side.merge(txn_side, on="_payment_id", how="inner")
        joined = joined[joined["order_id"].str.len() > 0]
        if len(joined):
            frames.append(joined[["order_id", "payment_provider"]])

    if not frames:
        return pd.DataFrame(columns=cols)
    combined = pd.concat(frames, ignore_index=True)
    # Same principle at the order level: an order_id that (via more than
    # one Payment ID - e.g. a retried checkout) ends up pointing at more
    # than one distinct provider is a genuine conflict, not something to
    # resolve by file-row-order - drop it rather than guess.
    out, _order_conflicts = _collapse_unambiguous(combined, "order_id", "payment_provider")
    return out.reset_index(drop=True)


def build_payment_gateway_lookups(consolidated_df, gokwik_provider_map=None):
    """
    Builds the two lookups the "Payment Gateway" column (Bank Reco
    (UTR-wise)) and the Reco working "Gateway" column both need:
      - order_id -> payment gateway label
      - utr      -> payment gateway label (works even for a settlement
        ledger row whose order_id could never be resolved - see
        engine.bank.NOT_FOUND - since it's keyed off the UTR itself, not
        the order)

    Base attribution is simply "which configured gateway/COD source's
    OWN raw file this money came from" (consolidated_df["source"]) -
    correct and unambiguous for Razorpay / Delhivery COD / Shiprocket
    COD / Prozo COD, all of which report their own settlements
    directly. For a row sourced from "Gokwik" specifically, gokwik_
    provider_map (see build_gokwik_payment_provider_map above) - when
    supplied and when this order_id resolves in it - refines the label
    to the actual downstream processor (e.g. "easebuzz"/"payu") instead
    of the bare "Gokwik". A Gokwik-sourced row with no resolvable
    order_id (or no provider map supplied at all) keeps the plain
    "Gokwik" label - a fair, honest fallback rather than a guess.

    Two fixes, both client-reported 2026-08-30 (item 7 and the
    Gateway/"payu" gaps feeding items 12/15):

    Fix 1 - order_id_lookup used to come ONLY from consolidated_df, so an
    order gokwik_provider_map could resolve but that has no settlement
    row yet at all (nothing collected/reported so far) never got a
    "Gateway" value, even though the provider was already known from the
    Gokwik Order+Transaction reports. Those orders are now unioned in
    from gokwik_provider_map directly, so "Gateway" is populated as soon
    as the provider is known - independent of whether money has moved
    yet.

    Fix 2 - one UTR can carry several consolidated_df ROWS (e.g. a Gokwik
    aggregator UTR settling multiple orders, some refined to "easebuzz",
    others still bare "Gokwik" because THEIR OWN order didn't resolve in
    gokwik_provider_map) - utr_lookup used to comma-join every distinct
    label into one string ("easebuzz, Gokwik"), but the client's own
    workbook always shows exactly ONE gateway name per UTR (client-
    reported: "one UTR should map to exactly ONE Payment Gateway... one
    ORDER can legitimately have multiple gateways/UTRs [part-prepaid +
    part-COD], but the SAME UTR should not show multiple gateway names").
    Resolved by amount-weighted majority vote within each UTR: sum each
    candidate label's total absolute settled amount, prefer a specific
    label (anything but the generic "Gokwik" fallback) when one exists,
    and pick whichever label carries the most money - a real tie (rare,
    two genuinely different specific gateways sharing one UTR) still
    resolves deterministically to whichever has the larger amount rather
    than silently falling back to a comma-joined string again.

    Returns (order_id_lookup, utr_lookup), both pandas Series indexed by
    order_id / normalized UTR respectively - empty Series (never None)
    when consolidated_df has nothing to build from (Fix 1's union still
    applies even then, so a Gokwik-only run with no consolidated_df yet
    can still resolve "Gateway" from gokwik_provider_map alone).
    """
    provider_by_order = {}
    if gokwik_provider_map is not None and not gokwik_provider_map.empty:
        provider_by_order = dict(zip(
            gokwik_provider_map["order_id"].astype(str), gokwik_provider_map["payment_provider"],
        ))

    if consolidated_df is None or consolidated_df.empty:
        if provider_by_order:
            return pd.Series(provider_by_order), pd.Series(dtype=object)
        empty = pd.Series(dtype=object)
        return empty, empty

    df = consolidated_df.copy()
    df["order_id"] = df["order_id"].astype(str)

    def _label_for(row):
        if row["source"] == "Gokwik":
            refined = provider_by_order.get(row["order_id"])
            if refined:
                return refined
        return row["source"]

    df["_gateway_label"] = df.apply(_label_for, axis=1)

    order_id_lookup = df[df["order_id"].str.len() > 0].groupby("order_id")["_gateway_label"].agg(
        lambda s: ", ".join(sorted(set(s)))
    )
    # Fix 1: union in any gokwik_provider_map order with no consolidated_df
    # row of its own at all - never overwrites an order consolidated_df
    # already resolved.
    if provider_by_order:
        missing = {
            oid: provider for oid, provider in provider_by_order.items()
            if oid not in order_id_lookup.index
        }
        if missing:
            order_id_lookup = pd.concat([order_id_lookup, pd.Series(missing)])

    has_utr = df["utr"].notna() & (df["utr"].astype(str).str.strip().str.len() > 0)
    utr_df = df[has_utr].copy()
    utr_df["_utr_norm"] = utr_df["utr"].astype(str).str.strip().str.upper()
    utr_df["_abs_amount"] = utr_df["amount"].abs() if "amount" in utr_df.columns else 1.0

    def _one_label_per_utr(group):
        # Fix 2: amount-weighted majority vote, preferring a specific
        # label over the generic "Gokwik" fallback whenever a specific
        # one is present at all.
        weights = group.groupby("_gateway_label")["_abs_amount"].sum()
        specific = weights.drop(labels=["Gokwik"], errors="ignore")
        pool = specific if len(specific) else weights
        return pool.idxmax()

    utr_lookup = utr_df.groupby("_utr_norm").apply(_one_label_per_utr)

    return order_id_lookup, utr_lookup
