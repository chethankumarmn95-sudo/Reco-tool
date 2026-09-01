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
                            recon_status_df=None, gateway_configs=None):
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

    New optional params, both required together for this fallback to run
    (silently skipped, not an error, if either is omitted - same
    degrade-gracefully philosophy as the rest of this module):
      recon_status_df: engine.bank.classify_order_bank_status()'s own
        output - specifically its "payment_type" column ("COD"/"Prepaid"/
        "Unknown"), the same already-validated COD/Prepaid signal
        refine_queries_with_settlement_status() and the Settlement Pending
        Summary already rely on (see classify_payment_type()'s own
        docstring for why a courier-COD settlement row always outranks a
        blank/ambiguous Payment Method field for this call's real data).
        Gates this fallback so it only ever fires for an order recon_
        status_df itself has already classified as COD - applying
        "<courier> COD" to an unresolved PREPAID order (Gokwik/Razorpay
        simply hasn't reported it yet) would be a confident wrong label,
        worse than leaving it blank.
      gateway_configs: config["gateways"] - used only to confirm
        f"{delivery_partner} COD" is actually a configured COD gateway
        label (payment_mode == "COD") before using it, so an unrecognised
        delivery_partner (e.g. "Self Delivery / Manual Report", which has
        no matching COD gateway entry) never produces a made-up label.

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

    fallback_method = df["payment_method"] if "payment_method" in df.columns else pd.Series(None, index=df.index)
    df["Payment Method"] = df["order_id"].map(method_by_order)
    df["Payment Method"] = df["Payment Method"].where(df["Payment Method"].notna(), fallback_method)

    # --- Payment Provider: Gokwik order_id -> downstream provider map -----
    provider_by_order = {}
    if gokwik_provider_map is not None and not gokwik_provider_map.empty:
        provider_by_order = dict(zip(
            gokwik_provider_map["order_id"].astype(str), gokwik_provider_map["payment_provider"],
        ))

    fallback_provider = df["Gateway"] if "Gateway" in df.columns else pd.Series(None, index=df.index)
    df["Payment Provider"] = df["order_id"].map(provider_by_order)
    df["Payment Provider"] = df["Payment Provider"].where(df["Payment Provider"].notna(), fallback_provider)

    # --- Payment Provider, third fallback: "<courier> COD" (2026-08-31, --
    # round 8 fix) - see this function's own docstring above for the full
    # story. Only fires where Payment Provider is STILL blank after both
    # fallbacks above, so it never overrides a real Gokwik/Gateway-sourced
    # answer.
    still_blank = df["Payment Provider"].isna()
    if still_blank.any() and "delivery_partner" in df.columns and gateway_configs:
        cod_gateway_labels = {
            str(cfg.get("label", "")).strip()
            for cfg in gateway_configs
            if str(cfg.get("payment_mode", "")).strip().lower() == "cod"
        }
        payment_type_by_order = {}
        if recon_status_df is not None and not recon_status_df.empty and "payment_type" in recon_status_df.columns:
            payment_type_by_order = dict(zip(
                recon_status_df["order_id"].astype(str), recon_status_df["payment_type"],
            ))
        if cod_gateway_labels and payment_type_by_order:
            is_cod = df["order_id"].map(payment_type_by_order).eq("COD")
            candidate = df["delivery_partner"].astype(str).str.strip() + " COD"
            candidate = candidate.where(candidate.isin(cod_gateway_labels))
            apply_mask = still_blank & is_cod & candidate.notna()
            df.loc[apply_mask, "Payment Provider"] = candidate[apply_mask]

    return df


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
