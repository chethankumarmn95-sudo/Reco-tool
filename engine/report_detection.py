"""
report_detection.py
--------------------
The tool's shared "what report is this, actually?" logic, used everywhere
a raw file gets uploaded (DTC orders/delivery-partners/gateways/bank
statement, Amazon MTR/Settlement Flat Files, Shiprocket COD's own two
sheets) - built once here so every upload slot in the app benefits from
the same identification logic instead of a one-off check per report.

Client's own framing (2026-08-21): the tool should be DATA-driven, not
NAME-driven - file name, sheet name, and sheet position are never used to
decide what a report is or whether it's the right one. Only the actual
column headers present matter (via engine.loaders.resolve_col's
case/whitespace-insensitive matching).

Two things live here:

1. A REGISTRY of every report type this app currently has a config for
   (built fresh from whatever configs/*.json are loaded - see
   build_report_registry() - never a hardcoded list to maintain by hand).
2. wrong_report_message(): given the file that failed its OWN report's
   column check, scans the registry for whatever OTHER report it actually
   looks like, and returns a specific "Wrong Report Detected: this looks
   like a <X> report" message when confident, or a specific "missing
   these columns" message when it doesn't confidently match anything -
   never a bare "Invalid File".
"""

from .loaders import signature_match_count, resolve_col, _normalize_col_name


def build_report_registry(config_labels):
    """
    Collects a {"label", "signature"} entry for every distinct raw report
    this app currently knows about, across EVERY loaded client/channel
    config - orders, delivery partners, gateways, bank statements, MTR
    reports, Settlement Flat Files, plus Shiprocket COD's own two source
    sheets. This is what lets "wrong report uploaded here" recognise ANY
    of the app's other report types, not just the ones that could
    plausibly land in this exact upload box - e.g. it catches a Delhivery
    file dropped into a Shiprocket box, or an Amazon MTR file dropped into
    a Shopify orders box, even though those live in entirely different
    configs.

    config_labels: the {label: config_dict} map already held in
    st.session_state["config_labels"] (see views/state_init.py) - every
    config the app has loaded, across every client/platform, not just the
    one currently active. Built fresh on every call (cheap - it's just
    walking small config dicts) so a config change is picked up
    immediately with no separate cache to invalidate.

    Each signature deliberately uses a handful of a source's more
    DISTINCTIVE columns (not just the bare minimum "does this file work
    at all" set validate_file_matches_source uses) - the two checks serve
    different purposes and are kept independent on purpose: the minimal
    set decides "can the engine actually run on this file", while this
    richer set decides "what does this file look like it IS", and a
    richer signature makes that second question far less likely to
    produce a false "looks like a Y report" guess.
    """
    registry = []
    seen_labels = set()

    def _add(label, signature):
        specs = [s for s in signature if s]
        if len(specs) < 2 or not label:
            return  # too weak a signature to ever be a confident match
        key = label.strip().lower()
        if key in seen_labels:
            return
        seen_labels.add(key)
        registry.append({"label": label, "signature": specs})

    for cfg in (config_labels or {}).values():
        if cfg.get("channel_type") == "marketplace":
            channel = cfg.get("channel_name") or "Marketplace"
            mtr_cols = cfg.get("mtr_columns", {}) or {}
            _add(f"{channel} MTR Report", [
                mtr_cols.get("order_id_col"), mtr_cols.get("sku_col"),
                mtr_cols.get("transaction_type_col"), mtr_cols.get("invoice_number_col"),
            ])
            settlement_cols = cfg.get("settlement_columns", {}) or {}
            _add(f"{channel} Settlement Flat File", [
                settlement_cols.get("settlement_id_col"), settlement_cols.get("transaction_type_col"),
                settlement_cols.get("deposit_date_col"), settlement_cols.get("total_amount_col"),
            ])
            bank_cfg = cfg.get("bank_statement")
            if bank_cfg:
                _add(bank_cfg.get("label") or f"{channel} Bank Statement", [
                    bank_cfg.get("amount_col"), bank_cfg.get("date_col"), bank_cfg.get("utr_col"),
                ])
        else:
            orders_cfg = cfg.get("orders", {}) or {}
            _add(orders_cfg.get("label") or "Orders", [
                orders_cfg.get("order_id_col"), orders_cfg.get("created_at_col"),
                orders_cfg.get("payment_method_col"), orders_cfg.get("month_col"),
            ])
            for d_cfg in cfg.get("delivery_partners", []) or []:
                _add(d_cfg.get("label"), [
                    d_cfg.get("order_id_col"), d_cfg.get("status_col"),
                    d_cfg.get("delivered_date_col"), d_cfg.get("rto_date_col"),
                ])
            for g_cfg in cfg.get("gateways", []) or []:
                _add(g_cfg.get("label"), [
                    g_cfg.get("order_id_col"), g_cfg.get("amount_col"),
                    g_cfg.get("type_col"), g_cfg.get("utr_col"), g_cfg.get("date_col"),
                ])
            for a_cfg in cfg.get("attribution_sources", []) or []:
                _add(a_cfg.get("label"), [
                    a_cfg.get("order_id_col"), a_cfg.get("payment_id_col"), a_cfg.get("payment_provider_col"),
                ])
            bank_cfg = cfg.get("bank_statement")
            if bank_cfg:
                _add(bank_cfg.get("label") or "Bank Statement", [
                    bank_cfg.get("amount_col"), bank_cfg.get("date_col"), bank_cfg.get("utr_col"),
                ])

    # Shiprocket COD's own two source sheets, registered directly from
    # engine.shiprocket_cod's own signatures - so e.g. a CRF-level sheet
    # dropped into a plain single-sheet gateway box elsewhere still gets
    # recognised for what it is, and vice versa.
    from .shiprocket_cod import AWB_SIGNATURE, CRF_SIGNATURE, AWB_SHEET_LABEL, CRF_SHEET_LABEL
    _add(f"Shiprocket {AWB_SHEET_LABEL}", AWB_SIGNATURE)
    _add(f"Shiprocket {CRF_SHEET_LABEL}", CRF_SIGNATURE)

    return registry


def _spec_aliases(spec):
    """A signature entry is either a single column name, or a list of
    acceptable aliases for the same field (per engine.loaders.resolve_col) -
    always returns the alias list form, so every caller here can treat
    every spec uniformly."""
    return spec if isinstance(spec, list) else [spec]


def build_column_doc_freq(registry):
    """
    normalized alias -> number of DISTINCT registry labels whose signature
    lists that alias (in any spec's alias list). This is what
    _weighted_signature_score() below uses to tell a DISTINCTIVE column
    (one that only ever shows up in a single report type's signature -
    e.g. Razorpay's own "transaction_entity") apart from a GENERIC one
    shared across many report types (e.g. "Order ID"/"Name"/"UTR" -
    repeated, under slightly different spellings, across nearly every
    gateway/delivery-partner config in this app, and "Amount"/
    "Settlement UTR" specifically shared between Gokwik and Razorpay).

    Client-reported (2026-08-21): a real Razorpay settlement export was
    being flagged as "looks like a Gokwik report" purely because 3 of
    Gokwik's 5 signature columns happen to be these generic, widely-shared
    names - a plain "how many columns matched" count has no way to tell
    that apart from a genuinely distinctive match. This is computed fresh
    from whatever's actually in the registry every time (same pattern as
    build_report_registry itself) - no per-column "is this generic"
    tagging to maintain by hand, and it adapts automatically as report
    types are added or removed from the app's configs.
    """
    freq = {}
    for entry in registry:
        seen_in_this_entry = set()
        for spec in entry["signature"]:
            for alias in _spec_aliases(spec):
                seen_in_this_entry.add(_normalize_col_name(alias))
        for key in seen_in_this_entry:
            freq[key] = freq.get(key, 0) + 1
    return freq


def _weighted_signature_score(df, signature, doc_freq):
    """
    Distinctiveness-weighted match score for one registry entry's
    signature against df - see build_column_doc_freq's docstring for why
    this exists. Each spec (field) contributes:
      - to `total`: 1 / doc_freq[its own most-distinctive alias] - the
        best-case weight this field COULD contribute, i.e. what it's
        worth if the LEAST-shared of its acceptable aliases is the one
        that actually appears.
      - to `score` (only when the spec actually resolves against df):
        1 / doc_freq[the alias that actually resolved] - always <= that
        field's own best-case weight, so score/total never exceeds 1.
    A field unique to one report type anywhere in the app contributes a
    full 1.0 when matched; a field shared across N report types
    contributes only 1/N - exactly inverting the old bug, where every
    matched column counted the same regardless of how many other reports
    also use that name.

    Returns (weighted_score, weighted_total, matched_min_doc_freq) - the
    third value is the LOWEST doc_freq among all columns that actually
    matched (i.e. how distinctive the single best matching column was),
    used by detect_wrong_report() below to require that at least one
    genuinely distinctive column - not just several generic ones adding up -
    was part of the match.
    """
    score, total = 0.0, 0.0
    matched_doc_freqs = []
    for spec in signature:
        aliases = _spec_aliases(spec)
        alias_freqs = [doc_freq.get(_normalize_col_name(a), 1) for a in aliases]
        total += 1.0 / min(alias_freqs)
        matched_col = resolve_col(df, spec)
        if matched_col is not None:
            matched_freq = doc_freq.get(_normalize_col_name(matched_col), 1)
            score += 1.0 / matched_freq
            matched_doc_freqs.append(matched_freq)
    return score, total, (min(matched_doc_freqs) if matched_doc_freqs else None)


# A candidate "other report" match needs ALL of these to be reported as
# confident - see detect_wrong_report()'s docstring for what each guards
# against.
MIN_RAW_COLUMN_MATCHES = 2
MIN_WEIGHTED_RATIO = 0.6
MIN_RATIO_MARGIN_OVER_OWN = 0.15
MAX_DISTINCTIVE_DOC_FREQ = 2


def identify_best_match(df, registry, exclude_labels=()):
    """
    Returns (label, score, total) for whichever registry entry df matches
    best BY RAW column count (used only for the human-readable "X of Y
    columns matched" text - see detect_wrong_report() below for the
    actual confidence decision, which uses distinctiveness weighting
    instead of this raw count). excludes any label in exclude_labels.
    (None, 0, 0) if df is empty or matches nothing in the registry at all.
    """
    if df is None or df.empty:
        return None, 0, 0
    exclude = {e.strip().lower() for e in exclude_labels if e}
    best_label, best_score, best_total = None, 0, 0
    for entry in registry:
        if entry["label"].strip().lower() in exclude:
            continue
        total = len(entry["signature"])
        if not total:
            continue
        score = signature_match_count(df, entry["signature"])
        if score > best_score:
            best_label, best_score, best_total = entry["label"], score, total
    return best_label, best_score, best_total


def detect_wrong_report(expected_label, df, registry, exclude_labels=None):
    """
    The one shared "does this file actually look like some OTHER known
    report, confidently enough to say so" decision - used both when a
    file fails its own report's minimal column check (wrong_report_message
    below) and when a raw_transform's own sheet-identification step fails
    outright (views/page_upload.py's _load_gateway_file).

    Client-reported (2026-08-21): a genuine Razorpay settlement export was
    being flagged "Wrong report detected... looks like a Gokwik report
    (3 of 5 expected columns matched)" - a false positive caused entirely
    by 3 generic column names (Order ID, Amount, Settlement UTR) that
    Gokwik's and Razorpay's raw exports both happen to use, under
    near-identical spellings. Real Razorpay/Gokwik files run through this
    function (see engine/report_detection.py's own test suite reasoning,
    verified directly against the client's two sample files):
        Real Razorpay file vs its OWN Razorpay signature:  ratio 0.83
        Real Razorpay file vs Gokwik's signature:          ratio 0.25
        Real Gokwik file vs Razorpay's signature:          ratio 0.25
        Real Gokwik file vs its OWN Gokwik signature:      ratio 1.00
    - i.e. the weighting on its own already keeps a real Razorpay file
    from ever reading as "mostly Gokwik", and a real Gokwik file is still
    caught cleanly if it's ever dropped in the Razorpay box.

    Two independent changes over the old plain-count approach:
      1. DISTINCTIVENESS WEIGHTING (build_column_doc_freq /
         _weighted_signature_score) - a column shared by many report
         types in this app's own registry counts for far less than one
         that's unique to a single report type. Computed fresh from the
         registry every time, so it covers every report type the app
         knows about, not just Razorpay/Gokwik, and adapts automatically
         as configs change.
      2. COMPARATIVE, not just absolute - the file's weighted ratio
         against the candidate OTHER report must beat its weighted ratio
         against its OWN expected report by a real margin (see
         MIN_RATIO_MARGIN_OVER_OWN), not just clear some fixed bar in
         isolation. A file that overwhelmingly matches its own type is
         never reported as "looks like X" merely for sharing a few
         generic columns with X.
    A confident match ALSO still requires (both carried over from the
    original design, kept as extra safety margins): at least
    MIN_RAW_COLUMN_MATCHES actual columns in common, and at least one of
    those matched columns being reasonably distinctive on its own
    (doc_freq <= MAX_DISTINCTIVE_DOC_FREQ) - so a "confident" match is
    never built ENTIRELY out of columns so generic they appear in most of
    the app's report types.

    Returns (other_label, raw_score, raw_total) when confident enough to
    report - raw_score/raw_total are plain column counts, for the
    human-readable message - else (None, 0, 0).
    """
    if df is None or df.empty:
        return None, 0, 0
    exclude_norm = {e.strip().lower() for e in (set(exclude_labels or ()) | {expected_label}) if e}
    doc_freq = build_column_doc_freq(registry)

    own_entry = next(
        (e for e in registry if e["label"].strip().lower() == (expected_label or "").strip().lower()), None
    )
    own_ratio = 0.0
    if own_entry:
        own_score, own_total, _ = _weighted_signature_score(df, own_entry["signature"], doc_freq)
        own_ratio = (own_score / own_total) if own_total else 0.0

    best_label, best_raw_score, best_raw_total, best_ratio, best_min_freq = None, 0, 0, 0.0, None
    for entry in registry:
        if entry["label"].strip().lower() in exclude_norm:
            continue
        raw_total = len(entry["signature"])
        if not raw_total:
            continue
        w_score, w_total, min_freq = _weighted_signature_score(df, entry["signature"], doc_freq)
        ratio = (w_score / w_total) if w_total else 0.0
        if ratio > best_ratio:
            raw_score = signature_match_count(df, entry["signature"])
            best_label, best_raw_score, best_raw_total, best_ratio, best_min_freq = (
                entry["label"], raw_score, raw_total, ratio, min_freq
            )

    if (
        best_label
        and best_raw_score >= MIN_RAW_COLUMN_MATCHES
        and best_ratio >= MIN_WEIGHTED_RATIO
        and best_min_freq is not None and best_min_freq <= MAX_DISTINCTIVE_DOC_FREQ
        and best_ratio >= own_ratio + MIN_RATIO_MARGIN_OVER_OWN
    ):
        return best_label, best_raw_score, best_raw_total
    return None, 0, 0


def wrong_report_message(expected_label, df, registry, missing_cols, exclude_labels=None):
    """
    Builds the right upload error message for a file that failed its own
    report's column check:

      - If the file confidently looks like some OTHER known report (see
        detect_wrong_report() above), say so BY NAME: "Wrong report
        detected... looks like a <X> report" - the client's explicit ask,
        so a Shiprocket file dropped into a Delhivery box doesn't just get
        a baffling "missing columns" message.
      - Otherwise, falls back to naming exactly which columns are
        missing - never a bare "Invalid File" (also the client's explicit
        ask - "errors should be generated only when the actual required
        data/columns are missing").

    exclude_labels defaults to just expected_label - pass additional
    labels to exclude when a single upload box's own signature happens to
    overlap with another registry entry (rare, but harmless to guard).
    """
    other_label, score, total = detect_wrong_report(expected_label, df, registry, exclude_labels=exclude_labels)
    if other_label:
        return (
            f"**Wrong report detected.** You uploaded this file for **{expected_label}**, but its "
            f"columns look like a **{other_label}** ({score} of {total} expected columns matched). "
            f"Please upload the correct **{expected_label}** report here instead."
        )
    missing_txt = ", ".join(str(m) for m in missing_cols) if missing_cols else "the expected columns"
    return (
        f"Invalid file for **{expected_label}**. This doesn't look like a {expected_label} report - "
        f"missing expected column(s): {missing_txt}."
    )
