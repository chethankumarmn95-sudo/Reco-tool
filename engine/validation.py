"""
validation.py
--------------
Two checks, both driven entirely by the config file (no hardcoding):

1. validate_file_matches_source() - when a file is uploaded into, say, the
   "Delhivery" box, check it actually looks like a Delhivery export (has the
   columns Delhivery config expects) rather than silently accepting whatever
   was dropped in.

2. check_mandatory_sources() - before running, make sure every source marked
   "mandatory": true in the config actually has a file uploaded.
"""


def _signature_columns(source_cfg):
    """
    Build the list of columns a file MUST contain to plausibly be this
    source - derived from the config's own column mappings, so there's no
    separate list to keep in sync. Each entry can be a single column name
    or a list of acceptable aliases (any one of which satisfies the check).

    Deliberately only the columns a source can't function at all without -
    order_id_col everywhere (nothing can be joined without it), plus
    status_col for a delivery partner file (its entire reason for
    existing), amount_col for a gateway file (same), and payment_id_col
    for an attribution-only source (e.g. Gokwik Order/Transaction Report,
    2026-08-25 - neither carries order_id_col AND payment_id_col both, so
    payment_id_col has to be checked as its own signature key or an
    attribution file with no recognisable columns at all would pass this
    check for free). financial_status/
    subtotal/etc. used to be required here too, which meant a genuine
    Shopify order export missing just "Subtotal" (a column the rest of the
    engine can perfectly well run without - see engine/reco.py's
    build_order_master) got rejected outright as "not a valid orders
    file" before ever reaching the code that could have handled it. Every
    other column config defines is looked up by header name at the point
    it's actually used (see engine/reco.py, engine/consolidator.py) and
    degrades to blank/zero there if genuinely absent - this check exists
    only to catch "wrong report dropped in this box" (validate_file_matches_source
    below), not to gate on completeness.
    """
    specs = []
    for key in ("order_id_col", "status_col", "amount_col", "payment_id_col", "payment_provider_col"):
        if source_cfg.get(key):
            specs.append(source_cfg[key])
    return specs


def validate_file_matches_source(df, source_cfg):
    """
    Returns (is_valid, missing_columns).
    is_valid is False if ANY signature field has no matching column in the
    uploaded file (checking all its aliases) - meaning it's very likely the
    wrong report was dropped in this box.
    """
    from .loaders import resolve_col

    specs = _signature_columns(source_cfg)
    missing = []
    for spec in specs:
        if resolve_col(df, spec) is None:
            # Report the primary (first-listed) name in the error for clarity
            missing.append(spec[0] if isinstance(spec, list) else spec)
    return (len(missing) == 0, missing)


def friendly_mismatch_message(source_cfg):
    label = source_cfg["label"]
    return (
        f"Invalid file. Please upload the **{label}** report here — "
        f"the file you uploaded doesn't have the columns a {label} export should have."
    )


def check_mandatory_sources(config, orders_file, delivery_files, gateway_files):
    """
    Returns a list of human-readable labels for every mandatory source that
    is still missing. Empty list means all good to run.
    """
    missing = []

    if config["orders"].get("mandatory", True) and orders_file is None:
        missing.append(config["orders"]["label"])

    for d_cfg in config["delivery_partners"]:
        if d_cfg.get("mandatory", True) and d_cfg["label"] not in delivery_files:
            missing.append(d_cfg["label"])

    for g_cfg in config["gateways"]:
        if g_cfg.get("mandatory", True) and g_cfg["label"] not in gateway_files:
            missing.append(g_cfg["label"])

    return missing
