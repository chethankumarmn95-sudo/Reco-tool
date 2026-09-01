"""
dedup.py
--------
Shared "have we already ingested this row" logic for the Upload Data page,
covering the two different data-lifecycle patterns the client described:

1. SALES / ORDER-LEVEL reports (DTC Orders export, Amazon MTR reports) -
   these are permanent records once issued (an MTR invoice line doesn't
   change after the fact), so re-uploading a wider date range than before
   (e.g. first "01-Aug to 08-Aug", later "01-Aug to 31-Aug") should ADD
   whatever's genuinely new and SKIP whatever's already been reconciled -
   never double-count it. See split_new_vs_seen().

2. STATUS / SETTLEMENT reports (delivery partner files, payment gateway
   files, Amazon Settlement Flat Files) - the same order can legitimately
   show a DIFFERENT status/amount in a later export (Undelivered ->
   Delivered, a settlement correction, etc.), so a later upload should be
   treated as an UPDATE to what was previously recorded, never silently
   ignored as "just a duplicate". See split_new_vs_updated().

Both patterns share one building block: a composite "row key" built from
whichever columns the client identified as uniquely identifying a real-
world record (e.g. Order ID + SKU + Transaction Type for MTR; Order ID +
Transaction Type for gateway/settlement files) - see build_row_key().
"""

import pandas as pd

# Common header spellings for a delivery partner's AWB/tracking number - no
# config currently defines this explicitly (see configs/*.json's
# delivery_partners blocks), so it's resolved best-effort by trying every
# alias here rather than requiring a config change just to get the
# Order ID + AWB duplicate/updated-record key the client asked for. Falls
# back to Order ID alone (whichever caller-supplied key column resolves)
# when none of these match - still correct, just slightly coarser (can't
# tell two shipments on the same order apart). Shared between
# views/page_upload.py (upload-time preview) and views/page_reconciliation.py
# (ledger update on save) so both sides use the exact same alias list.
AWB_COLUMN_ALIASES = [
    "AWB", "AWB Number", "AWB No", "AWB No.", "Waybill", "Waybill Number",
    "Tracking Number", "Tracking ID", "Tracking Number/AWB", "Courier AWB",
]


def build_row_key(df, key_cols):
    """
    One string key per row, built by joining the given columns (already-
    resolved actual column names - see engine.loaders.resolve_col - not
    raw config specs) with a separator that won't collide with real
    values. A blank/missing piece becomes "" rather than raising, so a key
    is still computed - and still comparable across uploads - even when
    one of the key columns is sparsely populated or (for an optional key
    column like an AWB number) entirely absent from this particular file.
    """
    n = len(df)
    parts = []
    for col in key_cols:
        if col and col in df.columns:
            series = df[col]
            if pd.api.types.is_numeric_dtype(series):
                # Round before stringifying, not just astype(str) - a
                # numeric key column (e.g. an amount, used to disambiguate
                # multiple settlement rows sharing the same order+type -
                # see the gateway accumulate key in views/page_upload.py)
                # can otherwise produce a DIFFERENT key text for the exact
                # same real value purely from how two separate exports
                # happened to be parsed (500 vs 500.0 vs 500.00), which
                # would defeat the whole point of including it: two rows
                # for the genuinely same settlement would stop matching
                # and start looking like two different ones.
                s = series.round(2).astype(str)
            else:
                s = series.astype(str).str.strip()
            s = s.where(~s.str.lower().isin(["nan", "none", "nat", ""]), "")
        else:
            s = pd.Series([""] * n, index=df.index)
        parts.append(s)
    key = parts[0]
    for p in parts[1:]:
        key = key.str.cat(p, sep="||")
    return key


def split_new_vs_seen(df, key_cols, seen_keys):
    """
    Sales/order-level pattern (see module docstring, case 1). Returns
    (new_df, duplicate_df) - rows whose composite key is already in
    `seen_keys` (built from every previously-saved period for this
    client/report - see engine.storage's ledger functions) are split into
    duplicate_df; only new_df should be reconciled/saved as this upload's
    contribution.
    """
    if df is None or df.empty:
        return df, (df.iloc[0:0] if df is not None else df)
    keys = build_row_key(df, key_cols)
    is_dup = keys.isin(seen_keys)
    return df[~is_dup].copy(), df[is_dup].copy()


def accumulate_df(existing_df, new_df, key_cols=None):
    """
    Merges new_df into existing_df so an upload slot's working dataset
    ACCUMULATES across separate upload events in the same session, instead
    of each new upload wholesale replacing whatever was already there.

    Client-reported (2026-08-21): every upload slot on the Upload Data page
    used to overwrite its own st.session_state entry with EXACTLY what
    st.file_uploader currently returns on THIS render - fine the first
    time, but wrong the moment the user revisits the page later to add
    just one more report: a file_uploader widget that isn't touched on a
    given visit returns nothing (Streamlit doesn't keep a widget's
    previous file "active" once the page has been navigated away from and
    back to), and the old unconditional-overwrite code wrote that
    "nothing" straight into session_state - silently deleting a source
    that had already been uploaded and confirmed earlier. First reported
    as Orders ("No orders file uploaded yet" reappearing after only
    adding a missing gateway file), but the exact same pattern existed at
    every other upload slot (delivery partners, gateways, bank statement,
    Amazon MTR/Settlement Flat Files) - which is also why uploading just
    one later month's Shopify data was silently dropping the Delivery
    Partner/Payment Gateway data from earlier in the session, even though
    neither of those uploaders was touched that time, and the Payment
    Gateway/Delivery Partner reconciliation layers then never reflected
    the newly-added month (views/page_reconciliation.py already recomputes
    every layer fresh from session_state on every "Run reconciliation"
    click - the bug was entirely in what upload made it into session_state
    to recompute FROM, not in the reconciliation logic itself).

    Fixed at this one shared primitive, used at every upload slot (see
    views/page_upload.py): a new upload never replaces the whole existing
    dataset - it's merged in by key instead:
      - A row in new_df whose key matches a row already in existing_df
        REPLACES that row (the newer upload is assumed correct for
        whatever it explicitly reports - a genuine status/amount
        correction for the same record, e.g. Undelivered -> Delivered).
      - A row in new_df with a key not seen before is simply ADDED.
      - Every row in existing_df whose key ISN'T touched by new_df is left
        exactly as it was.
    This is a strict superset of both "a full re-upload of the same wider
    export replaces every key it still contains" and "a separate month's
    file with no overlapping keys is pure addition" - one rule handles
    both without the caller needing to know which case it is.

    key_cols: resolved actual column names (see engine.loaders.resolve_col -
    NOT raw config specs) identifying a row uniquely, same convention as
    build_row_key() above. When no usable key column resolves on both
    sides (e.g. a raw bank statement export with no natural single ID
    column), falls back to a hash of each row's own full content - this
    still makes re-uploading the exact same file a safe no-op, while any
    row that's even slightly different (a genuinely new transaction) is
    still added.
    """
    if new_df is None or new_df.empty:
        return existing_df
    if existing_df is None or existing_df.empty:
        return new_df.reset_index(drop=True)

    resolved_keys = [c for c in (key_cols or []) if c and c in existing_df.columns and c in new_df.columns]
    if resolved_keys:
        existing_key = build_row_key(existing_df, resolved_keys)
        new_key = build_row_key(new_df, resolved_keys)
    else:
        existing_key = pd.util.hash_pandas_object(existing_df.astype(str), index=False).astype(str)
        new_key = pd.util.hash_pandas_object(new_df.astype(str), index=False).astype(str)

    keep_existing = existing_df[~existing_key.isin(set(new_key))]
    return pd.concat([keep_existing, new_df], ignore_index=True, sort=False)


def split_new_vs_updated(df, key_cols, seen_keys):
    """
    Status/settlement pattern (see module docstring, case 2). Returns
    (new_df, updated_df) - both subsets of the SAME upload, split purely
    for the upload-time preview ("N brand-new records, M records whose
    status/amount is changing from what a previous upload said"). Nothing
    is excluded here - the newest upload's data is what gets used for
    every one of these rows either way (that IS the "latest wins" rule);
    this split only tells the user which bucket each row falls in before
    they confirm.
    """
    if df is None or df.empty:
        empty = df.iloc[0:0] if df is not None else df
        return df, empty
    keys = build_row_key(df, key_cols)
    is_update = keys.isin(seen_keys)
    return df[~is_update].copy(), df[is_update].copy()
