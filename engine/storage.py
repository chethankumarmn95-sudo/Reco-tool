"""
storage.py
----------
Lets you save a reconciled month's results to disk and come back to it
later, without re-uploading files or losing last month's work. This is
what makes the tool "month-wise" rather than one-shot.

Storage is just pickled files on your own machine, one per saved month,
under data_store/<client>/<month_label>.pkl next to the app. Nothing
leaves your laptop.

Performance note (the "tool gets super slow after uploading bulk data"
report): list_runs()/list_amazon_runs() are called on EVERY render of the
Dashboard, Reports, Data Management, and Bank Linking pages - and in
Streamlit, every single widget click anywhere on those pages re-runs the
whole page top to bottom, so these functions fire constantly, not just
once per visit. They used to open and pickle.load() the FULL saved
payload - every dataframe in it, including a lakhs-of-rows Expense Ledger
- just to read 7 small metadata fields (month_label/saved_at/order_count/
date_min/date_max/financial_year/channel_name) for a dropdown label. For a
bulk upload that saves a multi-MB (or, per one real saved period seen in
this engagement, tens-of-MB) pickle, that meant fully deserializing that
whole file on every click, everywhere in the app. Fixed by writing a tiny
JSON "sidecar" file (<label>.meta.json) alongside each .pkl at save time,
containing ONLY those metadata fields - list_runs()/list_amazon_runs() now
read the sidecar (near-instant) instead of the full pickle. A saved run
from before this fix has no sidecar yet; it's read once via the old slow
path and a sidecar is written immediately after, so it's fast on every
listing from then on ("self-healing", no manual migration step needed).
"""

import os
import json
import pickle
import shutil
import uuid
import datetime as dt

STORE_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data_store")

# How long a soft-deleted item sits in the Recycle Bin before it's purged
# for good (see the "Recycle Bin / Soft Delete" section near the bottom of
# this file) - the client's own explicit "such as 30 days" suggestion.
RECYCLE_BIN_RETENTION_DAYS = 30


def _client_dir(client_key):
    safe_key = "".join(c if c.isalnum() or c in "-_" else "_" for c in client_key)
    path = os.path.join(STORE_ROOT, safe_key)
    os.makedirs(path, exist_ok=True)
    return path


def _safe_filename(month_label):
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in month_label).strip()
    return safe.replace(" ", "_") + ".pkl"


def _meta_path(pkl_path):
    return pkl_path[:-4] + ".meta.json" if pkl_path.endswith(".pkl") else pkl_path + ".meta.json"


def _write_meta_sidecar(pkl_path, meta):
    """Best-effort - a sidecar write failure (e.g. a locked/read-only
    folder) should never block the actual save, it just means that one
    listing falls back to the slower full-unpickle path next time."""
    try:
        with open(_meta_path(pkl_path), "w") as f:
            json.dump(meta, f)
    except Exception:
        pass


def _read_meta_sidecar(pkl_path):
    mp = _meta_path(pkl_path)
    if not os.path.exists(mp):
        return None
    try:
        with open(mp) as f:
            return json.load(f)
    except Exception:
        return None


_META_FIELDS = ["month_label", "saved_at", "order_count", "date_min", "date_max", "financial_year", "channel_name"]


def _meta_from_payload(payload, kind):
    return {"kind": kind, **{k: payload.get(k) for k in _META_FIELDS}}


def _list_saved(client_key, want_kind):
    """
    Shared listing logic for list_runs()/list_amazon_runs(): reads each
    saved period's metadata via its sidecar file when available (fast -
    see module docstring), falling back to a full pickle.load() only for
    older saved files that predate the sidecar feature - and writes the
    sidecar immediately after that fallback so the NEXT listing is fast
    too. want_kind=None means "don't filter by kind" (the DTC side, where
    every file in this client_key's folder is already known to be a DTC
    payload - see save_run()'s "no collision risk" note).
    """
    folder = _client_dir(client_key)
    runs = []
    for fname in os.listdir(folder):
        if not fname.endswith(".pkl"):
            continue
        path = os.path.join(folder, fname)
        meta = _read_meta_sidecar(path)
        if meta is None:
            try:
                with open(path, "rb") as f:
                    payload = pickle.load(f)
            except Exception:
                continue
            meta = _meta_from_payload(payload, payload.get("kind", "dtc"))
            _write_meta_sidecar(path, meta)  # self-heal so next time is fast
        if want_kind is not None and meta.get("kind") != want_kind:
            continue
        runs.append({
            "file": fname,
            "month_label": meta.get("month_label", fname),
            "saved_at": meta.get("saved_at", ""),
            "order_count": meta.get("order_count", 0),
            "date_min": meta.get("date_min"),
            "date_max": meta.get("date_max"),
            "financial_year": meta.get("financial_year", "Unknown FY"),
            "channel_name": meta.get("channel_name"),
        })
    runs.sort(key=lambda r: r["saved_at"], reverse=True)
    return runs


def financial_year_label(date):
    """
    Indian financial year: 1 April to 31 March.
    A date in April 2026 through March 2027 is "FY 2026-27".
    """
    import pandas as pd
    ts = pd.to_datetime(date, errors="coerce")
    if pd.isna(ts):
        return "Unknown FY"
    year = ts.year if ts.month >= 4 else ts.year - 1
    return f"FY {year}-{str(year + 1)[-2:]}"


def financial_year_end_date(fy_label):
    """
    "FY 2026-27" -> Timestamp(2027-03-31) - the last calendar day of the
    financial year this label names. "Unknown FY" or any unparseable label
    returns None rather than guessing.

    Added 2026-08-23 (client-reported spec) for the Reports page: when the
    user leaves the "To date" filter blank, the report covers the WHOLE
    selected financial year - but engine.bank.bank_reconciliation_by_utr's
    new "Settled Subsequent Period" logic still needs a real cutoff date to
    compare each bank credit's own date against (see that function's
    period_end_date parameter). Without this, "no To date picked" would
    have to mean "no cutoff at all", which would wrongly swallow a bank
    credit from a LATER financial year into "this period" the moment this
    engine's cross-period broadening (views/page_reports.py's
    _settlement_runs_for_bank_matching) is asked to search beyond the
    selected FY for a previous-period settlement.
    """
    import re
    import pandas as pd
    if not fy_label:
        return None
    match = re.match(r"FY (\d{4})-(\d{2})", str(fy_label))
    if not match:
        return None
    start_year = int(match.group(1))
    return pd.Timestamp(year=start_year + 1, month=3, day=31)


def financial_year_start_date(fy_label):
    """
    "FY 2026-27" -> Timestamp(2026-04-01) - the first calendar day of the
    financial year this label names. Sibling of financial_year_end_date()
    above, added 2026-08-31 (client-reported, Executive Summary Point 7 -
    "Receivable Collection Period Analysis"): that section's own
    reconciliation now needs a genuine LOWER bound on "the selected
    period" too, not just the upper bound financial_year_end_date() was
    built for - see engine.bank.bank_reconciliation_by_utr's new
    period_start_date parameter and this module's own callers in
    views/page_reports.py for the full story. Same "Unknown FY" or
    unparseable label -> None (never guessed) convention.
    """
    import re
    import pandas as pd
    if not fy_label:
        return None
    match = re.match(r"FY (\d{4})-(\d{2})", str(fy_label))
    if not match:
        return None
    start_year = int(match.group(1))
    return pd.Timestamp(year=start_year, month=4, day=1)


def financial_year_labels(dates):
    """
    Vectorized sibling of financial_year_label() above - takes a whole
    Series of dates and returns a Series of FY label strings in one pass
    (no per-row Python function call), for splitting a whole DataFrame's
    rows by financial year at once. Same rule, same "Unknown FY" fallback
    for anything unparseable/missing.

    Used by views/page_reconciliation.py to split a single reconciliation
    run into one saved period PER financial year before saving - the fix
    for "I uploaded FY 2025-26 data, then FY 2026-27 data, but both got
    combined into one report": if a single upload/run's own date range
    happens to straddle 1-April, this is what keeps the two financial
    years from ever landing in the same saved period file in the first
    place (see save_run()/save_amazon_run()'s docstrings for where this
    gets used).
    """
    import pandas as pd
    ts = pd.to_datetime(dates, errors="coerce")
    years = ts.dt.year.where(ts.dt.month >= 4, ts.dt.year - 1)
    labels = "FY " + years.astype("Int64").astype(str) + "-" + (years + 1).astype("Int64").astype(str).str[-2:]
    return labels.where(ts.notna(), "Unknown FY")


def month_labels(dates):
    """
    Vectorized per-row sibling of detect_month_label() below - one
    "Month YYYY" string per row (e.g. "April 2026"), instead of a single
    best-guess label for a whole dataframe. Used to split ONE reconciliation
    run's data into genuinely disjoint MONTHLY saved periods (see
    views/page_reconciliation.py's run_dtc_reconciliation/
    run_marketplace_reconciliation auto-save loops) - a calendar month is
    always entirely inside one financial year, so grouping by this instead
    of financial_year_labels() keeps the existing "never mix financial
    years in one saved period" guarantee AND adds "never mix calendar
    months in one saved period" as a strict refinement, not a separate rule.

    Client-reported (2026-08-22): before this existed, that auto-save loop
    only split by FINANCIAL YEAR - so uploading April's orders, running
    reconciliation (saved fine, one period, "April 2026"), then LATER
    uploading May's orders on top (orders_df accumulates across uploads by
    design - see engine.dedup.accumulate_df) and running again saved the
    ENTIRE now-April+May reco_df as ONE "period" again, since April and May
    both fall in the same financial year and the loop never split any
    finer than that - labelled with whatever month happened to be the
    MODE across that combined data (see detect_month_label), which could
    easily come out as "May 2026" purely because May had slightly more
    rows than April. Every subsequent month's run compounded the same way
    (a "June 2026" save that was actually April+May+June combined), so
    three saved periods that LOOKED disjoint by label were actually
    cumulative supersets of one another - and Dashboard/Reports'
    combine_runs (a plain concatenation across whichever saved periods are
    selected, with no cross-period dedup - it exists specifically to
    combine genuinely DISJOINT months) then double-, triple-, ...-counted
    every order saved under more than one label, while whichever period a
    filter happened to exclude looked like its data had vanished entirely
    - even though it existed, just buried inside a LATER period's
    cumulative snapshot instead of its own.

    Splitting the save loop by this function's month-level key instead
    fixes it at the source: each saved period ends up containing only the
    orders that actually belong to that calendar month, so two saved
    periods can never legitimately overlap on the same order_id again
    (barring a genuine data anomaly upstream), and combine_runs' simple
    concatenation is correct for whatever combination of saved months gets
    selected.
    """
    import pandas as pd
    ts = pd.to_datetime(dates, errors="coerce")
    labels = ts.dt.strftime("%B %Y")
    return labels.where(ts.notna(), "Unknown Month")


def detect_month_label(reco_df):
    """Best-guess default label for a run, from the data itself."""
    import pandas as pd
    try:
        dates = pd.to_datetime(reco_df["created_at"], errors="coerce")
        return dates.dt.strftime("%B %Y").mode().iloc[0]
    except Exception:
        return "Untitled run"


def save_run(client_key, month_label, reco_df, lookup_df, totals, channel_name=None,
             consolidated_df=None, bank_ledger_df=None):
    """
    consolidated_df / bank_ledger_df are optional additions (engine.consolidator
    build_consolidated_receipt() output / engine.bank load_bank_statement()
    output) - saved alongside reco_df/lookup_df so the Payment Gateway
    Settlement Report and the UTR-level bank reconciliation can be
    recomputed correctly over a COMBINED, multi-month date range on the
    Reports page (see engine.settlement / engine.bank), not just for the
    single month that was just run. Pass None if a bank statement wasn't
    uploaded this run - that's a normal, supported case (bank_statement is
    optional), it just means this saved month contributes nothing to a
    later combined bank reconciliation.
    """
    import pandas as pd

    dates = pd.to_datetime(reco_df["created_at"], errors="coerce")
    date_min = dates.min()
    date_max = dates.max()
    fy = financial_year_label(date_min) if pd.notna(date_min) else "Unknown FY"

    path = os.path.join(_client_dir(client_key), _safe_filename(month_label))
    payload = {
        "kind": "dtc",
        "month_label": month_label,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
        "reco_df": reco_df,
        "lookup_df": lookup_df,
        "totals": totals,
        "order_count": len(reco_df),
        "date_min": date_min.isoformat() if pd.notna(date_min) else None,
        "date_max": date_max.isoformat() if pd.notna(date_max) else None,
        "financial_year": fy,
        "channel_name": channel_name,
        "consolidated_df": consolidated_df,
        "bank_ledger_df": bank_ledger_df,
    }
    with open(path, "wb") as f:
        pickle.dump(payload, f)
    _write_meta_sidecar(path, _meta_from_payload(payload, "dtc"))
    return path


def list_runs(client_key):
    """Returns metadata for every saved month, newest first - reads each
    saved period's lightweight .meta.json sidecar rather than loading the
    (potentially large) dataframes into memory - see module docstring."""
    return _list_saved(client_key, want_kind=None)


def load_run(client_key, fname):
    path = os.path.join(_client_dir(client_key), fname)
    with open(path, "rb") as f:
        return pickle.load(f)


def delete_run(client_key, fname):
    path = os.path.join(_client_dir(client_key), fname)
    existed = False
    if os.path.exists(path):
        os.remove(path)
        existed = True
    meta_path = _meta_path(path)
    if os.path.exists(meta_path):
        os.remove(meta_path)
    return existed


def detect_amazon_month_label(order_reco_df):
    """detect_month_label()'s equivalent for the Amazon/marketplace payload
    shape - order_reco_df uses an "order_date" column, not "created_at"."""
    import pandas as pd
    try:
        dates = pd.to_datetime(order_reco_df["order_date"], errors="coerce")
        return dates.dt.strftime("%B %Y").mode().iloc[0]
    except Exception:
        return "Untitled period"


def save_amazon_run(client_key, month_label, order_reco_df, waterfall_df, expense_ledger_df,
                     settlement_summary_df, settlement_register_df, non_mtr_df, tie_out_df,
                     channel_name=None, subsequent_settlements_df=None, cutoff_date=None):
    """
    Amazon-channel equivalent of save_run() above - same one-pickle-per-
    saved-period storage under data_store/<client>/<label>.pkl, but for the
    marketplace-shaped result set (see engine/amazon_reco.py and
    engine/amazon_consolidator.py) instead of the DTC reco_df/lookup_df
    shape. Kept as separate functions rather than overloading save_run()
    itself, since the two payload shapes share almost nothing (no
    order_id-keyed reco_df, no "query" column, a settlement-id grain
    instead of a gateway/UTR grain, etc.) - forcing them through one
    function would just mean a pile of "if this is amazon" branches inside
    it. No collision risk with the DTC list_runs()/combine_runs() readers:
    they're always called with a DIFFERENT client_key (e.g. "esca_shopify"),
    which points at a different data_store/<client>/ folder entirely.
    """
    import pandas as pd
    dates = pd.to_datetime(order_reco_df["order_date"], errors="coerce") if order_reco_df is not None and not order_reco_df.empty else pd.Series([], dtype="datetime64[ns]")
    date_min = dates.min() if len(dates) else pd.NaT
    date_max = dates.max() if len(dates) else pd.NaT
    fy = financial_year_label(date_min) if pd.notna(date_min) else "Unknown FY"

    path = os.path.join(_client_dir(client_key), _safe_filename(month_label))
    payload = {
        "kind": "amazon",
        "month_label": month_label,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
        "order_reco_df": order_reco_df,
        "waterfall_df": waterfall_df,
        "expense_ledger_df": expense_ledger_df,
        "settlement_summary_df": settlement_summary_df,
        "settlement_register_df": settlement_register_df,
        "non_mtr_df": non_mtr_df,
        "tie_out_df": tie_out_df,
        "subsequent_settlements_df": subsequent_settlements_df,
        "cutoff_date": cutoff_date.isoformat() if cutoff_date is not None else None,
        "order_count": len(order_reco_df) if order_reco_df is not None else 0,
        "date_min": date_min.isoformat() if pd.notna(date_min) else None,
        "date_max": date_max.isoformat() if pd.notna(date_max) else None,
        "financial_year": fy,
        "channel_name": channel_name,
    }
    with open(path, "wb") as f:
        pickle.dump(payload, f)
    _write_meta_sidecar(path, _meta_from_payload(payload, "amazon"))
    return path


def list_amazon_runs(client_key):
    """Amazon equivalent of list_runs() - deliberately returns the SAME
    metadata shape (file/month_label/saved_at/order_count/date_min/
    date_max/financial_year/channel_name) so the existing filter widgets in
    views/filters.py (render_dashboard_filter_controls/
    render_filter_controls) work completely unmodified for the marketplace
    channel too. Reads each saved period's lightweight .meta.json sidecar
    rather than loading the full ledger into memory - see module
    docstring (this is the fix for the "tool gets super slow after
    uploading bulk data" report: this function runs on every single click
    on Dashboard/Reports/Data Management/Bank Linking, so it can never
    afford to fully unpickle a lakhs-of-rows Expense Ledger just to read a
    handful of small metadata fields)."""
    return _list_saved(client_key, want_kind="amazon")


def combine_amazon_runs(client_key, fnames):
    """
    Loads and concatenates multiple saved Amazon periods into one combined
    result set, for cross-period Dashboard/Reports viewing - mirrors
    combine_runs() above but returns a dict (the Amazon payload has many
    more frames than reco_df/lookup_df, a dict reads clearer at the call
    site than a 7-tuple).

    The order/ledger-grain frames (order_reco_df, expense_ledger_df,
    settlement_summary_df, settlement_register_df, non_mtr_df, tie_out_df)
    are simply concatenated - safe as long as the selected saved periods
    don't overlap in order-id/settlement-id, same assumption combine_runs()
    already makes for order_id on the DTC side.

    waterfall_df is NOT recomputed from raw MTR/ledger data (that would
    need re-loading every underlying MTR/flat-file frame, which isn't
    saved - it's large and fully reconstructable from the ledger). Instead:
    every row in a single period's waterfall is already a pure sum of that
    period's figures, and the "Payable"/"Receivable"/"Balance Receivable"
    subtotal rows are just linear combinations of those - so summing the
    Amount column by Particular label across periods gives exactly the
    same answer build_waterfall() would over the combined raw data.
    """
    import pandas as pd
    order_reco_parts, ledger_parts, summary_parts = [], [], []
    register_parts, non_mtr_parts, tie_out_parts, waterfall_parts = [], [], [], []
    subsequent_parts = []
    cutoff_dates = []

    for fname in fnames:
        payload = load_run(client_key, fname)
        if payload.get("cutoff_date"):
            cutoff_dates.append(pd.to_datetime(payload["cutoff_date"], errors="coerce"))
        if payload.get("order_reco_df") is not None:
            order_reco_parts.append(payload["order_reco_df"])
        if payload.get("expense_ledger_df") is not None:
            ledger_parts.append(payload["expense_ledger_df"])
        if payload.get("settlement_summary_df") is not None:
            summary_parts.append(payload["settlement_summary_df"])
        if payload.get("settlement_register_df") is not None:
            register_parts.append(payload["settlement_register_df"])
        if payload.get("non_mtr_df") is not None:
            non_mtr_parts.append(payload["non_mtr_df"])
        if payload.get("tie_out_df") is not None:
            tie_out_parts.append(payload["tie_out_df"])
        if payload.get("subsequent_settlements_df") is not None:
            subsequent_parts.append(payload["subsequent_settlements_df"])
        if payload.get("waterfall_df") is not None:
            waterfall_parts.append(payload["waterfall_df"])

    def _cat(parts):
        return pd.concat(parts, ignore_index=True) if parts else None

    combined_waterfall = None
    if waterfall_parts:
        # Row order: the union of every period's Particular labels, in
        # first-seen order - NOT just waterfall_parts[0]'s own rows. Some
        # memo rows (e.g. "Subsequent Settlements", "Receipts in transit" -
        # see engine/amazon_reco.py's build_waterfall) only appear in a
        # period's waterfall when that figure is non-zero for that period,
        # so the first saved period alone might be missing a row a later
        # period has - reindexing to only the first period's rows would
        # silently drop that row's total from the combined view.
        seen = []
        for wf in waterfall_parts:
            for p in wf["Particular"]:
                if p not in seen:
                    seen.append(p)
        combined_waterfall = (
            pd.concat(waterfall_parts, ignore_index=True)
            .groupby("Particular", sort=False)["Amount"].sum()
            .reindex(seen)
            .reset_index()
        )

    valid_cutoffs = [c for c in cutoff_dates if pd.notna(c)]

    return {
        "order_reco_df": _cat(order_reco_parts),
        "expense_ledger_df": _cat(ledger_parts),
        "settlement_summary_df": _cat(summary_parts),
        "settlement_register_df": _cat(register_parts),
        "non_mtr_df": _cat(non_mtr_parts),
        "tie_out_df": _cat(tie_out_parts),
        "subsequent_settlements_df": _cat(subsequent_parts),
        "waterfall_df": combined_waterfall,
        # The LATEST cut-off among the selected saved periods - shown on the
        # downloaded workbook's Dashboard sheet subheader for context when
        # multiple periods are combined; each individual period's own
        # cut-off is still visible via its own saved payload if needed.
        "cutoff_date": max(valid_cutoffs) if valid_cutoffs else None,
    }


def combine_runs(client_key, fnames, include_settlement=False):
    """Loads and concatenates multiple saved months into one combined
    reco_df/lookup_df, for cross-month reporting.

    include_settlement=True additionally combines each saved month's
    consolidated_df/bank_ledger_df (see save_run()) and returns them as a
    4th/5th... no - a 3rd/4th value, for the Gateway Settlement Report and
    UTR-level bank reconciliation on the Reports page. Older saved months
    (from before this feature existed) simply have no consolidated_df/
    bank_ledger_df to contribute - they're skipped rather than erroring.
    """
    import pandas as pd
    reco_parts, lookup_parts, consolidated_parts, bank_parts = [], [], [], []
    for fname in fnames:
        payload = load_run(client_key, fname)
        reco_parts.append(payload["reco_df"])
        lookup_parts.append(payload["lookup_df"])
        if include_settlement:
            if payload.get("consolidated_df") is not None:
                consolidated_parts.append(payload["consolidated_df"])
            if payload.get("bank_ledger_df") is not None:
                bank_parts.append(payload["bank_ledger_df"])

    combined_reco = pd.concat(reco_parts, ignore_index=True) if reco_parts else None
    combined_lookup = pd.concat(lookup_parts, ignore_index=True) if lookup_parts else None

    if not include_settlement:
        return combined_reco, combined_lookup

    combined_consolidated = pd.concat(consolidated_parts, ignore_index=True) if consolidated_parts else None
    combined_bank_ledger = pd.concat(bank_parts, ignore_index=True) if bank_parts else None
    return combined_reco, combined_lookup, combined_consolidated, combined_bank_ledger


# ---------------------------------------------------------------------------
# Ingestion ledger - "have we already seen this row before" across uploads
# ---------------------------------------------------------------------------
# Two small, deliberately lightweight JSON files per client (NOT part of the
# per-period .pkl payloads above) power the Upload Data page's duplicate/
# updated-record detection described in engine/dedup.py's module docstring:
#
#   _seen_keys__<report_key>.json    - a flat list of composite keys already
#                                       reconciled, for SALES/order-level
#                                       reports (skip-duplicate). Grows by
#                                       simple union on every save - a key
#                                       already reconciled once stays
#                                       "seen" forever.
#   _latest_records__<report_key>.json - {key: {...last-known field values,
#                                       "order_id":..., "period_file":...}}
#                                       for STATUS/SETTLEMENT reports
#                                       (latest-wins) - each save simply
#                                       overwrites whatever was there before
#                                       for that key.
#
# Both are plain JSON (not pickled), so they stay legible if anyone ever
# needs to inspect or hand-edit one, and they're independent of the
# per-period .pkl files - deleting/re-saving a period doesn't touch these,
# and these existing (or not) never blocks loading a period.


def _ledger_dir(client_key):
    path = os.path.join(_client_dir(client_key), "_ledger")
    os.makedirs(path, exist_ok=True)
    return path


def _seen_keys_path(client_key, report_key):
    return os.path.join(_ledger_dir(client_key), f"seen_keys__{report_key}.json")


def _latest_records_path(client_key, report_key):
    return os.path.join(_ledger_dir(client_key), f"latest_records__{report_key}.json")


def get_seen_keys(client_key, report_key):
    """The set of composite keys already reconciled for this report type
    (see engine.dedup.split_new_vs_seen) - empty set if nothing's been
    recorded yet (e.g. the very first upload for this client/report)."""
    path = _seen_keys_path(client_key, report_key)
    if not os.path.exists(path):
        return set()
    try:
        with open(path) as f:
            return set(json.load(f))
    except Exception:
        return set()


def add_seen_keys(client_key, report_key, keys):
    """Merges newly-reconciled keys into the ledger (plain union - a key
    already recorded just stays recorded). Call this once a batch of rows
    has actually been reconciled/saved, not merely uploaded, so a preview
    the user never confirms can't mark rows as "seen" that were never
    really processed."""
    existing = get_seen_keys(client_key, report_key)
    existing.update(str(k) for k in keys if k)
    try:
        with open(_seen_keys_path(client_key, report_key), "w") as f:
            json.dump(sorted(existing), f)
    except Exception:
        pass


def remove_seen_keys_for_order_ids(client_key, report_key, order_ids):
    """
    Removes exactly the seen_keys entries that belong to the given
    order_ids, leaving every other order's "already reconciled" record
    untouched - unlike clear_seen_keys() above (which wipes the WHOLE
    ledger for a report_key), this is what a Recycle Bin move needs for a
    SAVED PERIOD: the "orders"/"mtr__<segment>" ledgers accumulate across
    every period ever saved for that report type (see this module's
    ingestion-ledger section docstring - "a key already reconciled once
    stays seen forever"), so deleting just ONE saved period must only
    forget THAT period's own order_ids, not every order this client has
    ever reconciled.

    Works without needing any period/batch tagging at write time: every
    "skip"-mode composite key in this app is built with build_row_key()
    (see engine/dedup.py), which always puts the order id FIRST and joins
    pieces with "||" - true for both the single-column "orders" key (the
    order id alone, no "||" at all) and the 3-column "mtr__<segment>" key
    ("order_id||sku||transaction_type"). Splitting each stored key on the
    first "||" and comparing that first piece against order_ids therefore
    correctly isolates exactly this period's rows, however many other key
    pieces exist, with no separate bookkeeping required.

    IMPORTANT: the comparison runs through engine.loaders.normalize_order_id
    on both sides, not a plain string match. A seen_keys entry's order-id
    piece is built from the RAW upload's own order-id column (e.g. "#1001",
    straight from build_row_key() on the untouched orders_df/mtr_df) via
    engine.reco._update_dtc_ledger()/page_reconciliation._update_amazon_ledger()
    - but the order_ids a caller passes in here come from a SAVED PERIOD's
    reco_df/order_reco_df, whose own "order_id" column has already been run
    through normalize_order_id() (build_order_master() strips "#", leading/
    trailing space, and a stray ".0" - see that function's docstring) since
    that's what every join in the app keys on downstream. Comparing "#1001"
    against "1001" with a plain string match would silently match nothing
    at all - normalizing both sides the same way here is what makes this
    actually work, not an approximation of it.

    Returns the exact set of keys removed (empty set if none matched) -
    the caller (soft_delete_run below) records this in the Recycle Bin
    item's manifest so restore_run_from_recycle_bin() can put precisely
    these keys back if the deletion is undone.
    """
    import pandas as pd
    from .loaders import normalize_order_id

    order_id_set = set(normalize_order_id(pd.Series(list(order_ids), dtype=str))) if order_ids else set()
    if not order_id_set:
        return set()
    existing = get_seen_keys(client_key, report_key)
    if not existing:
        return set()
    existing_list = sorted(existing)
    prefixes = pd.Series([k.split("||", 1)[0] for k in existing_list])
    normalized_prefixes = normalize_order_id(prefixes)
    to_remove = {k for k, norm in zip(existing_list, normalized_prefixes) if norm in order_id_set}
    if not to_remove:
        return set()
    remaining = existing - to_remove
    try:
        with open(_seen_keys_path(client_key, report_key), "w") as f:
            json.dump(sorted(remaining), f)
    except Exception:
        pass
    return to_remove


def clear_seen_keys(client_key, report_key):
    """
    Deletes the "already reconciled" ledger for one specific report_key
    (see get_seen_keys/add_seen_keys above).

    Client-reported (2026-08-21): deleting an individual raw report (e.g.
    just the Razorpay gateway file, or an Orders/MTR file) from Data
    Management must also forget that this report type's rows were ever
    reconciled - otherwise a later re-upload of the exact same file still
    shows "X of X rows match data already reconciled - 0 new rows will be
    processed" (see views/page_upload.py's render_duplicate_check, mode=
    "skip"), even though every row this ledger was tracking has just been
    removed from the session entirely. Called by
    views/page_data_management.py's individual-file-delete action right
    after the file's data is removed from session_state, so the two never
    drift out of sync (ledger entries for rows nothing in the session can
    still point to).

    Returns True if a ledger file actually existed and was removed, False
    if there was nothing to clear (e.g. this report type was never
    reconciled/saved in the first place - deleting it is then a pure
    session-state no-op with nothing to clean up on disk).
    """
    path = _seen_keys_path(client_key, report_key)
    existed = os.path.exists(path)
    if existed:
        os.remove(path)
    return existed


def clear_latest_records(client_key, report_key):
    """
    clear_seen_keys()'s sibling for the "latest known values" ledger (see
    get_latest_records/record_latest_records) used by latest-wins report
    types - delivery partners, payment gateways, Amazon Settlement Flat
    Files. Same reasoning and same caller (views/page_data_management.py's
    individual-file-delete action) - without this, deleting e.g. Razorpay
    and re-uploading it would still show it as "N record(s) updating
    status/amount from a previous upload" against a ledger entry with no
    corresponding data left anywhere in the session.
    """
    path = _latest_records_path(client_key, report_key)
    existed = os.path.exists(path)
    if existed:
        os.remove(path)
    return existed


def get_latest_records(client_key, report_key):
    """{key: {field: value, ...}} of the most-recently-uploaded data seen
    for each composite key of this report type (see
    engine.dedup.split_new_vs_updated) - empty dict if nothing recorded yet."""
    path = _latest_records_path(client_key, report_key)
    if not os.path.exists(path):
        return {}
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return {}


def record_latest_records(client_key, report_key, records):
    """Overwrites whichever keys are present in `records` (a {key: {...}}
    dict, values must be JSON-serialisable) with their newest values -
    call this once a batch has actually been reconciled/saved, same
    timing rule as add_seen_keys() above."""
    existing = get_latest_records(client_key, report_key)
    existing.update(records)
    try:
        with open(_latest_records_path(client_key, report_key), "w") as f:
            json.dump(existing, f, default=str)
    except Exception:
        pass


def find_dtc_runs_containing_orders(client_key, order_ids):
    """
    Which ALREADY-SAVED DTC periods (see save_run()) contain any of the
    given order_ids - used to find which saved period(s) a freshly-
    uploaded delivery-partner file's updated statuses actually belong to,
    so patch_run_delivery_status() below can be pointed at exactly the
    right file(s) instead of the caller having to guess a period label.

    Unlike list_runs(), this DOES fully unpickle every saved period (the
    lightweight .meta.json sidecar has no order-id-level detail to check
    against) - acceptable here because, unlike list_runs(), this only
    runs when the user explicitly asks "does this new delivery file
    update anything I've already saved", not on every page rerun.
    """
    if not order_ids:
        return []
    order_id_set = {str(o) for o in order_ids}
    matches = []
    for meta in list_runs(client_key):
        if meta.get("kind") not in (None, "dtc"):
            continue
        try:
            payload = load_run(client_key, meta["file"])
        except Exception:
            continue
        reco_df = payload.get("reco_df")
        if reco_df is None or reco_df.empty or "order_id" not in reco_df.columns:
            continue
        hit_count = reco_df["order_id"].astype(str).isin(order_id_set).sum()
        if hit_count:
            matches.append({
                "file": meta["file"], "month_label": meta.get("month_label", meta["file"]),
                "matching_orders": int(hit_count),
            })
    return matches


def patch_run_delivery_status(client_key, fname, delivery_frames, delivery_configs):
    """
    Re-saves an already-saved DTC period with delivery status refreshed
    for whichever of its orders appear in the freshly-uploaded
    delivery_frames (see engine.reco.refresh_delivery_status for exactly
    what does and doesn't get touched), recomputing totals/order_count
    from the patched reco_df, and preserving every other saved field
    (lookup_df, consolidated_df, bank_ledger_df, channel_name, etc.)
    exactly as they were. Returns the number of orders actually updated.
    """
    from .reco import refresh_delivery_status
    from .summary import headline_totals
    import pandas as pd

    payload = load_run(client_key, fname)
    reco_df = payload.get("reco_df")
    if reco_df is None or reco_df.empty:
        return 0

    patched, updated_count = refresh_delivery_status(reco_df, delivery_frames, delivery_configs)
    if updated_count == 0:
        return 0

    payload["reco_df"] = patched
    payload["totals"] = headline_totals(patched)
    payload["order_count"] = len(patched)
    payload["saved_at"] = dt.datetime.now().isoformat(timespec="seconds")

    path = os.path.join(_client_dir(client_key), fname)
    with open(path, "wb") as f:
        pickle.dump(payload, f)
    _write_meta_sidecar(path, _meta_from_payload(payload, "dtc"))
    return updated_count


# ---------------------------------------------------------------------------
# Recycle Bin / soft delete
# ---------------------------------------------------------------------------
# Client-reported (2026-08-22): after deleting a saved period from Data
# Management and starting a fresh upload, the SAME Shopify Order report
# still showed "6,197 of 6,197 row(s) match data already reconciled - 0
# new row(s) will actually be processed". Root cause: delete_run() above
# only ever removed the saved period's .pkl/.meta.json - it never touched
# the separate ingestion ledger (seen_keys__orders.json etc., see the
# section above) that render_duplicate_check() checks against, so every
# order id that period had ever reconciled stayed permanently "seen" even
# after the period itself was gone. On top of that, the client explicitly
# asked that deletion NOT be instant/permanent at all - a Recycle Bin, so
# an accidental delete is recoverable, with data automatically purged for
# good after a retention window (30 days by default) and, critically,
# completely invisible to reconciliation/duplicate-checking for as long as
# it sits in the bin.
#
# Every soft-deleted item - whether a whole saved period (kind="dtc"/
# "amazon", moved via soft_delete_run) or a single raw upload slot (kind=
# "raw_file", moved via soft_delete_raw_file - see
# views/page_data_management.py's individual-file-delete action) - lands
# in the SAME bin, one subfolder per item:
#
#   data_store/<client>/_recycle_bin/<item_id>/
#       manifest.json  - {item_id, kind, label, deleted_at, expires_at,
#                          ledger_removals: [{report_key, ledger_kind,
#                          keys: [...]}], ...display fields}
#       payload.pkl    - whatever's needed to put the data back exactly as
#                         it was: the full saved-period payload dict for a
#                         "dtc"/"amazon" item, or {"session_key",
#                         "dict_entry", "data": <DataFrame>} for a
#                         "raw_file" item.
#
# manifest.json's ledger_removals is what makes "invisible to reconciliation
# while in the bin, but restorable" possible without any separate ledger
# format change: the EXACT keys cleared out of the ledger at delete time
# are captured right there in the manifest, so restoring later just replays
# them back in (add_seen_keys/record_latest_records, both already
# idempotent unions) - no need to guess what a period "would have"
# contributed by re-deriving it from scratch.


def _recycle_bin_dir(client_key):
    path = os.path.join(_client_dir(client_key), "_recycle_bin")
    os.makedirs(path, exist_ok=True)
    return path


def _recycle_bin_item_dir(client_key, item_id):
    return os.path.join(_recycle_bin_dir(client_key), item_id)


def _new_recycle_bin_item(client_key, kind, label, extra_meta, payload, ledger_removals,
                           retention_days=RECYCLE_BIN_RETENTION_DAYS):
    """Shared plumbing for soft_delete_run()/soft_delete_raw_file() below -
    writes payload.pkl + manifest.json into a fresh item folder and returns
    the new item_id. ledger_removals: list of {"report_key", "ledger_kind"
    ("seen"|"latest"), "keys": [...]} - already-applied ledger removals
    being recorded here purely so restore can reverse them; this function
    itself does not touch the ledger."""
    item_id = f"{dt.datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}"
    item_dir = _recycle_bin_item_dir(client_key, item_id)
    os.makedirs(item_dir, exist_ok=True)

    deleted_at = dt.datetime.now()
    expires_at = deleted_at + dt.timedelta(days=retention_days)
    manifest = {
        "item_id": item_id,
        "kind": kind,
        "label": label,
        "deleted_at": deleted_at.isoformat(timespec="seconds"),
        "expires_at": expires_at.isoformat(timespec="seconds"),
        "ledger_removals": ledger_removals,
        **extra_meta,
    }
    with open(os.path.join(item_dir, "manifest.json"), "w") as f:
        json.dump(manifest, f, default=str)
    with open(os.path.join(item_dir, "payload.pkl"), "wb") as f:
        pickle.dump(payload, f)
    return item_id


def soft_delete_run(client_key, fname, report_keys_to_scope, retention_days=RECYCLE_BIN_RETENTION_DAYS):
    """
    Moves an entire saved period (see save_run()/save_amazon_run()) to the
    Recycle Bin instead of permanently deleting it - the "Delete entire
    dataset" flow in views/page_data_management.py now calls this instead
    of delete_run() directly.

    report_keys_to_scope: the "skip"-mode ledger report_keys relevant to
    this saved period's own channel - ["orders"] for a DTC/Shopify period,
    or one "mtr__<segment>" per configured MTR segment for an Amazon
    period (the caller supplies this since storage.py deliberately knows
    nothing about configs/*.json - see remove_seen_keys_for_order_ids'
    docstring for why order-id-prefix matching makes this precise without
    any config/segment-specific logic here). "Latest"-mode ledgers
    (delivery partners, gateways, Settlement Flat Files) are deliberately
    NOT scrubbed here: that mode never excludes any row from processing
    (see views/page_upload.py's render_duplicate_check docstring, mode=
    "latest") - it only affects informational upload-preview wording, so
    leaving old "latest" entries in place while a period sits in the bin
    has no effect on reconciliation correctness, and un-scrubbing them
    perfectly on restore isn't worth the added complexity.

    Returns the new Recycle Bin item_id, or None if fname doesn't exist.
    """
    path = os.path.join(_client_dir(client_key), fname)
    if not os.path.exists(path):
        return None
    payload = load_run(client_key, fname)

    kind = payload.get("kind", "dtc")
    if kind == "amazon":
        order_ids = set(payload["order_reco_df"]["order_id"].astype(str)) if payload.get("order_reco_df") is not None and not payload["order_reco_df"].empty else set()
    else:
        order_ids = set(payload["reco_df"]["order_id"].astype(str)) if payload.get("reco_df") is not None and not payload["reco_df"].empty else set()

    ledger_removals = []
    for report_key in report_keys_to_scope:
        removed = remove_seen_keys_for_order_ids(client_key, report_key, order_ids)
        if removed:
            ledger_removals.append({"report_key": report_key, "ledger_kind": "seen", "keys": sorted(removed)})

    item_id = _new_recycle_bin_item(
        client_key, kind, fname,
        extra_meta={
            "fname": fname,
            "month_label": payload.get("month_label", fname),
            "order_count": payload.get("order_count", 0),
            "channel_name": payload.get("channel_name"),
        },
        payload=payload, ledger_removals=ledger_removals, retention_days=retention_days,
    )
    delete_run(client_key, fname)
    return item_id


def soft_delete_raw_file(client_key, label, session_key, dict_entry, data,
                          ledger_removals, retention_days=RECYCLE_BIN_RETENTION_DAYS):
    """
    Recycle-Bin equivalent of an individual raw-file delete (see
    views/page_data_management.py's _delete_raw_file) - the raw upload
    itself only ever lives in st.session_state (never persisted to disk on
    its own), so what's moved to the bin here is a snapshot of exactly the
    DataFrame being removed, plus the exact ledger keys the caller already
    cleared for it (report_key/ledger_kind/keys - the caller snapshots
    those from get_seen_keys()/get_latest_records() BEFORE calling
    clear_seen_keys()/clear_latest_records(), since this function doesn't
    touch the ledger itself, matching soft_delete_run()'s split of
    responsibilities above).

    session_key/dict_entry: where this data lives in session_state - see
    views/page_data_management.py's _dtc_raw_file_specs/
    _amazon_raw_file_specs (dict_entry=None for a scalar slot like
    orders_df/bank_df). Recorded here so restore_raw_file_from_recycle_bin
    below knows where to hand the restored DataFrame back to.

    Returns the new Recycle Bin item_id.
    """
    item_id = _new_recycle_bin_item(
        client_key, "raw_file", label,
        extra_meta={"session_key": session_key, "dict_entry": dict_entry},
        payload={"session_key": session_key, "dict_entry": dict_entry, "data": data},
        ledger_removals=ledger_removals, retention_days=retention_days,
    )
    return item_id


def list_recycle_bin(client_key):
    """Every item currently in the bin, newest-deleted first, with how
    many days remain before it's purged automatically - what the "Recycle
    Bin" section of Data Management lists (label + kind + deletion date,
    per the client's explicit "ideally... show the deleted files/data and
    their deletion date" ask, plus the countdown so the retention window
    is visible too, not just implied)."""
    folder = _recycle_bin_dir(client_key)
    now = dt.datetime.now()
    items = []
    for item_id in sorted(os.listdir(folder), reverse=True):
        manifest_path = os.path.join(folder, item_id, "manifest.json")
        if not os.path.exists(manifest_path):
            continue
        try:
            with open(manifest_path) as f:
                manifest = json.load(f)
        except Exception:
            continue
        expires_at = dt.datetime.fromisoformat(manifest["expires_at"]) if manifest.get("expires_at") else None
        days_remaining = max(0, (expires_at - now).days) if expires_at else None
        items.append({
            "item_id": manifest["item_id"],
            "kind": manifest.get("kind"),
            "label": manifest.get("label"),
            "month_label": manifest.get("month_label"),
            "order_count": manifest.get("order_count"),
            "channel_name": manifest.get("channel_name"),
            "deleted_at": manifest.get("deleted_at"),
            "expires_at": manifest.get("expires_at"),
            "days_remaining": days_remaining,
        })
    return items


def restore_run_from_recycle_bin(client_key, item_id):
    """
    Puts a soft-deleted saved period (kind="dtc"/"amazon") back exactly as
    it was: rewrites its .pkl + .meta.json sidecar under the normal saved-
    runs folder (so it reappears in list_runs()/list_amazon_runs()
    immediately), replays every ledger removal the original delete made
    (add_seen_keys - a plain union, safe even if some of those keys were
    somehow already re-added by a fresh upload in the meantime), then
    removes the Recycle Bin item.

    Returns the restored payload's fname, or None if item_id doesn't exist
    or isn't a saved-period item (raw_file items go through
    restore_raw_file_from_recycle_bin instead).
    """
    item_dir = _recycle_bin_item_dir(client_key, item_id)
    manifest_path = os.path.join(item_dir, "manifest.json")
    if not os.path.exists(manifest_path):
        return None
    with open(manifest_path) as f:
        manifest = json.load(f)
    if manifest.get("kind") not in ("dtc", "amazon"):
        return None

    with open(os.path.join(item_dir, "payload.pkl"), "rb") as f:
        payload = pickle.load(f)

    fname = manifest["fname"]
    path = os.path.join(_client_dir(client_key), fname)
    with open(path, "wb") as f:
        pickle.dump(payload, f)
    _write_meta_sidecar(path, _meta_from_payload(payload, manifest["kind"]))

    for removal in manifest.get("ledger_removals", []):
        if removal.get("ledger_kind") == "seen":
            add_seen_keys(client_key, removal["report_key"], removal["keys"])
        else:
            record_latest_records(client_key, removal["report_key"], {k: True for k in removal["keys"]})

    shutil.rmtree(item_dir, ignore_errors=True)
    return fname


def restore_raw_file_from_recycle_bin(client_key, item_id):
    """
    restore_run_from_recycle_bin()'s counterpart for a soft-deleted
    individual raw file (kind="raw_file") - replays its ledger removals
    the same way, removes the Recycle Bin item, and hands the restored
    DataFrame plus its original session_key/dict_entry back to the caller
    (views/page_data_management.py), since only the caller can reach
    st.session_state to actually put it back into the live upload state.

    Returns {"session_key", "dict_entry", "data"}, or None if item_id
    doesn't exist or isn't a raw_file item.
    """
    item_dir = _recycle_bin_item_dir(client_key, item_id)
    manifest_path = os.path.join(item_dir, "manifest.json")
    if not os.path.exists(manifest_path):
        return None
    with open(manifest_path) as f:
        manifest = json.load(f)
    if manifest.get("kind") != "raw_file":
        return None

    with open(os.path.join(item_dir, "payload.pkl"), "rb") as f:
        payload = pickle.load(f)

    for removal in manifest.get("ledger_removals", []):
        if removal.get("ledger_kind") == "seen":
            add_seen_keys(client_key, removal["report_key"], removal["keys"])
        else:
            record_latest_records(client_key, removal["report_key"], {k: True for k in removal["keys"]})

    shutil.rmtree(item_dir, ignore_errors=True)
    return payload


def permanently_delete_recycle_bin_item(client_key, item_id):
    """Empties one item from the bin right now, on the user's own explicit
    request (not the automatic 30-day purge below) - the ledger removals
    already applied at soft-delete time simply stay applied; there's
    nothing left to "finish" deleting other than the bin copy itself."""
    item_dir = _recycle_bin_item_dir(client_key, item_id)
    existed = os.path.isdir(item_dir)
    shutil.rmtree(item_dir, ignore_errors=True)
    return existed


def purge_expired_recycle_bin_items(client_key, retention_days=RECYCLE_BIN_RETENTION_DAYS):
    """
    Automatically empties every bin item whose retention window has
    passed - the client's explicit "after the retention period, the data
    should be permanently deleted automatically" requirement. Since this
    app has no background scheduler, this is called as a lazy sweep every
    time the Data Management page renders (views/page_data_management.py)
    - cheap even with many items, since it only ever reads each item's
    small manifest.json, never its (potentially large) payload.pkl.

    retention_days is accepted mainly for tests/overrides - normal calls
    rely on each item's own already-computed expires_at (set from
    RECYCLE_BIN_RETENTION_DAYS at the moment it was soft-deleted), so
    changing the module default later doesn't retroactively change when
    already-binned items expire.

    Returns the number of items purged.
    """
    folder = _recycle_bin_dir(client_key)
    now = dt.datetime.now()
    purged = 0
    for item_id in os.listdir(folder):
        manifest_path = os.path.join(folder, item_id, "manifest.json")
        if not os.path.exists(manifest_path):
            continue
        try:
            with open(manifest_path) as f:
                manifest = json.load(f)
            expires_at = dt.datetime.fromisoformat(manifest["expires_at"])
        except Exception:
            continue
        if now >= expires_at:
            shutil.rmtree(os.path.join(folder, item_id), ignore_errors=True)
            purged += 1
    return purged


# ---------------------------------------------------------------------------
# Persisted raw uploads
# ---------------------------------------------------------------------------
# Client-reported (2026-08-22): after re-pushing an unrelated engine fix -
# which meant restarting the Streamlit app to pick it up - the client's
# in-progress Razorpay upload attempt immediately failed with "the Shopify
# order report needs to be uploaded (and confirmed above) first", and Data
# Management's "Uploaded raw files" delete list came up completely empty
# ("I do not see an option to delete an individual report"). Both are the
# SAME root cause, not two separate bugs: every individually-uploaded-and-
# CONFIRMED raw file (Shopify orders, Razorpay, Gokwik, Shiprocket, ...)
# lived ONLY in Streamlit's st.session_state, which is wiped clean by any
# app restart or session expiry - unlike a "saved month/period"
# (save_run/save_amazon_run above), which was always written to disk, raw
# uploads had no such durability. The client had genuinely uploaded and
# confirmed the Shopify order report earlier that session, so from their
# perspective nothing was wrong with their own actions - the app had
# simply forgotten everything they'd fed it, through no fault of theirs.
#
# This section persists every confirmed raw upload here, one small pickle
# per file slot plus a manifest (same "manifest + payload" shape as the
# Recycle Bin above, for the same reason - crash-safe, human-inspectable,
# no format migration needed later), and views/page_upload.py calls
# save_raw_upload() right alongside every st.session_state[...] = ...
# assignment for a confirmed upload. views/state_init.py's init_state()
# and views/page_upload.py's reset_upload_state() call load_raw_uploads()
# to restore them into session_state - once at the very start of a fresh
# session (for whichever client/channel ends up active by default), and
# again any time the user explicitly switches platform/client account
# (from either the Upload Data page or the Settings page - both now go
# through reset_upload_state(), see its own docstring), so a raw file
# behaves the same whether the app has been running continuously or was
# just restarted, exactly like a saved period always has. Deleting a raw
# file (views/page_data_management.py's _delete_raw_file) calls
# clear_raw_upload() so it doesn't reappear on the next restart - it's
# already been moved to the Recycle Bin at that point, same as before this
# section existed; this only removes the SEPARATE always-current-copy this
# section maintains for restart durability.
#
# RAW_UPLOAD_SLOTS below is the single canonical list of which
# session_state keys are individually-deletable raw upload slots (see
# views/page_data_management.py) and how each one defaults when nothing's
# uploaded yet ("dict" -> {}, "scalar" -> None - a scalar slot only ever
# holds one file, e.g. orders_df/bank_df; a dict slot holds one file per
# label/segment/payment-mode, e.g. delivery_frames/gateway_frames/
# attribution_frames/mtr_files/settlement_files). Every caller that needs
# to reset and/or restore raw-upload session state (views/state_init.py,
# views/page_upload.py, views/page_settings.py) imports this rather than
# keeping its own hand-maintained copy of the key list.
#
# That hand-maintained-copy pattern is exactly what caused this class of
# bug TWICE before this constant existed: (2026-08-22) the very first
# version of this persistence layer wasn't wired into every reset path,
# and (2026-08-25) adding attribution_frames as a new raw-upload type
# updated views/page_upload.py's own copy of the key list but not
# views/state_init.py's separate copy - and a THIRD copy in
# views/page_settings.py's channel switcher was never wired up at all, so
# switching channel from Settings silently dropped every raw upload's
# on-disk copy out of session_state without restoring it. All three times
# the client-visible symptom was identical: a raw report that had genuinely
# already been uploaded and confirmed looked exactly like it had never
# been uploaded at all, and Data Management's per-file Delete option for
# it disappeared. Routing every caller through this one constant (and
# through reset_upload_state() itself where possible, see
# views/page_settings.py) means a future new raw-upload type only needs to
# be added here once.
RAW_UPLOAD_SLOTS = {
    "orders_df": "scalar",
    "bank_df": "scalar",
    "delivery_frames": "dict",
    "gateway_frames": "dict",
    "attribution_frames": "dict",
    "mtr_files": "dict",
    "settlement_files": "dict",
}
RAW_UPLOAD_SESSION_KEYS = tuple(RAW_UPLOAD_SLOTS.keys())
RAW_UPLOAD_DICT_KEYS = tuple(k for k, shape in RAW_UPLOAD_SLOTS.items() if shape == "dict")


def raw_upload_defaults():
    """A fresh {key: default} dict covering every raw-upload slot - {} for
    a dict-shaped slot, None for a scalar one. Built fresh on every call
    (never shared/mutated) so one caller's dict default can never leak
    into another caller's session_state by reference."""
    return {key: ({} if shape == "dict" else None) for key, shape in RAW_UPLOAD_SLOTS.items()}


def _raw_uploads_dir(client_key):
    d = os.path.join(_client_dir(client_key), "_raw_uploads")
    os.makedirs(d, exist_ok=True)
    return d


def _raw_uploads_manifest_path(client_key):
    return os.path.join(_raw_uploads_dir(client_key), "manifest.json")


def _read_raw_uploads_manifest(client_key):
    path = _raw_uploads_manifest_path(client_key)
    if not os.path.exists(path):
        return []
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return []


def _write_raw_uploads_manifest(client_key, entries):
    with open(_raw_uploads_manifest_path(client_key), "w") as f:
        json.dump(entries, f)


def save_raw_upload(client_key, session_key, dict_entry, df):
    """
    Persists ONE currently-uploaded-and-confirmed raw file slot to disk -
    session_key matches the session_state key it lives under (any key in
    RAW_UPLOAD_SLOTS above); dict_entry is None for a scalar slot
    (orders_df/bank_df - there's only ever one file) or the dict key
    within that slot for a multi-file one (the delivery/gateway label, MTR
    segment, or settlement payment_mode) - the exact same (session_key,
    dict_entry) shape views/page_data_management.py's raw-file-delete
    specs already use, so a delete and this persistence layer always agree
    on what identifies "one raw file".

    A no-op if df is None - nothing to persist, and callers (see
    views/page_upload.py) only ever call this right after a real,
    confirmed upload has landed in session_state.
    """
    if df is None:
        return
    entries = _read_raw_uploads_manifest(client_key)
    entries = [
        e for e in entries
        if not (e["session_key"] == session_key and e.get("dict_entry") == dict_entry)
    ]
    fname = f"{uuid.uuid4().hex}.pkl"
    with open(os.path.join(_raw_uploads_dir(client_key), fname), "wb") as f:
        pickle.dump(df, f)
    entries.append({
        "session_key": session_key,
        "dict_entry": dict_entry,
        "file": fname,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
    })
    _write_raw_uploads_manifest(client_key, entries)


def clear_raw_upload(client_key, session_key, dict_entry):
    """Removes ONE raw file slot's persisted copy - called when that slot
    is individually deleted (views/page_data_management.py's
    _delete_raw_file), so it doesn't silently reappear the next time the
    app restarts or the session is refreshed, even though a snapshot of it
    still separately lives in the Recycle Bin (restorable there, same as
    always). Returns True if something was actually removed."""
    entries = _read_raw_uploads_manifest(client_key)
    keep, removed = [], None
    for e in entries:
        if e["session_key"] == session_key and e.get("dict_entry") == dict_entry:
            removed = e
        else:
            keep.append(e)
    if removed is None:
        return False
    try:
        os.remove(os.path.join(_raw_uploads_dir(client_key), removed["file"]))
    except OSError:
        pass
    _write_raw_uploads_manifest(client_key, keep)
    return True


def load_raw_uploads(client_key):
    """
    Rebuilds every currently-persisted raw file slot for this client from
    disk, in the exact shape session_state needs - one key per entry in
    RAW_UPLOAD_SLOTS above, e.g.:
        {"orders_df": df, "bank_df": df,
         "delivery_frames": {label: df, ...}, "gateway_frames": {label: df, ...},
         "attribution_frames": {label: df, ...},
         "mtr_files": {segment: df, ...}, "settlement_files": {payment_mode: df, ...}}
    Only the slots that actually have a persisted file appear as keys -
    callers (views/state_init.py, views/page_upload.py) only overwrite
    session_state for keys present in this result, leaving anything else
    (e.g. a slot nothing has ever been uploaded for) at its existing
    default. A file that fails to unpickle (corrupted, truncated write)
    is silently skipped rather than blowing up every page load - the
    client can simply re-upload that one report, same as if it had never
    been uploaded at all.
    """
    result = {}
    for e in _read_raw_uploads_manifest(client_key):
        path = os.path.join(_raw_uploads_dir(client_key), e.get("file", ""))
        try:
            with open(path, "rb") as f:
                df = pickle.load(f)
        except Exception:
            continue
        if e.get("dict_entry") is None:
            result[e["session_key"]] = df
        else:
            result.setdefault(e["session_key"], {})[e["dict_entry"]] = df
    return result
