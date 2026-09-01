"""
period.py
---------
Classifies which reconciliation "period" each order belongs to, so a bank
credit that bundles settlements from more than one period - a common
e-commerce reality when a delayed COD/gateway payout lands together with
the current period's payouts - can be split out correctly instead of
silently inflating or understating the current period's numbers.

This mirrors what was previously done by hand in the client's sample
"Bank_reco" workbook: a "Reco period" Yes/No column, with "No" rows
further explained by an "Order month" column showing which other month
the order actually belongs to (e.g. "March sale receipt").

How "this period" is decided (confirmed with the client): whatever
order(s) are present in the Shopify order file uploaded for the CURRENT
reconciliation run define "this period" - auto-detected, no manual date
range to fill in every time. Orders that show up in a gateway/bank
receipt but are NOT part of the current run's order file are looked up in
previously saved months (engine/storage.py) to tell "Previous period"
apart from "Subsequent period" by comparing order dates - and if an order
can't be found anywhere at all, it's flagged "Order not found" rather
than silently guessed at.
"""

import pandas as pd

from . import storage

THIS_PERIOD = "This period"
PREVIOUS_PERIOD = "Previous period"
SUBSEQUENT_PERIOD = "Subsequent period"
NOT_FOUND = "Order not found"


def _historical_order_dates(client_key):
    """
    Builds a single order_id -> earliest known order date lookup across
    every previously saved month for this client, so an order that
    belongs to a different period than the one currently being run can
    still be dated (and therefore classified Previous vs Subsequent)
    instead of just being labelled "unknown".

    Returns an empty dict if nothing has ever been saved yet (e.g. the
    very first month run for a brand-new client) - callers should treat
    that as "can't tell, flag as not found" rather than an error.
    """
    if not client_key:
        return {}

    lookup = {}
    for run_meta in storage.list_runs(client_key):
        try:
            payload = storage.load_run(client_key, run_meta["file"])
        except Exception:
            continue
        reco_df = payload.get("reco_df")
        if reco_df is None or "order_id" not in reco_df.columns or "created_at" not in reco_df.columns:
            continue
        dates = pd.to_datetime(reco_df["created_at"], errors="coerce")
        for oid, d in zip(reco_df["order_id"].astype(str), dates):
            if pd.isna(d):
                continue
            if oid not in lookup or d < lookup[oid]:
                lookup[oid] = d
    return lookup


def classify_order_periods(order_ids, current_order_master, client_key=None):
    """
    order_ids: an iterable/Series of order_id values to classify (e.g.
               every distinct order_id that appears in the consolidated
               gateway receipts - which may include orders NOT in this
               run's order file at all, like a delayed settlement for an
               order sold last period).
    current_order_master: this run's own order-level table - needs
               "order_id" and "created_at" columns (engine/reco.py's Reco
               working table, or its build_order_master() output). Every
               order_id in here is, by definition, "This period".
    client_key: which client/channel's saved history to search for orders
               that aren't part of the current run - pass None to skip
               history lookup entirely (everything not in the current run
               is then "Order not found").

    Returns a pandas Series (same length/order as order_ids) of one of:
      "This period" / "Previous period" / "Subsequent period" / "Order not found"
    """
    order_id_series = pd.Series(list(order_ids)).astype(str)

    if current_order_master is not None and len(current_order_master):
        current_ids = set(current_order_master["order_id"].astype(str))
        current_dates = pd.to_datetime(current_order_master["created_at"], errors="coerce")
    else:
        current_ids = set()
        current_dates = pd.Series(dtype="datetime64[ns]")

    valid_dates = current_dates.dropna()
    period_start = valid_dates.min() if len(valid_dates) else None
    period_end = valid_dates.max() if len(valid_dates) else None

    historical_dates = _historical_order_dates(client_key)

    def classify_one(oid):
        if oid in current_ids:
            return THIS_PERIOD
        hist_date = historical_dates.get(oid)
        if hist_date is None or period_start is None or period_end is None:
            return NOT_FOUND
        if hist_date < period_start:
            return PREVIOUS_PERIOD
        if hist_date > period_end:
            return SUBSEQUENT_PERIOD
        # Falls inside the current window's date span but wasn't part of
        # this run's own order file (e.g. a combined-months history
        # overlap) - treat as belonging to this period rather than
        # manufacturing an unexplained 4th bucket.
        return THIS_PERIOD

    return order_id_series.apply(classify_one)
