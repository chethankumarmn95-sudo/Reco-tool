"""
filters.py
----------
Shared filtering logic for saved reconciliation runs - used by Dashboard
(cumulative FY view), Reports (month/date-wise downloads), and Exceptions
(filtered order-level export). One place for "what counts as FY 2026-27",
"what date range did the user pick", etc. so all three pages agree.

Sales Channel is a plain dropdown (not a multiselect) everywhere it
appears. Every saved run for a given client/channel context already shares
one channel, so a removable "chip" was never really filtering anything -
it was just one accidental click away from an empty selection. A dropdown
can't be fat-fingered the same way, and "All" is always there to see every
channel again.

Financial Year, by contrast, is deliberately NOT given an "All" option
(client's own explicit requirement, 2026-08-20: "data from different
financial years should not be mixed into a single report... the user
should be able to clearly select/view the required financial year").
Letting someone pick "All" here would silently sum/concatenate two
financial years' figures into one Dashboard view or one downloaded
report - exactly the bug that was reported. The dropdown always has one
financial year selected (defaulting to the current FY, or the most recent
saved one if the current FY has no saved data yet), and the user switches
it explicitly to look at a different year. This is enforced independently
of - and in addition to - the fix in views/page_reconciliation.py that
splits a single reconciliation run into one saved period PER financial
year before it's even saved, so the two fixes together mean a financial
year can never end up blended with another one anywhere in the app.
"""

import datetime as dt
import pandas as pd
import streamlit as st

from engine import storage

ALL_CHANNELS = "All channels"


def current_financial_year():
    today = dt.date.today()
    year = today.year if today.month >= 4 else today.year - 1
    return f"FY {year}-{str(year + 1)[-2:]}"


def available_financial_years(runs):
    fys = sorted({r["financial_year"] for r in runs if r.get("financial_year")}, reverse=True)
    return fys


def _channel_dropdown(channels, key):
    if not channels:
        return None
    options = [ALL_CHANNELS] + channels
    choice = st.selectbox("Sales Channel", options, index=0, key=key)
    return None if choice == ALL_CHANNELS else choice


def _apply_date_range(runs, date_from, date_to):
    """Keeps only the saved runs (whole months) that could contain data in
    the picked range. This is coarse - a month-level check - the exact
    row-level trim to date_from/date_to happens afterward in
    trim_to_date_range() once the matching runs have been loaded."""
    if not (date_from or date_to):
        return runs

    def in_range(r):
        dmin = pd.to_datetime(r.get("date_min"), errors="coerce")
        dmax = pd.to_datetime(r.get("date_max"), errors="coerce")
        if pd.isna(dmin) or pd.isna(dmax):
            return True  # can't check - don't exclude
        if date_from and dmax.date() < date_from:
            return False
        if date_to and dmin.date() > date_to:
            return False
        return True

    return [r for r in runs if in_range(r)]


def render_dashboard_filter_controls(runs, key_prefix="dash"):
    """
    A deliberately simpler filter set for the Dashboard than Reports/
    Exceptions use: Financial Year + Sales Channel, with a collapsed
    custom date range for when someone needs to zoom into a specific
    period. No Month multiselect - the Dashboard is meant to be glanced
    at, not configured.

    The date range applies live as soon as both fields are picked -
    Streamlit reruns the page automatically on every widget change, so
    there's no separate "Run" button to press.
    """
    fys = available_financial_years(runs)
    channels = sorted({r.get("channel_name") for r in runs if r.get("channel_name")})

    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        # No "All" option here on purpose - see this module's docstring.
        # `runs` is guaranteed non-empty by every caller (they check
        # `if not runs:` first), so `fys` is never empty either.
        fy_options = fys
        default_fy = current_financial_year() if current_financial_year() in fys else fys[0]
        fy_choice = st.selectbox("Financial Year", fy_options,
                                  index=fy_options.index(default_fy),
                                  key=f"{key_prefix}_fy")
    with c2:
        channel_choice = _channel_dropdown(channels, key=f"{key_prefix}_channel")
    with c3:
        with st.popover("📅 Custom date range"):
            date_from = st.date_input("From date", value=None, key=f"{key_prefix}_from")
            date_to = st.date_input("To date", value=None, key=f"{key_prefix}_to")

    filtered = [r for r in runs if r.get("financial_year") == fy_choice]
    if channel_choice:
        filtered = [r for r in filtered if r.get("channel_name") == channel_choice]

    filtered = _apply_date_range(filtered, date_from, date_to)

    return filtered, (date_from, date_to)


def render_filter_controls(runs, key_prefix, show_date_range=True, show_month=True):
    """
    Renders FY / [Month] / [Date range] / Channel filter widgets and
    returns the filtered list of run metadata dicts. Shared UI so
    Dashboard, Reports, and Exceptions all filter the same way.

    show_month=False drops the Month multiselect entirely (Reports uses
    this - Financial Year + Date range already narrows down to an exact
    scope, so Month was a second, redundant way to do the same job and
    just as removable-by-accident as the old Sales Channel chips were).
    When it's hidden, every month in the selected FY is included by
    default and the date range does the actual narrowing.

    Exceptions still shows Month (show_month defaults True) - useful there
    as a quick way to isolate one saved month's exceptions without also
    picking exact dates. Sales Channel is a dropdown, same reasoning as
    the Dashboard version above.
    """
    fys = available_financial_years(runs)
    channels = sorted({r.get("channel_name") for r in runs if r.get("channel_name")})

    n_cols = 1 + (1 if show_month else 0) + (2 if show_date_range else 0)
    cols = iter(st.columns(n_cols))

    with next(cols):
        # No "All" option here either - same reasoning as
        # render_dashboard_filter_controls above (see module docstring).
        fy_options = fys
        default_fy = current_financial_year() if current_financial_year() in fys else fys[0]
        fy_choice = st.selectbox("Financial Year", fy_options,
                                  index=fy_options.index(default_fy),
                                  key=f"{key_prefix}_fy")

    runs_in_fy = [r for r in runs if r.get("financial_year") == fy_choice]
    month_options = sorted({r["month_label"] for r in runs_in_fy})

    if show_month:
        with next(cols):
            month_choice = st.multiselect("Month", month_options, default=month_options, key=f"{key_prefix}_month")
    else:
        month_choice = month_options  # no picker - date range below does the narrowing instead

    if show_date_range:
        with next(cols):
            date_from = st.date_input("From date", value=None, key=f"{key_prefix}_from")
        with next(cols):
            date_to = st.date_input("To date", value=None, key=f"{key_prefix}_to")
    else:
        date_from = date_to = None

    channel_choice = _channel_dropdown(channels, key=f"{key_prefix}_channel")

    filtered = [r for r in runs_in_fy if r["month_label"] in month_choice]
    if channel_choice:
        filtered = [r for r in filtered if r.get("channel_name") == channel_choice]

    filtered = _apply_date_range(filtered, date_from, date_to)

    return filtered, (date_from, date_to)


def load_combined(client_key, filtered_runs):
    """Loads and concatenates the filtered runs' data, then trims to the
    exact date range if one was specified (a saved month can span more
    than the requested range at the edges)."""
    if not filtered_runs:
        return None, None
    fnames = [r["file"] for r in filtered_runs]
    return storage.combine_runs(client_key, fnames)


def load_combined_with_settlement(client_key, filtered_runs):
    """Like load_combined(), but also combines each saved month's
    consolidated receipt ledger and bank statement ledger - needed for the
    Reports page to recompute the Payment Gateway Settlement Report and
    the UTR-level bank reconciliation over whatever combined, multi-month
    date range is currently selected (see engine.settlement / engine.bank),
    the same way reco_df/lookup_df are already combined here."""
    if not filtered_runs:
        return None, None, None, None
    fnames = [r["file"] for r in filtered_runs]
    return storage.combine_runs(client_key, fnames, include_settlement=True)


def trim_to_date_range(df, date_col, date_from, date_to):
    if df is None or (not date_from and not date_to):
        return df
    if date_col not in df.columns:
        return df
    dates = pd.to_datetime(df[date_col], errors="coerce")
    mask = pd.Series(True, index=df.index)
    if date_from:
        mask &= dates.dt.date >= date_from
    if date_to:
        mask &= dates.dt.date <= date_to
    return df[mask]
