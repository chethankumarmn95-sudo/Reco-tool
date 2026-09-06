"""
theme.py
--------
Shared visual styling and small session-state helpers used by every page.
Keeping this in one place means the whole app (sidebar, cards, charts) has
one consistent look instead of each page inventing its own style.
"""

import streamlit as st

BRAND_PURPLE = "#6C5CE7"
BRAND_DARK = "#1A1A2E"
BRAND_BG = "#F7F7FB"

CUSTOM_CSS = f"""
<style>
/* Sidebar */
section[data-testid="stSidebar"] {{
    background-color: {BRAND_DARK};
}}
section[data-testid="stSidebar"] * {{
    color: #E8E8F0 !important;
}}
section[data-testid="stSidebar"] [data-testid="stNavSectionHeader"] {{
    color: #8888A0 !important;
}}

/* "‹ Back to Portal" / "Log out" buttons - Streamlit renders these with
its own light button background regardless of the dark sidebar around
them, so the blanket "make all sidebar text near-white" rule just above
was leaving pale text on a pale button (nearly invisible). Give buttons
their own explicit, readable styling instead of only inheriting the
generic sidebar text color. */
section[data-testid="stSidebar"] button {{
    background-color: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(232,232,240,0.35) !important;
}}
section[data-testid="stSidebar"] button:hover {{
    background-color: rgba(255,255,255,0.16) !important;
    border-color: #FFFFFF !important;
}}
section[data-testid="stSidebar"] button p,
section[data-testid="stSidebar"] button span,
section[data-testid="stSidebar"] button div {{
    color: #FFFFFF !important;
}}

/* KPI card look for st.metric */
div[data-testid="stMetric"] {{
    background: white;
    border: 1px solid #ECECF4;
    border-radius: 12px;
    padding: 16px 18px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    /* Let the card grow to fit its value instead of clipping it - see
    stMetricValue rule below for why this pairing is needed. */
    min-width: 0;
    overflow: visible;
}}
div[data-testid="stMetricLabel"] {{
    color: #6B6B80;
}}

/* Full-value KPI numbers - the fix for "₹50,92,570 shows as ₹50,92,...".
Streamlit's own stMetricValue CSS is a single-line box with
`white-space: nowrap` + `text-overflow: ellipsis`, sized to whatever
column width st.columns(N) hands it - so a wide rupee figure (lakhs/
crores, comma-grouped) truncates with "..." the moment N columns leave
less room than the text needs, regardless of screen size. Fixed by
letting the value wrap onto a second line and trimming its font size
slightly, rather than relying on a fixed column count to always be wide
enough - so the complete amount is always readable, on any screen size
and no matter how many KPI cards share a row. */
div[data-testid="stMetricValue"] {{
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word;
    font-size: 1.6rem;
    line-height: 1.25;
}}
div[data-testid="stMetricValue"] > div {{
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}}

/* Section headers */
h1, h2, h3 {{
    color: {BRAND_DARK};
}}

/* Health score badge */
.health-score-card {{
    background: linear-gradient(135deg, {BRAND_PURPLE}, #8E7CF3);
    border-radius: 14px;
    padding: 20px;
    color: white;
    text-align: center;
}}
</style>
"""


def apply_theme():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def has_reconciliation():
    return st.session_state.get("reco_df") is not None


def get_active_config():
    """The currently selected client/channel config, or None if not set yet."""
    return st.session_state.get("config")


def _autoload_from_storage():
    """
    If nothing is loaded in this session yet, but saved months exist on
    disk, load the current Financial Year's combined data automatically -
    same behavior as the Dashboard, so every page is consistent regardless
    of whether this is a fresh session or one that just ran a reconciliation.

    Never falls back to blending EVERY saved financial year together (that
    used to happen here when the current FY had no saved data yet - e.g.
    it's FY 2026-27 already but only FY 2025-26 has been saved). Per the
    client's own rule (see views/filters.py's module docstring), financial
    years must never mix - so the fallback here is the single MOST RECENT
    financial year that actually has saved data, same one financial year
    the Dashboard itself would default to, never a blend of all of them.
    """
    if has_reconciliation():
        return
    client_key = st.session_state.get("client_key")
    if not client_key:
        return

    from engine import storage
    from views.filters import current_financial_year, available_financial_years

    runs = storage.list_runs(client_key)
    if not runs:
        return

    fy = current_financial_year()
    runs_in_fy = [r for r in runs if r.get("financial_year") == fy]
    if not runs_in_fy:
        fys = available_financial_years(runs)  # newest first
        runs_in_fy = [r for r in runs if r.get("financial_year") == fys[0]] if fys else runs
    fnames = [r["file"] for r in runs_in_fy]
    combined_reco, combined_lookup = storage.combine_runs(client_key, fnames)

    from engine.summary import headline_totals
    st.session_state["reco_df"] = combined_reco
    st.session_state["lookup_df"] = combined_lookup
    st.session_state["totals"] = headline_totals(combined_reco)


def require_data_or_prompt(target_page_label="Upload Data"):
    """Standard 'no data yet' message used by pages that need a completed
    reconciliation to show anything. Returns True if data is available.
    Auto-loads saved data from storage first, so this works consistently
    whether or not a reconciliation was just run in this session."""
    _autoload_from_storage()
    if has_reconciliation():
        return True
    st.info(f"No reconciliation loaded yet. Go to **{target_page_label}** to upload files and run one, "
            f"or open **Data Management** to load a previously saved month.")
    return False
