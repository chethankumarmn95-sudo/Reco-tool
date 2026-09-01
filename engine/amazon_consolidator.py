"""
amazon_consolidator.py
------------------------
Turns the raw, combined Settlement Flat File (engine.amazon_loaders.
load_settlement_files output) into two things:

  1. settlement_summary_df - one row per settlement-id: the settlement's
     own start/end/deposit date and Amazon-reported total-amount. This
     comes from the one "header" row Amazon includes at the top of every
     settlement in the flat file (transaction-type is blank, total-amount
     is populated) - used later for settlement-to-bank matching
     (engine/amazon_bank.py), the same way a gateway settlement UTR is used
     in the Shopify pipeline.

  2. expense_ledger_df - the GRANULAR, line-by-line breakup of every fee,
     tax, promotion, reserve movement, and revenue component in the flat
     file. This is deliberately NOT pre-aggregated into a handful of fixed
     buckets (the client explicitly asked that this reconciliation NOT
     repeat last year's approach of merging several Amazon expense types
     into one column) - every distinct fee Amazon names (e.g. "Fixed
     closing fee", "Fixed closing fee CGST", "Fixed closing fee SGST" as
     three separate lines, not netted into one "Closing Fee") is kept as
     its own row here, tagged with exactly which source column it came
     from - so any number can always be traced back to the original flat
     file. Callers that want a wide, one-column-per-category view (e.g. for
     display) can pivot this long table; see pivot_expense_ledger() below.

Why this needs its own "melt" logic (not just resolve_col + sum):
Amazon's flat file is a transaction LEDGER, not a wide table - each row
carries at most a handful of populated "-type"/"-amount" column pairs, and
which pair is populated depends on the transaction-type of that row (an
"Order" row uses price-type/price-amount and item-related-fee-type/
item-related-fee-amount; a "ServiceFee" advertising row uses
item-related-fee-type/item-related-fee-amount for the base cost AND
other-fee-reason-description/other-fee-amount for the GST on it, in the
SAME row; a "Current Reserve Amount" row has no type column at all and
just uses other-amount). Rather than hardcoding a branch per transaction-
type (which breaks the moment Amazon adds a new fee type), this scans every
configured (amount_col, type_col) pair on every row and emits one ledger
line per pair that actually has a non-null amount - verified against every
transaction-type seen in the client's real Flat File COD/Online exports.
"""

import numpy as np
import pandas as pd

from .loaders import resolve_col, normalize_order_id


def build_settlement_summary(raw_settlement_df, settlement_cols_cfg):
    """
    One row per settlement-id (across COD + Online combined):
        settlement_id | payment_mode | start_date | end_date | deposit_date | total_amount
    """
    cols = ["settlement_id", "payment_mode", "start_date", "end_date", "deposit_date", "total_amount"]
    if raw_settlement_df is None or raw_settlement_df.empty:
        return pd.DataFrame(columns=cols)

    df = raw_settlement_df
    total_col = resolve_col(df, settlement_cols_cfg["total_amount_col"])
    sid_col = "_settlement_id_resolved"
    tt_col = resolve_col(df, settlement_cols_cfg["transaction_type_col"])

    if total_col is None or sid_col not in df.columns:
        return pd.DataFrame(columns=cols)

    is_summary_row = df[total_col].notna()
    if tt_col is not None:
        is_summary_row = is_summary_row & df[tt_col].isna()
    summary = df[is_summary_row].copy()
    if summary.empty:
        return pd.DataFrame(columns=cols)

    start_col = resolve_col(df, settlement_cols_cfg.get("settlement_start_col"))
    end_col = resolve_col(df, settlement_cols_cfg.get("settlement_end_col"))
    deposit_col = resolve_col(df, settlement_cols_cfg.get("deposit_date_col"))

    out = pd.DataFrame()
    out["settlement_id"] = summary[sid_col].astype(str).str.strip()
    out["payment_mode"] = summary["_payment_mode"]
    out["start_date"] = pd.to_datetime(summary[start_col], errors="coerce", utc=True).dt.tz_localize(None) if start_col else pd.NaT
    out["end_date"] = pd.to_datetime(summary[end_col], errors="coerce", utc=True).dt.tz_localize(None) if end_col else pd.NaT
    out["deposit_date"] = pd.to_datetime(summary[deposit_col], errors="coerce", utc=True).dt.tz_localize(None) if deposit_col else pd.NaT
    out["total_amount"] = pd.to_numeric(summary[total_col], errors="coerce").fillna(0.0)

    return out.drop_duplicates(subset=["settlement_id"]).reset_index(drop=True)[cols]


# transaction-types whose amount represents a reserve hold/release, not a
# real cost - net to ~0 across a full year, called out separately so the
# reconciliation waterfall doesn't miscount them as a genuine deduction.
_RESERVE_TRANSACTION_TYPES = {"current reserve amount", "previous reserve amount balance"}


def _bucket_for(transaction_type, type_col_used, label, revenue_price_types, tax_price_types):
    """
    Classifies a line into one of four MUTUALLY EXCLUSIVE money "kinds",
    independent of whether it's order- or settlement-level (see `level`,
    computed separately in build_expense_ledger() from order-id presence -
    a line's kind and its level are two independent axes, not one combined
    category, so a Tax-Collected line that happens to sit at settlement
    level doesn't need a fifth bucket name of its own):

      "Revenue"                 - Principal/Product Tax/Shipping/Shipping
                                   tax price-type lines. Already fully
                                   captured from MTR's Invoice Amount for
                                   the waterfall's "Sales as per MTR" line -
                                   MUST be excluded from any deduction
                                   total, or revenue gets double-counted.
      "Tax Collected (TDS/TCS)" - TDS 194-O / TCS price-type lines. Real
                                   cash withheld from the seller (matches
                                   the prior year's own Section 1
                                   "Collection Fee" deduction) - INCLUDED
                                   in the deduction total for a cash
                                   reconciliation, even though it's
                                   separately recoverable as an ITR credit
                                   (a bookkeeping distinction, not a cash
                                   one - see engine/amazon_reco.py).
      "Reserve Movement"        - Current/Previous Reserve Amount lines.
                                   Nets to ~0 across a full year - cash-
                                   flow timing, not a real cost. Excluded
                                   from the deduction total, kept as a
                                   memo line in the waterfall.
      "Deduction/Credit"        - everything else: every fee (Commission,
                                   FBA, Closing Fee, Easy Ship, Shipping
                                   Chargeback, Advertising, Storage,
                                   Warehouse/Inbound, Debt adjustments...),
                                   every promotion line, and every
                                   fulfilment-fee-refund credit. INCLUDED
                                   in the deduction total.
    """
    tt = str(transaction_type or "").strip().lower()
    if type_col_used == "price" and label in revenue_price_types:
        return "Revenue"
    if type_col_used == "price" and label in tax_price_types:
        return "Tax Collected (TDS/TCS)"
    if tt in _RESERVE_TRANSACTION_TYPES:
        return "Reserve Movement"
    return "Deduction/Credit"


def _bucket_for_vectorized(tt_series, type_col_used, label_series, revenue_price_types, tax_price_types):
    """
    Same classification as _bucket_for() above (kept as the scalar
    reference implementation/docstring source of truth), but computed for
    an entire (amount_col, type_col) pair's worth of rows at once via
    np.select instead of one Python function call per row - see
    build_expense_ledger()'s docstring for why this matters: a Settlement
    Flat File can run into the lakhs of rows for a full financial year, and
    a Python-level per-row loop over that (previously: `for pos in
    rows_idx: bucket = _bucket_for(...)`) was the single largest cost in
    generating the reconciliation output. Returns a numpy array of bucket
    labels, same length/order as tt_series/label_series.
    """
    tt_lower = tt_series.astype(str).str.strip().str.lower()
    is_reserve = tt_lower.isin(_RESERVE_TRANSACTION_TYPES)

    if type_col_used == "price":
        is_revenue = label_series.isin(revenue_price_types)
        is_tax = label_series.isin(tax_price_types)
    else:
        false_arr = np.zeros(len(tt_series), dtype=bool)
        is_revenue = pd.Series(false_arr, index=tt_series.index)
        is_tax = pd.Series(false_arr, index=tt_series.index)

    return np.select(
        [is_revenue.to_numpy(), is_tax.to_numpy(), is_reserve.to_numpy()],
        ["Revenue", "Tax Collected (TDS/TCS)", "Reserve Movement"],
        default="Deduction/Credit",
    )


def build_expense_ledger(raw_settlement_df, settlement_cols_cfg):
    """
    The granular, line-by-line ledger described in the module docstring.
    Columns:
        settlement_id | payment_mode | order_id | sku | shipment_id |
        posted_date | transaction_type | category | amount | bucket |
        level | source_amount_col | source_type_col | row_ref

    `bucket` (Revenue / Tax Collected (TDS/TCS) / Reserve Movement /
    Deduction-Credit) and `level` (Order / Settlement, from whether the
    flat file row carried an order-id) are independent axes - see
    _bucket_for()'s docstring above for exactly why they're kept separate
    rather than combined into one composite category name.

    row_ref is the row's position in the combined raw settlement frame -
    kept purely so any figure in this ledger can be pointed straight back
    at the exact source row during an audit, without needing to re-search
    the multi-hundred-thousand-row raw file by eye.

    Performance note: this melt scans every configured (amount_col,
    type_col) pair (a fixed, small number - e.g. 9 for Amazon's flat file)
    against the WHOLE settlement frame, which can run into the lakhs of
    rows for a full financial year's COD + Online export. Every step below
    is a vectorised pandas/numpy operation over a full column or boolean-
    masked slice - there is deliberately no Python-level `for row in
    df...`/`.apply()` loop scanning individual transaction lines. An
    earlier version of this function built the ledger with a nested
    `for amount_spec, type_spec in pairs: for pos in rows_idx: ...` loop -
    correct, but for a large real client file that meant several hundred
    thousand individual Python-level `.loc[]` scalar lookups plus one dict
    construction per ledger line, which is what made a full-year Settlement
    Flat File take minutes to reconcile instead of seconds. This function
    produces byte-for-byte the same ledger (same columns, same values, same
    row order - verified by direct comparison against the old row-by-row
    implementation) purely by replacing "loop over rows" with "operate on
    the whole column/slice at once".
    """
    cols = [
        "settlement_id", "payment_mode", "order_id", "sku", "shipment_id", "posted_date",
        "transaction_type", "category", "amount", "bucket", "level", "source_amount_col",
        "source_type_col", "row_ref",
    ]
    if raw_settlement_df is None or raw_settlement_df.empty:
        return pd.DataFrame(columns=cols)

    df = raw_settlement_df
    tt_col = resolve_col(df, settlement_cols_cfg["transaction_type_col"])
    if tt_col is None:
        return pd.DataFrame(columns=cols)

    # Only real transaction rows carry a transaction-type; the one blank-
    # transaction-type row per settlement is the summary row handled by
    # build_settlement_summary() above, not a ledger line.
    tx = df[df[tt_col].notna()].copy()
    if tx.empty:
        return pd.DataFrame(columns=cols)

    sid_col = "_settlement_id_resolved"
    order_col = resolve_col(tx, settlement_cols_cfg.get("order_id_col"))
    merchant_order_col = resolve_col(tx, settlement_cols_cfg.get("merchant_order_id_col"))
    sku_col = resolve_col(tx, settlement_cols_cfg.get("sku_col"))
    shipment_col = resolve_col(tx, settlement_cols_cfg.get("shipment_id_col"))
    posted_col = resolve_col(tx, settlement_cols_cfg.get("posted_date_col"))

    # fillna("") is essential here, not cosmetic: pandas' string dtype lets
    # a NaN survive .astype(str)/.str.replace() (normalize_order_id) as an
    # actual float NaN rather than becoming empty/"nan" text - and
    # Python's bool(nan) is True, so every `if oid` check below would
    # otherwise treat a MISSING order-id as present, mis-classifying every
    # settlement-level row (Advertising, Storage, Reserve movements, Debt
    # adjustments...) as "Order-level" simply because pandas' null wasn't
    # falsy. Caught by this module's own settlement total tie-out test
    # showing zero settlement-level rows, which should never happen.
    #
    # order-id vs merchant-order-id: Amazon's flat file carries BOTH
    # columns, and for the vast majority of rows they're identical - but
    # NOT always. Some transaction rows (certain refund/adjustment/service-
    # fee lines, confirmed against a client report of orders whose expense
    # lines were present in the flat file but invisible in this ledger)
    # leave "order-id" blank while "merchant-order-id" is populated for the
    # exact same order. Using order-id alone silently misclassified those
    # rows as settlement-level (no order attribution at all) even though
    # Amazon's own file clearly ties them to a real order - so every row
    # here falls back to merchant-order-id, per row, whenever order-id
    # itself is blank, rather than only ever reading one of the two
    # columns and ignoring the other entirely.
    if order_col:
        primary = normalize_order_id(tx[order_col])
    else:
        primary = pd.Series(pd.NA, index=tx.index, dtype="object")
    if merchant_order_col:
        fallback = normalize_order_id(tx[merchant_order_col])
        blank = primary.isna() | (primary.astype(str).str.strip() == "")
        primary = primary.where(~blank, fallback)
    order_id_series = primary.fillna("")
    sku_series = tx[sku_col].astype(str).where(tx[sku_col].notna(), None) if sku_col else pd.Series(None, index=tx.index)
    shipment_series = tx[shipment_col] if shipment_col else pd.Series(None, index=tx.index)
    posted_series = pd.to_datetime(tx[posted_col], errors="coerce", utc=True).dt.tz_localize(None) if posted_col else pd.Series(pd.NaT, index=tx.index)
    transaction_type_series = tx[tt_col].astype(str).str.strip()
    payment_mode_series = tx["_payment_mode"]
    settlement_id_series = tx[sid_col].astype(str).str.strip()

    revenue_price_types = set(settlement_cols_cfg.get("revenue_price_types", []))
    tax_price_types = set(settlement_cols_cfg.get("tax_collection_price_types", []))
    fallback_label_cols_spec = settlement_cols_cfg.get("fallback_label_priority", [])
    resolved_fallback_cols = [c for c in (resolve_col(tx, s) for s in fallback_label_cols_spec) if c]

    frames = []
    for amount_spec, type_spec in settlement_cols_cfg["amount_type_pairs"]:
        amount_col = resolve_col(tx, amount_spec)
        if amount_col is None:
            continue
        type_col = resolve_col(tx, type_spec) if type_spec else None

        amt = pd.to_numeric(tx[amount_col], errors="coerce")
        has_value = amt.notna() & (amt != 0)
        if not has_value.any():
            continue

        rows_idx = tx.index[has_value]
        amt_vals = amt[has_value]

        if type_col is not None:
            label_vals = tx.loc[rows_idx, type_col].astype(str).str.strip()
            label_vals = label_vals.where(tx.loc[rows_idx, type_col].notna(), transaction_type_series.loc[rows_idx])
        else:
            label_vals = pd.Series(None, index=rows_idx, dtype=object)
            for fb_col in resolved_fallback_cols:
                fb_vals = tx.loc[rows_idx, fb_col]
                label_vals = label_vals.where(label_vals.notna(), fb_vals)
            label_vals = label_vals.where(label_vals.notna(), transaction_type_series.loc[rows_idx])
            label_vals = label_vals.astype(str).str.strip()

        type_kind = "price" if type_spec == "price-type" else None
        tt_vals = transaction_type_series.loc[rows_idx]
        oid_vals = order_id_series.loc[rows_idx]
        is_order_level = oid_vals.astype(bool)

        bucket_vals = _bucket_for_vectorized(tt_vals, type_kind, label_vals, revenue_price_types, tax_price_types)

        frame = pd.DataFrame({
            "settlement_id": settlement_id_series.loc[rows_idx].to_numpy(),
            "payment_mode": payment_mode_series.loc[rows_idx].to_numpy(),
            "order_id": oid_vals.where(is_order_level, None).to_numpy(),
            "sku": (sku_series.loc[rows_idx].to_numpy() if sku_col else None),
            "shipment_id": (shipment_series.loc[rows_idx].to_numpy() if shipment_col else None),
            "posted_date": (posted_series.loc[rows_idx].to_numpy() if posted_col else pd.NaT),
            "transaction_type": tt_vals.to_numpy(),
            "category": label_vals.to_numpy(),
            "amount": amt_vals.astype(float).to_numpy(),
            "bucket": bucket_vals,
            "level": np.where(is_order_level.to_numpy(), "Order", "Settlement"),
            "source_amount_col": amount_col,
            "source_type_col": type_col,
            "row_ref": rows_idx.to_numpy(),
        }, index=rows_idx)
        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=cols)

    return pd.concat(frames, ignore_index=True)[cols]


def pivot_expense_ledger(expense_ledger_df, index_cols=("settlement_id",)):
    """
    Wide view of the expense ledger for display: one row per index_cols
    combination, one column per distinct category, values summed. This is
    the "each expense in a separate column" option the client asked for as
    an alternative to the long/row-per-expense view - both are available,
    same underlying ledger, so they can never drift apart.
    """
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=list(index_cols))
    pivot = expense_ledger_df.pivot_table(
        index=list(index_cols), columns="category", values="amount", aggfunc="sum", fill_value=0.0,
    )
    pivot.columns = [str(c) for c in pivot.columns]
    return pivot.reset_index()


# Buckets that represent a real cash deduction/credit against the seller -
# i.e. everything except "Revenue" (already captured from MTR's Invoice
# Amount - re-adding it here would double-count sales) and "Reserve
# Movement" (nets to ~0, cash-flow timing only - see _bucket_for's
# docstring). "Tax Collected (TDS/TCS)" IS included here even though it's
# separately recoverable via the ITR, because it still reduces the actual
# cash Amazon pays out - matching the prior year workbook's own treatment
# of TDS/TCS as part of its Section 1 order-level deduction total.
_CASH_DEDUCTION_BUCKETS = {"Deduction/Credit", "Tax Collected (TDS/TCS)"}


def order_ids_with_settlement_row(expense_ledger_df):
    """
    The full set of order-ids that appear ANYWHERE in the settlement flat
    file at the ORDER level - every bucket (Revenue, Tax Collected,
    Reserve Movement, Deduction/Credit), not just the cash-deduction
    buckets summarize_ledger_by_order() rolls up.

    Why this needs to be separate from summarize_ledger_by_order(): "has
    Amazon settled this order at all" is a broader question than "does it
    have a net cash deduction/credit". An order whose only order-level
    line in that settlement is a bare Principal/tax reversal (no separate
    fee line posted in the same settlement) still genuinely HAS a
    settlement row - it just nets to a zero cash-deduction total. Using
    summarize_ledger_by_order()'s cash-bucket-filtered order-id set for
    order_reco_df's has_settlement_row flag (as build_order_reconciliation
    used to) could mark such an order "settlement pending" even though its
    line is sitting right there in the Expense Ledger sheet - confirmed
    against a client report of exactly this mismatch (has_settlement_row
    False for orders they could manually find in the ledger). Using this
    broader "any order-level row" set instead means has_settlement_row can
    never disagree with what's actually visible in the Expense Ledger
    detail - the same "two views of one ledger must never quietly
    diverge" principle already applied elsewhere in this module.
    """
    if expense_ledger_df is None or expense_ledger_df.empty:
        return set()
    df = expense_ledger_df[expense_ledger_df["level"] == "Order"]
    return set(df["order_id"].dropna().astype(str))


def _pivot_expense_ledger_wide(df, group_keys):
    """
    Shared horizontal-pivot logic for pivot_expense_ledger_by_order() and
    pivot_expense_ledger_by_settlement() below: one row per group_keys
    combination, one column per distinct category (all buckets - Revenue,
    Tax Collected, Reserve Movement, Deduction/Credit - nothing dropped),
    plus two roll-up columns so the wide view can be cross-checked against
    the order/settlement-level totals elsewhere in this engine without
    re-deriving them by hand:
      "Total Revenue"                    - Revenue bucket only
      "Total Deductions/Credits (Cash)"  - Deduction/Credit + Tax Collected
                                            buckets (_CASH_DEDUCTION_BUCKETS)
    """
    settlement_ids = (
        df.groupby(group_keys)["settlement_id"]
        .agg(lambda s: ", ".join(sorted(set(s.astype(str)))))
        .rename("settlement_id(s)")
    )
    pivot = df.pivot_table(index=group_keys, columns="category", values="amount", aggfunc="sum", fill_value=0.0)
    pivot.columns = [str(c) for c in pivot.columns]
    category_cols = list(pivot.columns)

    cash = (
        df[df["bucket"].isin(_CASH_DEDUCTION_BUCKETS)].groupby(group_keys)["amount"].sum()
        .rename("Total Deductions/Credits (Cash)")
    )
    revenue = df[df["bucket"] == "Revenue"].groupby(group_keys)["amount"].sum().rename("Total Revenue")

    out = pivot.join(settlement_ids).join(revenue).join(cash).reset_index()
    out["Total Revenue"] = out["Total Revenue"].fillna(0.0).round(2)
    out["Total Deductions/Credits (Cash)"] = out["Total Deductions/Credits (Cash)"].fillna(0.0).round(2)
    for c in category_cols:
        out[c] = out[c].round(2)

    ordered_cols = list(group_keys) + ["settlement_id(s)", "Total Revenue"] + category_cols + [
        "Total Deductions/Credits (Cash)",
    ]
    return out[[c for c in ordered_cols if c in out.columns]]


def pivot_expense_ledger_by_order(expense_ledger_df):
    """
    Horizontal (wide) view of the ORDER-level expense ledger for the
    downloadable report: one row per order_id (+ payment_mode + cutoff
    status when present), one column per distinct fee/tax/revenue
    category, values summed - instead of one row per individual line item,
    which runs into the lakhs for a full year (exactly the "redesign as
    horizontal, wherever possible, to keep it compact/readable" ask). No
    detail is lost - every category that would have been its own set of
    rows becomes its own column instead, so any order's full fee breakup
    is still visible, just read across a row instead of down a column. The
    full line-by-line long-format ledger (settlement_id | order_id | sku |
    posted_date | row_ref | ...) is still exported separately as "Expense
    Ledger (Full Detail)" for anyone who needs to trace a figure back to
    its exact source row.
    """
    group_keys = ["order_id", "payment_mode"]
    empty_cols = group_keys + ["settlement_id(s)", "Total Revenue", "Total Deductions/Credits (Cash)"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=empty_cols)

    df = expense_ledger_df[expense_ledger_df["level"] == "Order"].copy()
    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    if "cutoff_status" in df.columns:
        group_keys = group_keys + ["cutoff_status"]

    out = _pivot_expense_ledger_wide(df, group_keys)
    return out.sort_values("order_id").reset_index(drop=True)


def pivot_expense_ledger_by_settlement(expense_ledger_df):
    """
    Same idea as pivot_expense_ledger_by_order() above, but for the
    SETTLEMENT-level (non-order-attributable) lines - Advertising, MCF
    fulfilment fee, Storage/Warehouse, Debt adjustments, Reserve movements,
    etc. (see summarize_settlement_level_adjustments's docstring) - one row
    per settlement instead of one row per category per settlement.
    """
    group_keys = ["settlement_id", "payment_mode"]
    empty_cols = group_keys + ["Total Revenue", "Total Deductions/Credits (Cash)"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=empty_cols)

    df = expense_ledger_df[expense_ledger_df["level"] == "Settlement"].copy()
    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    if "cutoff_status" in df.columns:
        group_keys = group_keys + ["cutoff_status"]

    out = _pivot_expense_ledger_wide(df, group_keys)
    return out.drop(columns=["settlement_id(s)"], errors="ignore").sort_values("settlement_id").reset_index(drop=True)


def summarize_ledger_by_order(expense_ledger_df, known_order_ids=None):
    """
    Order-level rollup of the expense ledger's cash deduction/credit lines
    (see _CASH_DEDUCTION_BUCKETS above), for joining into the MTR-based
    reconciliation table (engine/amazon_reco.py):
        order_id | order_deductions | order_credits

    known_order_ids: optional set of order-ids that exist in the MTR
    ledger. Some flat file rows carry an order-id-shaped reference that
    ISN'T a real Amazon marketplace order - most commonly MCF (Multi-
    Channel Fulfilment) references, e.g. "S02-xxxxxxx-xxxxxxx", for Amazon
    fulfilling a Shopify/D2C order. The prior year's own workbook explicitly
    booked MCF COD collections/fees under the Shopify P&L, not Amazon's -
    so when known_order_ids is supplied, rows whose order-id isn't in it
    are EXCLUDED here (they'd never join to anything in the MTR-driven
    order table anyway) and should instead be surfaced separately via
    non_mtr_order_level_items() below, so that money is still visible
    somewhere rather than silently vanishing from every report.
    """
    cols = ["order_id", "order_deductions", "order_credits"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=cols)

    df = expense_ledger_df[
        (expense_ledger_df["level"] == "Order") & (expense_ledger_df["bucket"].isin(_CASH_DEDUCTION_BUCKETS))
    ].copy()
    if known_order_ids is not None:
        df = df[df["order_id"].isin(known_order_ids)]
    if df.empty:
        return pd.DataFrame(columns=cols)

    grouped = df.groupby("order_id")["amount"].agg(
        order_deductions=lambda s: s[s < 0].sum(),
        order_credits=lambda s: s[s > 0].sum(),
    )
    return grouped.reset_index()[cols]


def non_mtr_order_level_items(expense_ledger_df, known_order_ids):
    """
    The complement of summarize_ledger_by_order()'s known_order_ids filter:
    order-level cash deduction/credit lines whose order-id does NOT appear
    in MTR - predominantly MCF (Shopify/D2C fulfilled-by-Amazon) pass-
    through collections and fees. Kept visible here (grouped by category)
    rather than silently dropped, even though they're excluded from the
    Amazon marketplace waterfall - see summarize_ledger_by_order's
    docstring for why they're excluded from that total.
        category | bucket | amount | order_count
    """
    cols = ["category", "bucket", "amount", "order_count"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=cols)

    df = expense_ledger_df[
        (expense_ledger_df["level"] == "Order") & (expense_ledger_df["bucket"].isin(_CASH_DEDUCTION_BUCKETS))
    ].copy()
    df = df[~df["order_id"].isin(known_order_ids or set())]
    if df.empty:
        return pd.DataFrame(columns=cols)

    grouped = df.groupby(["category", "bucket"]).agg(
        amount=("amount", "sum"), order_count=("order_id", "nunique"),
    ).reset_index()
    return grouped[cols]


def summarize_settlement_level_adjustments(expense_ledger_df):
    """
    Settlement-level (non-order-attributable) rollup of the expense
    ledger's cash deduction/credit lines (see _CASH_DEDUCTION_BUCKETS
    above), by settlement_id and category - the equivalent of the prior
    workbook's "Section 2" columns (Advertising, MCF fulfilment fee,
    Storage/Warehouse, Debt/Other), kept category-by-category rather than
    merged. Reserve movements are deliberately excluded here too (see
    reserve_movement_summary() below for that memo figure on its own).
    """
    cols = ["settlement_id", "payment_mode", "category", "bucket", "amount"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=cols)

    df = expense_ledger_df[
        (expense_ledger_df["level"] == "Settlement") & (expense_ledger_df["bucket"].isin(_CASH_DEDUCTION_BUCKETS))
    ].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    grouped = df.groupby(["settlement_id", "payment_mode", "category", "bucket"])["amount"].sum().reset_index()
    return grouped[cols]


def reserve_movement_summary(expense_ledger_df):
    """Net reserve movement (Current Reserve withheld + Previous Reserve
    released) - a memo figure only, see _bucket_for's docstring."""
    if expense_ledger_df is None or expense_ledger_df.empty:
        return 0.0
    return float(expense_ledger_df[expense_ledger_df["bucket"] == "Reserve Movement"]["amount"].sum())
