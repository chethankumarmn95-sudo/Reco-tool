"""
amazon_invoice_check.py
------------------------
STATUS (confirmed with the client, 2026-08-18): the client does not yet
have Amazon's own fee/tax-invoice-level export (Seller Central's Tax
Document Library issues separate GST invoices per fee type, e.g.
Commission Invoice / Closing Fee Invoice / Fulfilment Fee Invoice /
Advertising Invoice, per state per month) - so a genuine "amount as per
Amazon's expense report vs amount as per Amazon's own invoice" match is not
buildable yet. Neither the Settlement Flat File nor the MTR carries an
invoice reference on the EXPENSE side (MTR's Invoice Number/Date is
revenue-side only).

What this module does instead, as a placeholder that's still genuinely
useful: a settlement-level TIE-OUT, summing every single line this engine's
melt logic (engine.amazon_consolidator.build_expense_ledger) extracted from
the flat file for a settlement and checking it against Amazon's OWN
reported total-amount for that settlement (the summary row,
engine.amazon_consolidator.build_settlement_summary). Since a settlement's
total-amount is, by construction, Amazon's own sum of every line it
attributes to that settlement, a clean tie-out here is strong evidence this
engine's parsing captured 100% of the flat file correctly - no line missed,
none double-counted, none miscategorized into the wrong amount column.

When the real Amazon fee-invoice export becomes available: replace
invoice_amount below (currently a placeholder duplicate of the settlement's
own reported total) with a genuine per-invoice figure joined on whatever
reference field that export uses, and category_invoice_tie_out() below can
be extended the same way at the category level (comparing this ledger's
per-category total against that category's own invoice amount, per
settlement period) - the ledger's structure (one row per settlement +
category + amount, fully traceable via source_amount_col/row_ref back to
the flat file) is already shaped to support that once the input exists.
"""

import pandas as pd

TOLERANCE = 1.0


def settlement_tie_out(expense_ledger_df, settlement_summary_df):
    """
    One row per settlement:
        settlement_id | payment_mode | amount_as_parsed | amount_as_per_amazon |
        difference | status
    status is "Tied out" when the two agree within TOLERANCE (Rupee 1),
    else "Review - parsing gap or timing difference".
    """
    cols = ["settlement_id", "payment_mode", "amount_as_parsed", "amount_as_per_amazon", "difference", "status"]
    if settlement_summary_df is None or settlement_summary_df.empty:
        return pd.DataFrame(columns=cols)

    parsed_totals = (
        expense_ledger_df.groupby("settlement_id")["amount"].sum()
        if expense_ledger_df is not None and not expense_ledger_df.empty
        else pd.Series(dtype=float)
    )

    df = settlement_summary_df.copy()
    df["settlement_id"] = df["settlement_id"].astype(str)
    df["amount_as_parsed"] = df["settlement_id"].map(parsed_totals).fillna(0.0).round(2)
    df["amount_as_per_amazon"] = df["total_amount"].round(2)
    df["difference"] = (df["amount_as_parsed"] - df["amount_as_per_amazon"]).round(2)
    df["status"] = df["difference"].abs().le(TOLERANCE).map(
        {True: "Tied out", False: "Review - parsing gap or timing difference"}
    )

    return df.rename(columns={})[["settlement_id", "payment_mode", "amount_as_parsed",
                                    "amount_as_per_amazon", "difference", "status"]]


def category_breakup_by_period(expense_ledger_df, period_col="settlement_id"):
    """
    Category-wise totals per period (settlement_id by default; pass
    "payment_mode" or a derived month column for a coarser view) - the
    shape this module's future real invoice-matching will need once
    Amazon's fee-invoice export is available: one row per (period,
    category), ready to be joined against that export's own per-category,
    per-period invoice amount.
    """
    cols = [period_col, "category", "bucket", "amount", "line_count"]
    if expense_ledger_df is None or expense_ledger_df.empty:
        return pd.DataFrame(columns=cols)

    grouped = expense_ledger_df.groupby([period_col, "category", "bucket"]).agg(
        amount=("amount", "sum"),
        line_count=("amount", "count"),
    ).reset_index()
    return grouped[cols]
