"""
amazon_bank.py
--------------
Matches Amazon settlements to the bank statement.

An Amazon settlement is exactly like a COD courier's remittance in the
Shopify pipeline (engine/bank.py's build_cod_settlement_batches /
match_batches_to_bank): ONE lump bank credit that covers many orders'
worth of money at once, with no clean per-order UTR to match against on
either side. The settlement's own total-amount and deposit-date (captured
in engine.amazon_consolidator.build_settlement_summary) are the reliable
match keys - so this module is a thin, settlement-shaped wrapper around
the SAME amount+date batch-matching logic already built and tested for
COD, rather than a new matching algorithm.
"""

import pandas as pd

from .bank import match_batches_to_bank  # amount+date batch matcher, reused as-is


def build_settlement_batches(settlement_summary_df):
    """
    Reshapes the settlement summary into the generic "batch" shape
    match_batches_to_bank() expects: batch_id | source | settlement_date |
    order_ids | batch_amount.

    Settlements with a zero (or negative, after Amazon's own fee deductions
    exceeded revenue) total-amount are dropped before matching - there's no
    bank credit to look for when Amazon itself reports nothing was paid out,
    exactly like the prior year's "Nil / Negative" settlement status.
    """
    cols = ["batch_id", "source", "settlement_date", "order_ids", "batch_amount"]
    if settlement_summary_df is None or settlement_summary_df.empty:
        return pd.DataFrame(columns=cols)

    df = settlement_summary_df[settlement_summary_df["total_amount"] > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    out = pd.DataFrame()
    out["batch_id"] = df["settlement_id"].astype(str)
    out["source"] = "Amazon " + df["payment_mode"].astype(str)
    out["settlement_date"] = pd.to_datetime(df["deposit_date"], errors="coerce")
    out["order_ids"] = [[] for _ in range(len(df))]
    out["batch_amount"] = df["total_amount"].round(2)
    return out[cols]


def settlement_register(settlement_summary_df, bank_ledger_df, date_window_days=5):
    """
    One row per settlement (the Amazon equivalent of the prior workbook's
    "4. Settlement Register"):
        settlement_id | payment_mode | start_date | end_date | deposit_date |
        settlement_amount | status | bank_amount | bank_date | match_note

    status is one of "Nil / Negative" (nothing due - see
    build_settlement_batches above), "Matched" (bank credit found by
    amount+date), or "Unmatched" (needs manual review).
    """
    cols = [
        "settlement_id", "payment_mode", "start_date", "end_date", "deposit_date",
        "settlement_amount", "status", "bank_amount", "bank_date", "match_note",
    ]
    if settlement_summary_df is None or settlement_summary_df.empty:
        return pd.DataFrame(columns=cols)

    batches = build_settlement_batches(settlement_summary_df)
    match = match_batches_to_bank(batches, bank_ledger_df, date_window_days=date_window_days)

    df = settlement_summary_df.copy()
    df["settlement_id"] = df["settlement_id"].astype(str)

    if not match.empty:
        match = match.rename(columns={"batch_id": "settlement_id"})
        df = df.merge(match, on="settlement_id", how="left")
    else:
        df["matched"] = False
        df["bank_amount"] = None
        df["bank_date"] = pd.NaT
        df["match_note"] = None

    def _status(row):
        if row["total_amount"] <= 0:
            return "Nil / Negative"
        return "Matched" if bool(row.get("matched")) else "Unmatched"

    df["status"] = df.apply(_status, axis=1)
    df = df.rename(columns={"total_amount": "settlement_amount"})
    return df[cols]


def split_received_by_cutoff(settlement_register_df, cutoff_date):
    """
    "Bank receipts must also be split by receipt date" (the client's own
    rule - see engine/amazon_reco.py's module docstring). Independent of
    which settlement a receipt belongs to, or of that settlement's own
    deposit-date: this only asks "did the ACTUAL matched bank credit
    (bank_date) land on/before the cut-off?"

    Returns (received_by_cutoff, received_after_cutoff) as floats, summed
    over whatever rows are in settlement_register_df - callers should pass
    the "within cut-off" settlement subset (see classify_settlements_by_
    cutoff) so received_after_cutoff means specifically "settlement was
    within cut-off per Amazon's own report, but the bank credit lagged
    past it" (receipts in transit), not "receipt for a subsequent
    settlement" (those are already excluded/tracked separately via
    subsequent_settlements_summary).
    """
    if settlement_register_df is None or settlement_register_df.empty or cutoff_date is None:
        total = float(settlement_register_df["bank_amount"].fillna(0).sum()) if (
            settlement_register_df is not None and not settlement_register_df.empty) else 0.0
        return total, 0.0

    bank_dates = pd.to_datetime(settlement_register_df["bank_date"], errors="coerce")
    amounts = settlement_register_df["bank_amount"].fillna(0)
    by_cutoff = float(amounts[bank_dates <= cutoff_date].sum())
    after_cutoff = float(amounts[bank_dates > cutoff_date].sum())
    return by_cutoff, after_cutoff


def settlement_register_summary(settlement_register_df):
    """
    Payment-mode-wise (COD vs Online) match-rate summary, mirroring the
    prior workbook's "5. Recon Summary" section B (Settlement Summary -
    All Matched):
        Type | Count | Settlement Total | Bank Receipts | Variance |
        Matched | Nil/Negative | Unmatched
    """
    cols = ["Type", "Count", "Settlement Total (Rs)", "Bank Receipts (Rs)", "Variance (Rs)",
            "Matched", "Nil/Negative", "Unmatched"]
    if settlement_register_df is None or settlement_register_df.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for mode, grp in settlement_register_df.groupby("payment_mode"):
        settlement_total = float(grp["settlement_amount"].sum())
        bank_total = float(grp["bank_amount"].fillna(0).sum())
        rows.append({
            "Type": mode,
            "Count": len(grp),
            "Settlement Total (Rs)": round(settlement_total, 2),
            "Bank Receipts (Rs)": round(bank_total, 2),
            "Variance (Rs)": round(settlement_total - bank_total, 2),
            "Matched": f"{int((grp['status'] == 'Matched').sum())}/{len(grp)}",
            "Nil/Negative": f"{int((grp['status'] == 'Nil / Negative').sum())}/{len(grp)}",
            "Unmatched": f"{int((grp['status'] == 'Unmatched').sum())}/{len(grp)}",
        })

    total_settlement = float(settlement_register_df["settlement_amount"].sum())
    total_bank = float(settlement_register_df["bank_amount"].fillna(0).sum())
    n = len(settlement_register_df)
    rows.append({
        "Type": "TOTAL",
        "Count": n,
        "Settlement Total (Rs)": round(total_settlement, 2),
        "Bank Receipts (Rs)": round(total_bank, 2),
        "Variance (Rs)": round(total_settlement - total_bank, 2),
        "Matched": f"{int((settlement_register_df['status'] == 'Matched').sum())}/{n}",
        "Nil/Negative": f"{int((settlement_register_df['status'] == 'Nil / Negative').sum())}/{n}",
        "Unmatched": f"{int((settlement_register_df['status'] == 'Unmatched').sum())}/{n}",
    })
    return pd.DataFrame(rows)[cols]
