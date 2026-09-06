"""
settlement.py
-------------
Payment-Gateway-wise Settlement Report: for each gateway (Gokwik,
Razorpay, Delhivery COD, Shiprocket COD, ...), how much has actually
settled, how much is still pending, and how much was deducted as fees /
charges - shown separately per gateway rather than only as one lump total
across all of them.

Two different data sources feed this, deliberately:
  - "Settlement Done" and "Deductions" come from the gateway's OWN
    settlement export (engine/consolidator.py's consolidated receipt
    ledger) - money the gateway itself says it has already remitted.
  - "Pending for Settlement" comes from the ORDER side (engine/reco.py's
    Reco working table): orders that are Delivered (so payment is
    genuinely due) but haven't shown up in that gateway's settlement
    export at all yet. A gateway that hasn't settled an order yet won't
    have a row in the consolidated ledger to look at, so this can only be
    computed from the order side, not the receipt side.
"""

import pandas as pd

from .bank import to_naive_timestamp
from .reco import _COD_SETTLEMENT_PENDING_PHRASES
from .attribution import cod_component_of_gateway_label

UNATTRIBUTED_COD = "Unattributed COD"
UNATTRIBUTED_PREPAID = "Unattributed Prepaid"


def _partner_name_in_label(label):
    """"Delhivery COD" -> "delhivery", "Shiprocket COD" -> "shiprocket" -
    used to connect a COD gateway back to the courier that actually
    collects for it, since COD collection is done by whichever courier
    delivered the order, not a universal "COD gateway"."""
    return label.lower().replace("cod", "").strip()


def expected_gateway_for_order(payment_method_text, delivery_partner, gateway_configs):
    """
    Best-effort guess at which gateway/COD-channel is responsible for
    settling an order that hasn't been received yet, so unsettled money
    can be attributed to a gateway instead of only appearing as one
    unexplained lump sum.

    Order of preference:
      1. The order's own Payment Method text directly names a Prepaid
         gateway (e.g. Shopify's Payment Method literally says
         "Razorpay" or "Gokwik") - trust that over anything else.
      2. Otherwise, if this looks like a COD order (Payment Method says
         so, or is blank/unclear), attribute it to whichever COD gateway
         matches the courier that actually delivered it - e.g. Delhivery
         -> "Delhivery COD", Shiprocket -> "Shiprocket COD".
      3. If none of that resolves cleanly, "Unattributed COD" /
         "Unattributed Prepaid" - so the amount is never silently
         dropped, it just needs a human to confirm which gateway it
         really belongs to.

    2026-09-06 (round 18, client-reported): fixed a latent bug in how a
    genuinely missing value was blanked here - `payment_method_text or ""`
    looks like it blanks a missing value, but a NaN FLOAT (pandas' own
    representation of a missing value after a merge - not Python's None)
    is truthy, so `nan or ""` evaluates to `nan` itself, not "" - which
    then stringifies to the literal text "nan" instead of blank. That
    made `not text` False (the text isn't actually empty, it's "nan"),
    so a genuinely COD order whose Payment Method came through as a
    post-merge NaN (not a raw None) skipped the COD-courier-by-
    delivery_partner branch below entirely and fell all the way to the
    "Unattributed Prepaid" catch-all - one of the exact failure modes
    behind the client's own "Unattributed Prepaid" report (see
    engine/settlement_pending.py::build_settlement_pending_report()'s own
    docstring). Using pd.isna() (true for both None and NaN) rather than
    Python truthiness fixes this for both parameters.
    """
    text = "" if pd.isna(payment_method_text) else str(payment_method_text).strip().lower()
    partner = "" if pd.isna(delivery_partner) else str(delivery_partner).strip().lower()

    prepaid_cfgs = [c for c in gateway_configs if str(c.get("payment_mode", "")).lower() != "cod"]
    cod_cfgs = [c for c in gateway_configs if str(c.get("payment_mode", "")).lower() == "cod"]

    for cfg in prepaid_cfgs:
        if cfg["label"].lower() in text:
            return cfg["label"]

    is_cod_text = bool(text) and any(k in text for k in ("cod", "cash on delivery", "cashondelivery"))
    if is_cod_text or not text:
        for cfg in cod_cfgs:
            if partner and _partner_name_in_label(cfg["label"]) in partner:
                return cfg["label"]
        if is_cod_text:
            return UNATTRIBUTED_COD

    return UNATTRIBUTED_PREPAID


SAME_MONTH = "Same Month"
NEXT_MONTH = "Next Month"
OTHER_PERIOD = "Other Period"


def _settlement_bucket(order_id, receipt_date, period_by_order_id, period_end):
    """
    Which of the three "Gateway Settlement" sheet categories a single
    gateway transaction row belongs to (client-reported: without this
    split, a Payment Gateway report covering a WIDER date range than the
    Sales report - e.g. Sales Apr-Jun, Gateway report Apr-Jul - silently
    let July's settlements count as if they were part of the Apr-Jun
    reconciliation period, with no way to see which settlements actually
    landed on time vs late):

      - OTHER_PERIOD: the order itself doesn't belong to this
        reconciliation period at all (Previous/Subsequent period, or not
        found anywhere) - e.g. a delayed settlement for an order sold
        last period. Kept visible as its own bucket rather than either
        silently mixed into this period's own total or silently dropped.
      - SAME_MONTH: the order belongs to this period AND the gateway's
        own receipt/settlement date falls within this period's own date
        span (or has no date to check at all - treated as on-time rather
        than manufacturing a false "late" flag from missing data).
      - NEXT_MONTH: the order belongs to this period, but the gateway
        settled it AFTER this period's own date span ended - the
        "settled in the following month" spillover case.

    Client-reported (2026-08-21): crashed with "Cannot compare tz-naive and
    tz-aware timestamps" the moment a real Shopify/gateway export was used -
    Shopify's own "Created at"/receipt-date columns commonly carry a
    timezone offset (e.g. "+0530"), so receipt_date here can be tz-aware
    while period_end (derived from reco_df's own created_at column, via a
    separate pd.to_datetime call in gateway_settlement_summary below) can
    come out tz-naive, or the two can carry different offsets - pandas
    raises rather than comparing across that mismatch. Both sides are
    stripped to naive via to_naive_timestamp() (see engine/bank.py, where
    this exact class of bug was already fixed once for this same underlying
    reason) immediately before comparing, so this only ever compares two
    naive Timestamps regardless of what tz-awareness either source carried.
    """
    period = (period_by_order_id or {}).get(str(order_id), "Order not found")
    if period != "This period":
        return OTHER_PERIOD
    if pd.isna(receipt_date) or period_end is None or pd.isna(period_end):
        return SAME_MONTH
    receipt_ts = to_naive_timestamp(pd.Timestamp(receipt_date))
    period_end_ts = to_naive_timestamp(pd.Timestamp(period_end))
    return SAME_MONTH if receipt_ts <= period_end_ts else NEXT_MONTH


def gateway_settlement_summary(consolidated_df, reco_df, gateway_configs, period_by_order_id=None):
    """
    One row per gateway:
        Payment Gateway | Settlement Done - This Period (Same Month) |
        Settlement Done - This Period (Next Month) |
        Settlement Done - Other Period | Settlement Done | Deductions |
        Pending for Settlement | Pending Orders | Total (Settled + Pending)

    Settlement Done = receipt - deduction - refund, for that gateway, from
                      what it has actually reported settling (mirrors the
                      existing "Net settlement" headline figure, just
                      split by gateway instead of totalled across all of
                      them - see engine/summary.py headline_totals()) -
                      now further split into the three period/timing
                      buckets described in _settlement_bucket() above, so
                      the sheet's own total still equals the sum of the
                      three (nothing silently added or dropped), while
                      making it possible to see exactly how much of that
                      total was on-time vs late vs a different period's
                      settlement entirely.
    Deductions      = fees/charges/tax that gateway held back (kept as one
                      overall figure - the split above only applies to
                      Settlement Done, per the client's own ask).
    Pending for Settlement / Pending Orders = Delivered orders attributed
                      to that gateway (see expected_gateway_for_order()
                      above) where money is still owed - i.e. orders with
                      final_delivery_status == "Delivered" and diff > 0
                      that this gateway hasn't remitted yet.

    period_by_order_id: dict order_id (str) -> "This period" / "Previous
    period" / "Subsequent period" / "Order not found" (see
    engine/period.py classify_order_periods()) - the SAME classification
    already used for the UTR-level bank reconciliation, so "this period"
    means exactly the same thing everywhere in the app. Pass None to skip
    the split (every rupee then falls under "Same Month", matching this
    function's behaviour before the split existed - a safe default for
    any caller that hasn't computed the classification).
    """
    settlement_by_bucket = {}
    if consolidated_df is not None and not consolidated_df.empty:
        df = consolidated_df.copy()

        period_end = None
        if reco_df is not None and len(reco_df) and "created_at" in reco_df.columns:
            valid_dates = pd.to_datetime(reco_df["created_at"], errors="coerce").dropna()
            period_end = valid_dates.max() if len(valid_dates) else None

        df["_bucket"] = [
            _settlement_bucket(oid, rd, period_by_order_id, period_end)
            for oid, rd in zip(df["order_id"], df.get("receipt_date", pd.Series([pd.NaT] * len(df))))
        ]

        payments = df[~df["is_refund"]]
        refunds = df[df["is_refund"]]
        receipt = payments.groupby("source")["amount"].sum()
        deduction = payments.groupby("source")["deduction"].sum()
        refund_total = refunds.groupby("source")["amount"].sum()
        settlement_done = receipt.sub(deduction, fill_value=0).sub(refund_total, fill_value=0)

        receipt_b = payments.groupby(["source", "_bucket"])["amount"].sum()
        deduction_b = payments.groupby(["source", "_bucket"])["deduction"].sum()
        refund_b = refunds.groupby(["source", "_bucket"])["amount"].sum()
        settlement_done_b = receipt_b.sub(deduction_b, fill_value=0).sub(refund_b, fill_value=0)
        for (label, bucket), value in settlement_done_b.items():
            settlement_by_bucket[(label, bucket)] = value
    else:
        receipt = pd.Series(dtype=float)
        deduction = pd.Series(dtype=float)
        settlement_done = pd.Series(dtype=float)

    pending_by_gateway = pd.Series(dtype=float)
    pending_orders_by_gateway = pd.Series(dtype="int64")

    if reco_df is not None and len(reco_df) and "final_delivery_status" in reco_df.columns:
        due = reco_df[(reco_df["final_delivery_status"] == "Delivered") & (reco_df["diff"] > 1)].copy()
        if len(due):
            if "payment_method" not in due.columns:
                due["payment_method"] = ""
            if "delivery_partner" not in due.columns:
                due["delivery_partner"] = ""
            due["_expected_gateway"] = due.apply(
                lambda r: expected_gateway_for_order(r["payment_method"], r["delivery_partner"], gateway_configs),
                axis=1,
            )
            pending_by_gateway = due.groupby("_expected_gateway")["diff"].sum()
            pending_orders_by_gateway = due.groupby("_expected_gateway")["order_id"].count()

    all_labels = sorted(
        {cfg["label"] for cfg in gateway_configs}
        | set(pending_by_gateway.index)
        | set(receipt.index)
    )

    rows = []
    for label in all_labels:
        rows.append({
            "Payment Gateway": label,
            "Settlement Done - This Period (Same Month)": round(float(settlement_by_bucket.get((label, SAME_MONTH), 0.0)), 2),
            "Settlement Done - This Period (Next Month)": round(float(settlement_by_bucket.get((label, NEXT_MONTH), 0.0)), 2),
            "Settlement Done - Other Period": round(float(settlement_by_bucket.get((label, OTHER_PERIOD), 0.0)), 2),
            "Settlement Done": round(float(settlement_done.get(label, 0.0)), 2),
            "Deductions": round(float(deduction.get(label, 0.0)), 2),
            "Pending for Settlement": round(float(pending_by_gateway.get(label, 0.0)), 2),
            "Pending Orders": int(pending_orders_by_gateway.get(label, 0)),
        })

    out_cols = [
        "Payment Gateway",
        "Settlement Done - This Period (Same Month)", "Settlement Done - This Period (Next Month)",
        "Settlement Done - Other Period", "Settlement Done", "Deductions",
        "Pending for Settlement", "Pending Orders", "Total (Settled + Pending)",
    ]
    if not rows:
        return pd.DataFrame(columns=out_cols)

    out = pd.DataFrame(rows)
    out["Total (Settled + Pending)"] = (out["Settlement Done"] + out["Pending for Settlement"]).round(2)
    return out[out_cols].sort_values("Payment Gateway").reset_index(drop=True)


# ---------------------------------------------------------------------------
# "Gateway Settlement overall" / "Gateway Recon Recoperiod" - client-reported
# 2026-08-30 (items 5/6), replacing the single gateway_settlement_summary()
# sheet above with the client's own reference workbook's two-sheet design.
# ---------------------------------------------------------------------------
# "Gateway Settlement overall" (broad, all-time-per-run audit view) is
# sourced ENTIRELY from the already-computed 'Bank Reco (UTR-wise)' sheet
# (engine.bank.bank_reconciliation_by_utr) grouped by ITS OWN "Payment
# Gateway" column - the client's own workbook builds it with plain SUMIFS
# formulas against that sheet, so replicating it in Python off the same
# dataframe keeps the two sheets permanently consistent with each other
# (a rupee that moves on Bank Reco (UTR-wise) always moves the same way
# here) rather than recomputing settlement-by-gateway a second, possibly
# divergent way. This deliberately REPLACES gateway_settlement_summary()
# above as the "Gateway Settlement" sheet's data source; that older
# function grouped by consolidated_df's raw, UN-refined "source" column
# (so a Gokwik-routed order stayed "Gokwik", never "easebuzz"/"payu") and
# used a coarser 3-bucket scheme that conflated several different real
# scenarios (previous-period-settled-now, subsequent-period-still-
# subsequent, and "order not found at all") into one "Other Period"
# figure - both fixed here.
GATEWAY_SETTLEMENT_OVERALL_COLUMNS = [
    "Payment Gateway",
    "Settlement Done - This Period (Same Month)",
    "Settlement Done - This Period (Next Month)",
    "Settlement Done - Other Period",
    "Order ID Not Found - Settled",
    "Order ID Not Found - Settled After Reco Period",
    "Total Settlement Done",
    "PG Deductions Reco period",
    "PG Deductions other period",
]

_GATEWAY_OVERALL_SOURCE_COLS = {
    "Settlement Done - This Period (Same Month)": "Amount (this period settled this period)",
    "Settlement Done - This Period (Next Month)": "Amount (this period transaction Settled Subsequent Period)",
    "Settlement Done - Other Period": "Settled (previous period transaction settled this period)",
    "Order ID Not Found - Settled": "Order ID Not Found - Settled During Reco Period",
    "Order ID Not Found - Settled After Reco Period": "Order ID Not Found - Settled After Reco Period",
}


def gateway_settlement_overall(utr_bank_reco_df, reco_df):
    """
    One row per REFINED gateway label (Payment Gateway column on 'Bank
    Reco (UTR-wise)' - already one-gateway-per-UTR post the item-7 fix in
    engine.attribution.build_payment_gateway_lookups):
        Payment Gateway | Settlement Done - This Period (Same Month) |
        Settlement Done - This Period (Next Month) |
        Settlement Done - Other Period | Order ID Not Found - Settled |
        Order ID Not Found - Settled After Reco Period |
        Total Settlement Done | PG Deductions Reco period |
        PG Deductions other period

    The five "Settlement Done"/"Order ID Not Found" columns are a
    straight per-gateway SUM of 'Bank Reco (UTR-wise)'s own like-named
    columns (utr_bank_reco_df - see engine.bank.bank_reconciliation_by_utr
    for what each one means) - Total Settlement Done is their row sum,
    matching the client's own `=SUM(B:F)` formula exactly.

    "PG Deductions Reco period" sums Reco working's own total_deduction
    column (reco_df), grouped by "Payment Provider" (the same refined,
    order-level gateway attribution engine.attribution.
    attach_payment_columns already computes) rather than by the UTR-level
    "Payment Gateway" column above - matching the client's own formula,
    which deliberately sources deductions from Reco working rather than
    Bank Reco (UTR-wise).

    "PG Deductions other period" is always 0 here - a disclosed gap, not
    guessed at. In the client's own reference workbook this is a literal,
    hand-typed number in every row (not a formula), evidently sourced from
    outside this engine entirely (a broader-window gateway deduction
    figure this tool has no equivalent data for, since its own pipeline
    is date-filtered to the selected period). Left as 0 rather than
    invented.
    """
    if utr_bank_reco_df is None or utr_bank_reco_df.empty or "Payment Gateway" not in utr_bank_reco_df.columns:
        return pd.DataFrame(columns=GATEWAY_SETTLEMENT_OVERALL_COLUMNS)
    missing = [c for c in _GATEWAY_OVERALL_SOURCE_COLS.values() if c not in utr_bank_reco_df.columns]
    if missing:
        return pd.DataFrame(columns=GATEWAY_SETTLEMENT_OVERALL_COLUMNS)

    grouped = utr_bank_reco_df.groupby("Payment Gateway")[list(_GATEWAY_OVERALL_SOURCE_COLS.values())].sum()
    grouped = grouped.rename(columns={v: k for k, v in _GATEWAY_OVERALL_SOURCE_COLS.items()})

    deduction_by_gateway = pd.Series(dtype=float)
    if reco_df is not None and len(reco_df) and "Payment Provider" in reco_df.columns and "total_deduction" in reco_df.columns:
        # 2026-09-06 (round 16, order #30456): engine.attribution.combine_
        # prepaid_and_cod_label() can now hand "Payment Provider" a
        # COMBINED "<courier> COD, <prepaid provider>" label for a genuine
        # part-prepaid/part-COD order - grouped here under its COD leg
        # (cod_component_of_gateway_label()) so it lands in the SAME
        # "Payment Gateway" row 'Bank Reco (UTR-wise)' already uses for
        # that courier, rather than creating its own orphaned combined-
        # label row with no matching Settlement Done figure to reconcile
        # against.
        deduction_by_gateway = reco_df.groupby(
            reco_df["Payment Provider"].apply(cod_component_of_gateway_label)
        )["total_deduction"].sum()

    all_labels = sorted(set(grouped.index) | {i for i in deduction_by_gateway.index if pd.notna(i) and str(i).strip()})
    if not all_labels:
        return pd.DataFrame(columns=GATEWAY_SETTLEMENT_OVERALL_COLUMNS)

    out = grouped.reindex(all_labels).fillna(0.0)
    out["Total Settlement Done"] = out[list(_GATEWAY_OVERALL_SOURCE_COLS.keys())].sum(axis=1)
    out["PG Deductions Reco period"] = deduction_by_gateway.reindex(all_labels).fillna(0.0)
    out["PG Deductions other period"] = 0.0
    out = out.reset_index().rename(columns={"index": "Payment Gateway"})
    out = out[GATEWAY_SETTLEMENT_OVERALL_COLUMNS].round(2)

    # 2026-09-06 (round 21, client-reported): drop a gateway row that
    # contributed genuinely NOTHING this period - every "Settlement Done"/
    # "Order ID Not Found" figure AND both deduction columns are all
    # exactly 0. Client's own example: a "Gokwik" row - Gokwik is a
    # checkout aggregator, not the rail that actually moves money (see
    # engine/attribution.py's module docstring), so once every order it
    # touched has been refined to its real downstream processor (easebuzz,
    # PayU, ...) for BOTH Bank Reco (UTR-wise)'s "Payment Gateway" AND Reco
    # working's "Payment Provider", a bare "Gokwik" label has nothing left
    # to sum under it - it's a phantom row, not a genuine gateway with a
    # real (if small) figure to report. A general "all-zero" filter rather
    # than hardcoding "Gokwik" specifically, since the same phantom-row
    # shape could equally appear for any other gateway label that ends up
    # fully refined away for a given period.
    all_zero = (out[["Total Settlement Done", "PG Deductions Reco period", "PG Deductions other period"]]
                .abs().lt(0.01).all(axis=1))
    out = out.loc[~all_zero].reset_index(drop=True)
    return out


GATEWAY_RECON_RECOPERIOD_COLUMNS = [
    "Payment Gateway", "Group", "Order Value", "Gateway deduction", "Refund",
    "Pending Settlement", "Part prepaid and part post paid",
    "Order Value partially not received", "Received in Bank", "Check", "Diff",
]


def _pending_query_phrases(gateway_label):
    """The exact query-column text engine.reco.refine_queries_with_settlement_status
    would have assigned an order still-pending under this gateway - reused
    here (not re-derived) so "Pending Settlement" below always agrees with
    what the Query column on Reco working actually shows for the same
    order. A COD gateway also matches its own " not reflecting"/"
    partially reflecting" suffixed variants."""
    if gateway_label in _COD_SETTLEMENT_PENDING_PHRASES:
        base = _COD_SETTLEMENT_PENDING_PHRASES[gateway_label]
        return {base, f"{base} not reflecting", f"{base} partially reflecting"}
    return {f"{gateway_label} Setlment Pending"}


def gateway_recon_by_period(reco_df, gateway_settlement_overall_df, gateway_configs):
    """
    Client's own "Gateway Recon Recoperiod" sheet (2026-08-30, items 5/6) -
    a period-scoped, per-gateway reconciliation table, one row per
    "Payment Provider" (Reco working's own refined, order-level gateway
    attribution - engine.attribution.attach_payment_columns):
        Payment Gateway | Group | Order Value | Gateway deduction |
        Refund | Pending Settlement | Part prepaid and part post paid |
        Order Value partially not received | Received in Bank | Check |
        Diff

    Every column except "Part prepaid and part post paid" (see below) is
    derived straight from Reco working (reco_df) plus
    gateway_settlement_overall() above, matching the client's own
    formulas column for column:
      - Order Value / Gateway deduction / Refund: that gateway's own
        total / total_deduction / refund_amount, summed.
      - Pending Settlement: total (order value), summed only for orders
        whose Query text is one of _pending_query_phrases() for this
        gateway - i.e. exactly the orders engine.reco.
        refine_queries_with_settlement_status() flagged as still sitting
        with this gateway/courier, not yet bank-credited.
      - Order Value partially not received: diff, summed for orders whose
        Receipt Status (engine.reco.attach_receipt_status) is "Partialy
        Received" or "Partially received and Refunded".
      - Received in Bank: this gateway's (Same Month + Next Month)
        Settlement Done from gateway_settlement_overall_df - i.e. money
        genuinely bank-matched for THIS reconciliation period (excludes
        Other Period/Not-Found, the same period-scoping already applied
        to the Executive Summary's Section 5 - see engine/formatting.py).
        This is a disclosed simplification of the client's own formula,
        which instead sums a separate line-item "Conso Receipt" sheet
        filtered to a "Reco period" flag; this engine doesn't build that
        sheet, but the two are the same figure by construction (an order
        counts as "Reco period" there under the same "This period" test
        that already splits Same Month/Next Month here).
      - Check: Order Value - Gateway deduction - Refund - Pending
        Settlement - Order Value partially not received (matches the
        client's own `=C-D-E-H-F` formula for every gateway that doesn't
        carry a "Part prepaid" figure - see below).
      - Diff: Received in Bank - Check - should be close to 0 when
        everything reconciles; a large Diff flags exactly the gateway
        where money isn't tying out.

    "Part prepaid and part post paid" is always 0 here - a disclosed gap.
    In the client's own reference workbook this is a hand-entered figure
    for the rare order that is genuinely split across two different
    payment rails (part COD, part prepaid) - this engine has no column
    today that marks an order as a split payment, so there's no rule to
    derive this from rather than guess at.

    A final "Total" row sums Order Value / Gateway deduction / Refund /
    Pending Settlement / Order Value partially not received / Received in
    Bank (never Check/Diff/Part-prepaid - matching the client's own Total
    row exactly). Orders with no resolvable "Payment Provider" at all get
    their own "Unresolved / #N/A" row (Order Value/deduction/refund still
    shown; Check/Diff left blank, since there's no gateway to reconcile
    them against) rather than being silently dropped.
    """
    if reco_df is None or reco_df.empty or "Payment Provider" not in reco_df.columns:
        return pd.DataFrame(columns=GATEWAY_RECON_RECOPERIOD_COLUMNS)

    df = reco_df.copy()
    gw_raw = df["Payment Provider"]
    # 2026-09-06 (round 16, order #30456): same COD-leg normalization as
    # gateway_settlement_overall() above - a COMBINED "<courier> COD,
    # <prepaid provider>" label groups under its COD leg here too, so this
    # sheet's own "Received in Bank" lookup (keyed off gateway_settlement_
    # overall_df's un-combined "Payment Gateway" labels below) actually
    # finds a match instead of a silently-orphaned all-zero row.
    gw = gw_raw.apply(cod_component_of_gateway_label).astype(str).str.strip()
    blank_mask = gw_raw.isna() | gw.isin(["", "nan", "None", "#N/A", "NA"])
    df["_gw"] = gw.where(~blank_mask, None)

    received_in_bank = pd.Series(dtype=float)
    if gateway_settlement_overall_df is not None and not gateway_settlement_overall_df.empty:
        g = gateway_settlement_overall_df.set_index("Payment Gateway")
        same = g["Settlement Done - This Period (Same Month)"] if "Settlement Done - This Period (Same Month)" in g.columns else pd.Series(dtype=float)
        nxt = g["Settlement Done - This Period (Next Month)"] if "Settlement Done - This Period (Next Month)" in g.columns else pd.Series(dtype=float)
        received_in_bank = same.add(nxt, fill_value=0.0)

    payment_mode_by_label = {
        cfg["label"]: str(cfg.get("payment_mode", "")).strip().lower() for cfg in (gateway_configs or [])
    }

    has_query = "query" in df.columns
    has_receipt_status = "receipt_status" in df.columns and "diff" in df.columns
    partial_labels = {"Partialy Received", "Partially received and Refunded"}

    rows = []
    for label in sorted(df["_gw"].dropna().unique()):
        sub = df[df["_gw"] == label]
        order_value = float(sub["total"].sum()) if "total" in sub.columns else 0.0
        deduction = float(sub["total_deduction"].sum()) if "total_deduction" in sub.columns else 0.0
        refund = float(sub["refund_amount"].sum()) if "refund_amount" in sub.columns else 0.0
        phrases = _pending_query_phrases(label)
        pending = float(sub.loc[sub["query"].isin(phrases), "total"].sum()) if has_query and "total" in sub.columns else 0.0
        partial_not_received = (
            float(sub.loc[sub["receipt_status"].isin(partial_labels), "diff"].sum()) if has_receipt_status else 0.0
        )
        rib = float(received_in_bank.get(label, 0.0))
        check = order_value - deduction - refund - pending - partial_not_received
        rows.append({
            "Payment Gateway": label,
            "Group": "COD" if payment_mode_by_label.get(label) == "cod" else "Prepaid",
            "Order Value": round(order_value, 2),
            "Gateway deduction": round(deduction, 2),
            "Refund": round(refund, 2),
            "Pending Settlement": round(pending, 2),
            "Part prepaid and part post paid": 0.0,
            "Order Value partially not received": round(partial_not_received, 2),
            "Received in Bank": round(rib, 2),
            "Check": round(check, 2),
            "Diff": round(rib - check, 2),
        })

    unresolved = df[df["_gw"].isna()]
    if len(unresolved):
        catch_all = "COD Delivered Amount not Received"
        pending_unresolved = (
            float(unresolved.loc[unresolved["query"] == catch_all, "total"].sum())
            if has_query and "total" in unresolved.columns else 0.0
        )
        rows.append({
            "Payment Gateway": "Unresolved / #N/A",
            "Group": "COD",
            "Order Value": round(float(unresolved["total"].sum()), 2) if "total" in unresolved.columns else 0.0,
            "Gateway deduction": round(float(unresolved["total_deduction"].sum()), 2) if "total_deduction" in unresolved.columns else 0.0,
            "Refund": round(float(unresolved["refund_amount"].sum()), 2) if "refund_amount" in unresolved.columns else 0.0,
            "Pending Settlement": round(pending_unresolved, 2),
            "Part prepaid and part post paid": 0.0,
            "Order Value partially not received": None,
            "Received in Bank": None,
            "Check": None,
            "Diff": None,
        })

    if not rows:
        return pd.DataFrame(columns=GATEWAY_RECON_RECOPERIOD_COLUMNS)

    out = pd.DataFrame(rows)[GATEWAY_RECON_RECOPERIOD_COLUMNS]
    total_row = {
        "Payment Gateway": "Total", "Group": None,
        "Order Value": round(float(out["Order Value"].sum()), 2),
        "Gateway deduction": round(float(out["Gateway deduction"].sum()), 2),
        "Refund": round(float(out["Refund"].sum()), 2),
        "Pending Settlement": round(float(out["Pending Settlement"].sum()), 2),
        "Part prepaid and part post paid": None,
        "Order Value partially not received": round(float(pd.to_numeric(out["Order Value partially not received"], errors="coerce").sum()), 2),
        "Received in Bank": round(float(pd.to_numeric(out["Received in Bank"], errors="coerce").sum()), 2),
        "Check": None, "Diff": None,
    }
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    return out
