"""
Synthetic (no real client data) smoke test for the new reconciliation
categorization + settlement-batch matching + settlement pending report.
Exercises every one of the 6 categories, the COD batch-matching fallback,
and the broadened bank-narration UTR extraction, using made-up orders that
mirror the ESCA Shopify config's shape but contain no real client data.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import json
import pandas as pd

from engine.reco import run_shopify_pipeline
from engine.consolidator import build_consolidated_receipt, summarize_receipts_by_order, receipt_detail_by_order
from engine.lookup import build_sku_detail, build_order_lookup
from engine.bank import (
    load_bank_statement, classify_order_bank_status, extract_utr_from_narration,
    build_cod_settlement_batches, match_batches_to_bank, matched_order_level_utrs,
    COD_NOT_DELIVERED, COD_SETTLEMENT_PENDING, COD_BANK_MATCHED,
    PREPAID_SETTLEMENT_PENDING, PREPAID_BANK_MATCHED, EXCEPTION_MANUAL_REVIEW,
    PREPAID_PAYMENT_NOT_RECEIVED,
)
from engine.settlement_pending import build_settlement_pending_report, settlement_pending_summary_by_gateway, reconciliation_health_by_gateway

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "configs", "esca_shopify.json")) as f:
    config = json.load(f)

TODAY = pd.Timestamp("2026-08-18")  # matches the session's "today" - fixed so test is deterministic

# --- Orders (Shopify-shaped) -------------------------------------------------
orders_rows = [
    # order_id, financial_status, fulfillment_status, total, payment_method, created_at
    ("1001", "pending", "unfulfilled", 500, "Cash on Delivery (COD)", "2026-07-01"),   # COD, RTO -> not delivered, no receipt expected
    ("1002", "pending", "unfulfilled", 600, "Cash on Delivery (COD)", "2026-07-02"),   # COD, delivered, no settlement row yet -> pending
    ("1003", "paid", "fulfilled", 700, "Cash on Delivery (COD)", "2026-06-01"),        # COD, delivered, settlement row w/ matching UTR -> bank matched
    ("1004", "paid", "fulfilled", 800, "Cash on Delivery (COD)", "2026-06-02"),        # COD, delivered, settlement row, blank UTR, but batch matches by amount/date
    ("1005", "paid", "fulfilled", 900, "Cash on Delivery (COD)", "2026-05-01"),        # COD, delivered, settlement row long ago, no bank match anywhere -> exception
    ("1006", "paid", "fulfilled", 1000, "Razorpay", "2026-08-01"),                     # Prepaid, paid, no settlement row yet -> pending
    ("1007", "paid", "fulfilled", 1100, "Razorpay", "2026-06-05"),                     # Prepaid, settlement row w/ matching UTR -> bank matched
    ("1008", "pending", "unfulfilled", 1200, "Razorpay", "2026-08-10"),                # Prepaid, not yet paid, no settlement -> "payment not received"
    ("1009", "paid", "fulfilled", 1300, "Razorpay", "2026-05-01"),                     # Prepaid, settlement long ago, no bank match -> exception
]
orders_records = []
for oid, fin, fulfil, total, pm, created in orders_rows:
    orders_records.append({
        "Name": oid, "Created at": created, "Financial Status": fin,
        "Fulfillment Status": fulfil, "Subtotal": total, "Shipping": 0, "Taxes": 0,
        "Total": total, "Payment Method": pm,
    })
orders_df = pd.DataFrame(orders_records)

# --- Delivery partner (Shiprocket) ------------------------------------------
delivery_rows = [
    ("1001", "RTO DELIVERED", "2026-07-05"),
    ("1002", "DELIVERED", "2026-07-03"),
    ("1003", "DELIVERED", "2026-06-03"),
    ("1004", "DELIVERED", "2026-06-04"),
    ("1005", "DELIVERED", "2026-05-03"),
    ("1006", "DELIVERED", "2026-08-03"),
    ("1007", "DELIVERED", "2026-06-07"),
    ("1008", "DELIVERED", "2026-08-12"),
    ("1009", "DELIVERED", "2026-05-03"),
]
shiprocket_df = pd.DataFrame([
    {"Order ID": oid, "Status": status, "Order Delivered Date": ddate, "RTO Delivered Date": ddate if "RTO" in status else None}
    for oid, status, ddate in delivery_rows
])
delivery_frames = {"Shiprocket": shiprocket_df}

# --- Gateways: Shiprocket COD (1003 matched UTR, 1004 blank UTR -> batch, 1005 no match) ---
shiprocket_cod_rows = [
    {"Order Id": "1003", "COD Available - Line Allocation": 700, "Freight Charges - Line Allocation": 20,
     "Remittance Date": "2026-06-10", "UTR": "AXISP00111111"},
    {"Order Id": "1004", "COD Available - Line Allocation": 800, "Freight Charges - Line Allocation": 20,
     "Remittance Date": "2026-06-11", "UTR": ""},  # blank UTR, distinct remittance date - must be caught by batch matching
    {"Order Id": "1005", "COD Available - Line Allocation": 900, "Freight Charges - Line Allocation": 20,
     "Remittance Date": "2026-05-10", "UTR": "AXISP00999999"},  # has a UTR, but it never shows up in the bank statement -> exception
]
shiprocket_cod_df = pd.DataFrame(shiprocket_cod_rows)

razorpay_rows = [
    {"Shopify Order ID": "1007", "amount": 1100, "fee (exclusive tax)": 20, "tax": 2,
     "transaction_entity": "payment", "settled_at": "2026-06-08", "settlement_utr": "RATNP00222222"},
    {"Shopify Order ID": "1009", "amount": 1300, "fee (exclusive tax)": 20, "tax": 2,
     "transaction_entity": "payment", "settled_at": "2026-05-08", "settlement_utr": "RATNP00888888"},  # never in bank statement -> exception
]
razorpay_df = pd.DataFrame(razorpay_rows)

gateway_frames = {"Shiprocket COD": shiprocket_cod_df, "Razorpay": razorpay_df}

# --- Bank statement: one NEFT credit matching 1003's UTR directly, one lump
# credit matching the 1004 batch's total (780 = 800-20) by amount+date, one
# broadened-format (RTGS) credit matching Razorpay 1007's UTR. Nothing for
# 1005 or 1009 - they should end up as exceptions. ---
bank_rows = [
    {"Date": "2026-06-12", "Amount": 680.0, "Narration": "NEFT AXISP00111111 SHIPROCKET PVT LTD"},  # matches 1003 (700-20=680)
    {"Date": "2026-06-13", "Amount": 780.0, "Narration": "NEFT AXISBATCHXXXX SHIPROCKET PVT LTD SETTLEMENT"},  # batch total for 1004 (800-20=780), no UTR match but amount+date match (settlement date 06-11, within 5-day window)
    {"Date": "2026-06-09", "Amount": 1078.0, "Narration": "RTGS RATNP00222222 RAZORPAY SOFTWARE"},  # matches 1007 (1100-22=1078), broadened non-NEFT prefix
]
bank_df = pd.DataFrame(bank_rows)

# --- Run the pipeline (mirrors page_reconciliation.py) ----------------------
consolidated = build_consolidated_receipt(gateway_frames, config["gateways"])
receipt_summary = summarize_receipts_by_order(consolidated)
reco_df = run_shopify_pipeline(orders_df, delivery_frames, receipt_summary, config)

bank_ledger = load_bank_statement(bank_df, config["bank_statement"])
recon_status_df = classify_order_bank_status(
    reco_df, consolidated, bank_ledger, config["gateways"], bank_statement_uploaded=True, as_of_date=TODAY,
)

result = dict(zip(recon_status_df["order_id"], recon_status_df["Reconciliation Category"]))
print("=== Per-order Reconciliation Category ===")
for oid in [r[0] for r in orders_rows]:
    print(f"  {oid}: {result[oid]}")

expected = {
    "1001": COD_NOT_DELIVERED,
    "1002": COD_SETTLEMENT_PENDING,
    "1003": COD_BANK_MATCHED,
    "1004": COD_BANK_MATCHED,      # via settlement-batch amount/date fallback
    "1005": EXCEPTION_MANUAL_REVIEW,
    "1006": PREPAID_SETTLEMENT_PENDING,
    "1007": PREPAID_BANK_MATCHED,  # via broadened RTGS narration extraction
    "1008": PREPAID_PAYMENT_NOT_RECEIVED,
    "1009": EXCEPTION_MANUAL_REVIEW,
}

failures = [oid for oid, cat in expected.items() if result.get(oid) != cat]
print()
if failures:
    print("FAILURES:")
    for oid in failures:
        print(f"  {oid}: expected {expected[oid]!r}, got {result.get(oid)!r}")
    sys.exit(1)
print("ALL 9 CATEGORY ASSERTIONS PASSED")

# --- Narration extraction unit checks ---------------------------------------
narration_cases = [
    ("NEFT AXISP00785504798 DELHIVERY  LIMITED UTIB0000...", "AXISP00785504798"),  # original validated NEFT case, must still work
    ("RTGS RATNP00222222 RAZORPAY SOFTWARE", "RATNP00222222"),
    ("UPI-CR/12345678/SOMENAME/OK", "CR/12345678/SOMENAME/OK".upper() if False else None),  # not asserted, just illustrative
]
assert extract_utr_from_narration("NEFT AXISP00785504798 DELHIVERY  LIMITED UTIB0000...") == "AXISP00785504798"
assert extract_utr_from_narration("RTGS RATNP00222222 RAZORPAY SOFTWARE") == "RATNP00222222"
assert extract_utr_from_narration("AD SPEND DEBIT FACEBOOK ADS") is None  # no digit+letter 8+ char token -> correctly no match
print("NARRATION EXTRACTION CHECKS PASSED")

# --- Settlement Pending Report + summaries -----------------------------------
sku_detail_df = pd.DataFrame(columns=["order_id", "skus", "quantity"])
receipt_detail_df = receipt_detail_by_order(consolidated)
lookup_df = build_order_lookup(reco_df, sku_detail_df, receipt_detail_df, recon_status_df)
assert "Reconciliation Category" in lookup_df.columns
assert (lookup_df.loc[lookup_df["Order ID"] == "1001", "Bank status"] == COD_NOT_DELIVERED).all()
assert "Not checked - no bank statement uploaded" not in lookup_df["Bank status"].values
print("ORDER LOOKUP CHECKS PASSED (no misleading 'Not checked' fallback anywhere)")

pending_df = build_settlement_pending_report(reco_df, recon_status_df, receipt_detail_df, config["gateways"], receipt_summary_df=receipt_summary)
pending_ids = set(pending_df["Order ID"])
assert pending_ids == {"1002", "1005", "1006", "1009"}, pending_ids  # exactly the pending/exception ones, not the matched/not-delivered ones
print("SETTLEMENT PENDING REPORT CHECKS PASSED:", sorted(pending_ids))

# 1005: gross COD Available = 900, Freight deduction = 20 -> net Settlement Amount should be 880, not 900
row_1005 = pending_df[pending_df["Order ID"] == "1005"].iloc[0]
assert row_1005["Gateway Amount"] == 900.0, row_1005["Gateway Amount"]
assert row_1005["Settlement Amount"] == 880.0, row_1005["Settlement Amount"]
print("NET SETTLEMENT AMOUNT CHECK PASSED (Gateway Amount 900 vs net Settlement Amount 880)")

summary_df = settlement_pending_summary_by_gateway(pending_df, config["gateways"])
assert set(summary_df["Group"]) <= {"COD", "Prepaid"}
print(summary_df.to_string(index=False))

health_df = reconciliation_health_by_gateway(reco_df, recon_status_df, receipt_detail_df, config["gateways"])
print()
print(health_df.to_string(index=False))

# --- Regression: Settlement Amount for orders with NO settlement row yet ---
# (bug found in post-delivery review: these previously showed Settlement
# Amount = 0, derived from receipt_amount which is 0 when nothing has been
# collected/settled yet, instead of the order's own outstanding Total.)
row_1002 = pending_df[pending_df["Order ID"] == "1002"].iloc[0]  # COD, no settlement row
assert row_1002["Gateway Amount"] == 0.0, row_1002["Gateway Amount"]
assert row_1002["Settlement Amount"] == 600.0, row_1002["Settlement Amount"]  # order Total, not 0
row_1006 = pending_df[pending_df["Order ID"] == "1006"].iloc[0]  # Prepaid, no settlement row
assert row_1006["Gateway Amount"] == 0.0, row_1006["Gateway Amount"]
assert row_1006["Settlement Amount"] == 1000.0, row_1006["Settlement Amount"]  # order Total, not 0
print("SETTLEMENT AMOUNT FOR NO-SETTLEMENT-ROW ORDERS CHECK PASSED (order Total, not 0)")

# --- Regression: empty consolidated_df must not crash (pandas 3.0.2 empty
# typed-Series .map() bug found in post-delivery review) - this is the
# normal "no gateway/COD file uploaded yet" partial-run state, not an edge
# case, so it must classify cleanly rather than raising. ---
empty_consolidated = build_consolidated_receipt({}, config["gateways"])
assert empty_consolidated.empty
recon_status_empty = classify_order_bank_status(
    reco_df, empty_consolidated, bank_ledger, config["gateways"], bank_statement_uploaded=True, as_of_date=TODAY,
)
assert len(recon_status_empty) == len(reco_df)
print("EMPTY CONSOLIDATED_DF NO-CRASH CHECK PASSED")

# --- Regression: settlement-batch matching must not double-count a bank
# credit already claimed by a direct order-level UTR match (bug found in
# post-delivery review). Build a batch scenario where a resolved order
# (1003, direct UTR match) shares its settlement date with another COD
# order whose batch, if 1003's amount were wrongly included, would total to
# an amount matching a DIFFERENT bank credit than the real one. Simplest
# direct check: the excluded UTR set must contain 1003's own UTR, and the
# COD batches computed with that exclusion must not include order 1003. ---
excluded = matched_order_level_utrs(consolidated, bank_ledger)
assert "AXISP00111111" in excluded, excluded
batches_excl = build_cod_settlement_batches(consolidated, config["gateways"], bank_ledger)
all_batched_orders = set()
for ids in batches_excl["order_ids"]:
    all_batched_orders.update(ids)
assert "1003" not in all_batched_orders, all_batched_orders  # already resolved by direct UTR match - must not reappear in a batch
print("SETTLEMENT-BATCH DOUBLE-COUNT EXCLUSION CHECK PASSED")

# --- Regression: UPI/IMPS slash-delimited narration extraction (bug found
# in post-delivery review: the old separator class didn't include "/", so
# slash-delimited UPI narrations either extracted nothing or the wrong
# token). ---
assert extract_utr_from_narration("UPI/402912345678/username/YESB0000/paytm") == "402912345678"
assert extract_utr_from_narration("UPI-CR/12345678/SOMENAME/OK") == "12345678"
# No recognised prefix at all, pure-digit RRN-style reference embedded in
# free text - must prefer the long all-digit token over any shorter mixed
# alphanumeric fragment elsewhere in the narration.
assert extract_utr_from_narration("SETTLEMENT REF 501234567890 FROM SOMEBANK0001") == "501234567890"
print("UPI/IMPS SLASH-DELIMITED NARRATION EXTRACTION CHECK PASSED")

# --- Regression: timezone-aware source dates must not crash classification
# (real bug reported after delivery: "TypeError: Cannot subtract tz-naive
# and tz-aware datetime-like objects"). Shopify's own "Created at" export
# commonly includes a UTC offset (e.g. "2026-07-01 10:23:45 +0530") - the
# synthetic fixtures above happened to use plain naive dates, which is why
# this only surfaced against a real export. Build a minimal tz-aware
# fixture mirroring that exact shape and confirm it classifies cleanly. ---
tz_orders_df = pd.DataFrame([{
    "Name": "2001", "Created at": "2026-07-01 10:23:45 +0530", "Financial Status": "paid",
    "Fulfillment Status": "fulfilled", "Subtotal": 500, "Shipping": 0, "Taxes": 0,
    "Total": 500, "Payment Method": "Cash on Delivery (COD)",
}])
tz_shiprocket_df = pd.DataFrame([{
    "Order ID": "2001", "Status": "DELIVERED", "Order Delivered Date": "2026-07-03 09:00:00 +0530",
    "RTO Delivered Date": None,
}])
tz_cod_df = pd.DataFrame([{
    "Order Id": "2001", "COD Available - Line Allocation": 500, "Freight Charges - Line Allocation": 10,
    "Remittance Date": "2026-07-10 00:00:00 +0530", "UTR": "AXISP00TZTEST01",
}])
tz_bank_df = pd.DataFrame([{
    "Date": "2026-07-12", "Amount": 490.0, "Narration": "NEFT AXISP00TZTEST01 SHIPROCKET PVT LTD",
}])
tz_consolidated = build_consolidated_receipt({"Shiprocket COD": tz_cod_df}, config["gateways"])
tz_receipt_summary = summarize_receipts_by_order(tz_consolidated)
tz_reco_df = run_shopify_pipeline(tz_orders_df, {"Shiprocket": tz_shiprocket_df}, tz_receipt_summary, config)
tz_bank_ledger = load_bank_statement(tz_bank_df, config["bank_statement"])
tz_recon_status_df = classify_order_bank_status(
    tz_reco_df, tz_consolidated, tz_bank_ledger, config["gateways"], bank_statement_uploaded=True, as_of_date=TODAY,
)
tz_result = dict(zip(tz_recon_status_df["order_id"], tz_recon_status_df["Reconciliation Category"]))
assert tz_result["2001"] == COD_BANK_MATCHED, tz_result
print("TIMEZONE-AWARE SOURCE DATES CHECK PASSED (no crash, classified correctly)")

print("\nALL SYNTHETIC TESTS PASSED")
