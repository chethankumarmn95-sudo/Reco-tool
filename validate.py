import json
import sys
import pandas as pd

sys.path.insert(0, "/home/claude/reco_tool")
from engine.consolidator import build_consolidated_receipt, summarize_receipts_by_order
from engine.reco import run_shopify_pipeline
from engine.summary import headline_totals, month_summary

SRC = "/mnt/user-data/uploads/ESCA_D2C_Q1_Sales_Reco.xlsx"

with open("/home/claude/reco_tool/configs/esca_shopify.json") as f:
    config = json.load(f)

print("Loading source sheets...")
orders_df = pd.read_excel(SRC, sheet_name="Shopify Order report")
shiprocket_df = pd.read_excel(SRC, sheet_name="Shiprocket")
delhivery_df = pd.read_excel(SRC, sheet_name="Delhivery report")
gokwik_df = pd.read_excel(SRC, sheet_name="Gokwik")
razorpay_df = pd.read_excel(SRC, sheet_name="Razorpay")
delhivery_cod_df = pd.read_excel(SRC, sheet_name="Delhivery COD")
shiprocket_cod_df = pd.read_excel(SRC, sheet_name="ShiprocketAWB level report")

delivery_frames = {"Shiprocket": shiprocket_df, "Delhivery": delhivery_df}
gateway_frames = {
    "Gokwik": gokwik_df,
    "Razorpay": razorpay_df,
    "Delhivery COD": delhivery_cod_df,
    "Shiprocket COD": shiprocket_cod_df,
}

print("Building consolidated receipt (Layer 2)...")
consolidated = build_consolidated_receipt(gateway_frames, config["gateways"])
receipt_summary = summarize_receipts_by_order(consolidated)
print(f"  consolidated receipt rows: {len(consolidated)}")

print("Running Reco working pipeline (Layer 3)...")
reco_df = run_shopify_pipeline(orders_df, delivery_frames, receipt_summary, config)
print(f"  reco working rows (orders): {len(reco_df)}")

print("\n=== HEADLINE TOTALS (my engine) ===")
totals = headline_totals(reco_df)
for k, v in totals.items():
    print(f"  {k}: {v:,}")

print("\n=== YOUR ACTUAL RECO SUMMARY (for comparison) ===")
print("  Total orders: 14,049")
print("  Gross order value: 9,345,549.96")
print("  Receipt before deduction: 8,711,845.41")
print("  Total deduction: 792,477.29")
print("  Net settlement: 7,839,942.57")
print("  Total diff: 633,704.55")
print("  Open queries (orders): 292")

print("\n=== MONTH SUMMARY (my engine) ===")
print(month_summary(reco_df).to_string(index=False))

reco_df.to_csv("/home/claude/reco_tool/output_reco_working_v1.csv", index=False)
print("\nSaved full detail to output_reco_working_v1.csv for row-level inspection.")
