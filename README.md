# D2C Sales Reconciliation Tool

A website (running on your own laptop, for now) that replicates your Excel
reconciliation workbook: upload the source files, get the Reco working +
Reco Summary equivalent, automatically — now with a full SaaS-style
dashboard and sidebar navigation.

Validated against ESCA's actual Q1 FY26 Shopify data — order count, gross
order value, receipt before deduction, total deduction, and net settlement
all tie out to within ₹2 on a ₹9.3M book.

---

## How it's built (so you understand what you're running)

```
reco_tool/
├── app.py                   <- entrypoint: sets up the sidebar navigation shell
├── views/                   <- one file per page in the sidebar
│   ├── page_dashboard.py       (home page - KPIs, charts, health score)
│   ├── page_upload.py          (file uploads + validation)
│   ├── page_reconciliation.py  (Run button, runs the engine)
│   ├── page_orders_overview.py (order search + full detail table)
│   ├── page_rto_refunds.py     (RTO/refund breakdown)
│   ├── page_bank_linking.py    (bank UTR matching status)
│   ├── page_reports.py         (Excel download)
│   ├── page_data_management.py (save/load/delete saved months)
│   ├── page_exceptions.py      (combined exception list)
│   ├── page_settings.py        (client/channel switcher)
│   ├── page_activity_log.py    (saved-month history)
│   ├── charts.py               (plotly charts with safe fallback)
│   ├── theme.py                (shared CSS + small helpers)
│   └── state_init.py           (session state defaults)
├── engine/                  <- the actual reconciliation logic (unchanged
│   │                            by the UI restructure - same validated math)
│   ├── loaders.py           <- reads uploaded files, resolves column aliases
│   ├── consolidator.py      <- Layer 2: normalizes all gateways into one ledger
│   ├── reco.py               <- Layer 3: joins orders + delivery + receipts
│   ├── summary.py           <- Layer 4: month/status/query summaries
│   ├── lookup.py            <- Order Lookup Dashboard detail builder
│   ├── bank.py               <- bank statement UTR matching + UTR-wise reco
│   ├── period.py            <- this/previous/subsequent reconciliation period split
│   ├── settlement.py        <- Payment Gateway Settlement Report (settled/pending/deducted)
│   ├── storage.py           <- multi-month save/load/delete on disk
│   ├── validation.py        <- file validation + mandatory-source checks
│   └── formatting.py        <- Excel styling (headers, currency, dashboard sheet)
├── configs/
│   └── esca_shopify.json    <- column-name mapping for THIS client/channel
├── requirements.txt
└── validate.py               <- script used to check the engine against your real numbers
```

**Why the UI is split into views/:** each file is one sidebar page. Adding
a new page later means adding a new `page_*.py` file and one line in
`app.py`'s page list — the engine code never needs to change.

**Why configs are separate from code:** `configs/esca_shopify.json` is the
only file that knows ESCA's column names. To reconcile a different client,
or Amazon/Flipkart for ESCA, add a new config file with that source's
column mappings — it'll show up automatically in the Settings page's
channel switcher.

---

## One-time setup (do this once)

1. **Install Python** (skip if you already have it):
   Go to https://www.python.org/downloads/ → download → run the installer.
   On the first screen, tick **"Add Python to PATH"** before clicking Install.

2. **Get these files onto your laptop**: unzip the `reco_tool` folder
   anywhere, e.g. `Documents/reco_tool`.

3. **Open a terminal in that folder**:
   - Windows: open the `reco_tool` folder in File Explorer, type `cmd` in
     the address bar, press Enter.
   - Mac: right-click the folder → "New Terminal at Folder" (or open
     Terminal and type `cd ` then drag the folder in).

4. **Install the required packages** (one-time, needs internet):
   ```
   pip install -r requirements.txt
   ```
   If you update the tool later and see a "ModuleNotFoundError" for a
   package, re-run this command — it means a new dependency was added.

## Running the website (every time you want to use it)

In the same terminal:
```
streamlit run app.py
```
Your browser will open automatically to something like `http://localhost:8501`.
That's the tool — running privately on your machine, nobody else can see it.

To stop it, go back to the terminal and press `Ctrl+C`.

## Using it

1. You'll land on the **Dashboard** first. If nothing's been reconciled
   yet, it'll point you to Upload Data.
2. **Upload Data**: upload each source file in the format you already
   export (Shopify order report, Shiprocket, Delhivery, Unicommerce,
   Gokwik, Razorpay, Delhivery COD, ShiprocketAWB level report, and
   optionally your Bank Statement).
3. **Reconciliation**: click **Run reconciliation**.
4. Head back to **Dashboard** for the overview, or explore **Orders
   Overview**, **RTO & Refunds**, **Bank Statement / UTR Linking**,
   **Exceptions**, or **Reports** (Excel download) via the sidebar.
5. **Data Management**: save this month's results under a label so you
   can come back to it later without re-uploading, or load/combine
   previously saved months.

---

## Known gaps in this version — to fix together next

- **Amazon and Flipkart configs don't exist yet** - the engine and the
  channel switcher are both ready for them, but I need a sample export of
  each to build accurate column mappings.
- **Activity Log is a simple first version** - it currently shows saved
  reconciliations only, not a full action-by-action audit trail.
- **Running locally only** - deploying this to a shareable link (so
  Mahanthy or ESCA's directors can use it without installing anything) is
  a small next step, not a rebuild.
- **UTR-level bank reconciliation "Remarks" is a best-effort automated
  read, not a final answer.** It uses clear, disclosed rules (see
  `engine/bank.py`'s docstring) to approximate Matched / Settled-with-
  other-period / Bank-statement-not-found - but this is a judgment call
  in real bookkeeping, and should always get a human review pass before
  being relied on, same as any reconciliation exception list.
- **Gateway attribution for "Pending for Settlement" is a best-effort
  guess** for orders that haven't settled yet (see `engine/settlement.py`)
  - based on Shopify's Payment Method text plus which courier delivered
  a COD order. An order that can't be matched to a specific gateway shows
  up as "Unattributed COD"/"Unattributed Prepaid" rather than being
  silently dropped, but still needs a human to confirm which gateway it
  really belongs to.
- **The six reconciliation categories (see below) are decided by
  configurable timing assumptions, not confirmed SLAs.** `expected_settlement_days`
  per gateway in `configs/esca_shopify.json`, and the "how long is normal
  before this counts as an Exception" grace periods in `engine/bank.py`,
  are reasonable starting defaults - not a verified fact about any specific
  gateway's or courier's actual contract. Please confirm real settlement
  terms with each gateway/courier and adjust the config if they differ.
- **The COD settlement-batch match (amount + date, not a hard reference) is
  a heuristic, same as the existing UTR-level Remarks column** - it's a
  strong lead for the reconciling accountant, not a substitute for tracing
  every rupee by reference number. See `engine/bank.py`'s docstring.

## Recent updates (this round)

- **Post-delivery self-review fixes (four issues caught by a follow-up code
  review before this was called done, all now fixed and covered by new
  regression checks in `test_reconciliation_categories.py`):**
  1. Orders with no settlement row yet (COD not-yet-remitted, Prepaid not-
     yet-settled) were showing **Settlement Amount = ₹0** in the Settlement
     Pending Report, because that figure was derived from the (also zero)
     gateway receipt amount instead of the order's own outstanding Total -
     understating exactly the orders that most need chasing. Now uses the
     order Total for those rows; unchanged (net, post-deduction) for orders
     that do have a settlement row.
  2. `classify_order_bank_status()` could crash (`TypeError: Cannot cast
     DatetimeArray to dtype float64`) on a perfectly normal partial run -
     reconciling before any gateway/COD file has been uploaded yet - due to
     a pandas edge case mapping an empty, typed lookup Series. Fixed by
     using plain dict lookups instead.
  3. The COD settlement-batch amount/date fallback match could double-count
     a single bank credit - once via a direct order-level UTR match, and
     again via an unrelated batch's amount/date match against the same
     bank row. Both `classify_order_bank_status()` and the Bank Linking
     page's own batch display now exclude bank credits already claimed by
     a direct UTR match (`matched_order_level_utrs()`) before searching.
  4. Bank-narration UTR extraction didn't handle slash-delimited UPI/IMPS
     formats (e.g. "UPI/402912345678/username/YESB0000/paytm") - it either
     missed the reference or picked the wrong token. Separator handling now
     includes "/", and the fallback path prefers a long all-digit token
     (typical UPI RRN style) over a mixed alphanumeric one.
- **Order Lookup / Bank Reconciliation now use six clear reconciliation
  categories instead of the old, sometimes-misleading three-state view.**
  Previously, any order with no gateway receipt row at all - including a
  COD order that was RTO'd, cancelled, still in transit, or never got a
  usable delivery status - showed "Not checked - no bank statement
  uploaded", even when a bank statement HAD been uploaded (that text really
  meant "nothing to check for this order", not "no bank statement"). And
  some COD orders that a courier had genuinely already remitted, and that
  had already landed in the bank, still showed "Bank receipt not
  identified" - because COD couriers usually repeat one lump bank credit's
  reference imperfectly (or not at all) on every individual order row of
  their settlement export, so matching order-by-order missed them.
  `engine/bank.py`'s new `classify_order_bank_status()` now distinguishes:
  COD - Not Delivered (No Receipt Expected); COD - Delivered & Settlement
  Pending; COD - Settlement Received & Bank Matched; Prepaid - Payment
  Received & Settlement Pending; Prepaid - Settlement Received & Bank
  Matched; and Exception / Manual Reconciliation Required for anything
  that's gone past its expected window without a match. It also adds a
  settlement-BATCH fallback match (group a courier's same-day settlement
  rows, match the batch total to a bank credit by amount + date, not just
  by UTR) and broadens bank-narration UTR extraction beyond NEFT-only, so
  those "already credited but not identified" cases get traced. See the
  Order Lookup's new "Reconciliation Category" / "Category Reason" columns,
  the Bank Statement / UTR Linking page's new breakdown, and
  `engine/bank.py`'s module-section docstring for the full story.
- **New Settlement Pending Report**, gateway/partner-wise: a dedicated
  order-wise detail table (Order ID, Payment Gateway, UTR/reference, Order
  Date, Payment/Receipt Date, Order/Gateway/Settlement Amount, Expected/
  Actual Settlement Date, Bank Credit Date, Settlement Status, Days
  Pending) plus a summary split into "Payment Gateway Settlement Pending"
  (Razorpay, Gokwik, ...) and "COD Settlement Pending" (Shiprocket COD,
  Delhivery COD, ...) - on the Reconciliation page, the Reports page, and
  as new sheets in the downloaded workbook. See `engine/settlement_pending.py`.
- **Dashboard now shows a gateway-wise / delivery-partner-wise
  reconciliation health table** - how much is Received in Bank vs Pending
  Settlement vs Pending Bank Matching vs a genuine Exception, per gateway/
  courier, so you can see at a glance where money is actually stuck.
- **A synthetic (no real order data) regression test,
  `test_reconciliation_categories.py`**, exercises all six categories, the
  settlement-batch fallback match, and the broadened UTR extraction end to
  end - run `python3 test_reconciliation_categories.py` any time after
  changing this logic to check nothing regressed.
- **Payment Gateway Settlement Report**: the Reconciliation page, the
  Reports page (and its downloaded workbook, new "Gateway Settlement"
  sheet) now show Settlement Done / Pending for Settlement / Deductions
  separately for each payment gateway (Gokwik, Razorpay, Delhivery COD,
  Shiprocket COD, ...) instead of only as one lump total. See
  `engine/settlement.py`.
- **Bank reconciliation now matches how the sample workbook actually does
  it**: one bank credit (UTR) very often bundles several orders' payouts
  together, sometimes from more than one reconciliation period (e.g. a
  delayed COD remittance for last month's orders landing in this month's
  bank credit). The tool now splits each UTR's total into this-period vs
  other-period amounts, compares it to the real bank credit, and flags
  the difference - see the new "Bank Reconciliation by Settlement (UTR)"
  section on the Bank Statement / UTR Linking page, the Reports page, and
  `engine/bank.py` / `engine/period.py`.
- **Fixed the #1 known gap from last round**: bank statement UTR matching
  no longer needs a dedicated UTR column in your export - it's now
  extracted straight from the Narration text (validated against your
  actual bank statement: NEFT credits format as "NEFT &lt;reference&gt;
  &lt;remitter name&gt;..."), matching **310 of 312** UTRs your own sample
  workbook had captured by hand, and the underlying settlement figures
  tie out **exactly** (₹0 difference) against your sample's own UTR-wise
  summary table on all 320 shared UTRs.
- **Fixed a real crash**: `ModuleNotFoundError: No module named 'plotly'`
  happened because a new dependency wasn't installed. Fixed the immediate
  cause, and made charts fall back to Streamlit's built-in charts if
  plotly is ever missing again, instead of crashing the page.
- **Full dashboard-first restructure**: Dashboard is now the home page
  (not the upload form), with a real left sidebar covering all 11
  requested sections. The reconciliation engine itself didn't change -
  only how the UI is organized - so all previously validated numbers
  still match exactly.
- **Multi-channel groundwork**: Settings page lists every available
  client/channel config as a switchable option, ready for Amazon/Flipkart
  configs to be added later.

## Adding a new client or channel

Copy `configs/esca_shopify.json`, rename it, and change the values to
match the new client/channel's column names. It'll appear automatically
in the Settings page's channel list — no Python knowledge needed, just
careful copy-editing of column names.
