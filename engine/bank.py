"""
bank.py
-------
The final layer of your reconciliation pipeline: matching the bank
statement to the gateway receipts, via UTR - exactly like your existing
Excel working links bank credits to reconciliation entries.

Three levels of matching, the first two mirroring the two tables in the
client's sample Bank_reco.xlsx workbook, the third added on top of them:

  1. Order/transaction level (load_bank_statement, match_consolidated_to_bank) -
     has this gateway transaction's own UTR been traced to a bank credit at all.

  2. UTR / settlement-batch level (build_settlement_ledger,
     bank_reconciliation_by_utr) - one bank credit (one UTR) very often
     bundles several orders' settlements together, and not always from
     the same reconciliation period (e.g. a delayed COD remittance for
     last month's orders lands in this month's bank credit alongside this
     month's own orders). This splits each UTR's total into "this period"
     vs "other period", compares that combined total to what the bank
     statement actually shows for that UTR, and flags the difference -
     exactly the "Amount during [period]" / "Setled belong to other
     month" / "Total" / "Bank" / "Defference" / "Remarks" columns in the
     sample workbook.

  3. Per-order reconciliation category (classify_order_bank_status, plus
     its COD settlement-batch fallback build_cod_settlement_batches /
     match_batches_to_bank) - the newer, client-confirmed layer that
     decides, for every single order, one of six statuses: never-delivered
     COD (no receipt expected), COD/Prepaid settlement pending, COD/Prepaid
     bank-matched, or a genuine exception needing manual reconciliation.
     This is what feeds the Order Lookup "Reconciliation Category" column
     and the Settlement Pending Report (engine/settlement_pending.py) - see
     the dedicated docstring above classify_order_bank_status() below for
     the full story of why this layer exists.

IMPORTANT - manual review still required throughout this module: the
Remarks classification in bank_reconciliation_by_utr() is a best-effort,
disclosed-rules read of the same categories used in the sample workbook
(Matched / Settled with other-period transaction(s) / Bank statement not
found) - the sample workbook's own Remarks column was itself partly a
manual, judgment-based entry (not driven by one formula throughout - some
rows tied out numerically but were still flagged for follow-up). The
per-order Reconciliation Category and the settlement-batch amount/date
matching below are the same kind of disclosed heuristic, not a hard
reference-number match. Treat every column in this module as a strong
starting point for the reconciling accountant's review, not a final answer.
"""

import re
import pandas as pd

from .loaders import resolve_col_or_raise, resolve_col
from .period import THIS_PERIOD, PREVIOUS_PERIOD, SUBSEQUENT_PERIOD, NOT_FOUND

# Rupee 1 tolerance - matches the rounding tolerance already used elsewhere
# in this engine (see engine/reco.py flag_queries' diff > 1 / diff < -1 checks).
TOLERANCE = 1.0

# RBI/NPCI-standardised UTR/reference lengths, per settlement rail - a
# fixed national format, not a per-bank convention: NEFT, RTGS, and IMFT
# references are 16 alphanumeric characters; IMPS/UPI RRNs are 12 digits.
# Used by extract_utr_from_narration() below to cap a rail-prefixed
# capture at its rail's true length, so a narration with no separator
# between the reference and the remitter's name (see that function's
# docstring) doesn't overshoot into the name.
RAIL_UTR_LENGTH = {"NEFT": 16, "RTGS": 16, "INFT": 16, "IMPS": 12, "UPI": 12}

# Client-reported (2026-08-21): a payment gateway's OWN "UTR" field for a
# settlement sometimes uses a completely different reference format from
# what actually prints in the bank statement's narration for that same
# credit - e.g. ICICI: the gateway recorded "ICICN22025102703875831" while
# the bank narration extracts to "ICIN230003875831". These aren't the same
# string at all, but they share the SAME trailing 8 characters
# ("03875831") - the bank's own running-serial suffix for that reference,
# which stays constant even though the prefix/date encoding wrapped around
# it differs between the two systems. UTR_SUFFIX_MATCH_LEN is the fallback
# match length used by match_utr_against_bank() below when an exact
# normalized match fails - long enough that two genuinely different
# transactions sharing it by coincidence is vanishingly unlikely (this is a
# bank-generated serial, not a small enumerable code), short enough to
# survive the prefix/date formatting differences seen in the client's
# example.
UTR_SUFFIX_MATCH_LEN = 8


def to_naive_timestamp(value):
    """
    Strip timezone info from a parsed Timestamp (scalar), if present, so
    the ageing/grace-period date arithmetic in this module never tries to
    subtract a tz-aware Timestamp from a tz-naive one (or vice versa) -
    pandas raises a hard TypeError ("Cannot subtract tz-naive and
    tz-aware datetime-like objects") the moment that happens, rather than
    silently getting it wrong. This module only ever needs whole calendar
    days, not time-of-day/timezone precision, so dropping the source's
    offset here is safe.

    Real-world trigger: Shopify's "Created at" / fulfillment timestamps
    are commonly exported WITH a timezone offset (e.g. "2026-07-01
    10:23:45 +0530"), which pandas keeps as a tz-aware Timestamp once
    parsed - while `pd.Timestamp.now()` (the default `as_of_date`) and
    most courier/bank date columns are naive. The synthetic test fixtures
    used to build this module's logic happened to use plain naive dates,
    so this only surfaced once run against a real, timezone-stamped
    export.
    """
    ts = value if isinstance(value, pd.Timestamp) else pd.Timestamp(value)
    if pd.notna(ts) and ts.tzinfo is not None:
        ts = ts.tz_localize(None)
    # 2026-08-31 fix (client-reported: "TypeError: Invalid comparison
    # between dtype=datetime64[us] and Timestamp", raised the moment a
    # report was generated on the client's own machine). Some pandas
    # builds parse dates - from a legacy .xls/OLE workbook, or from a
    # freshly re-combined multi-month bank ledger - at microsecond
    # resolution (datetime64[us]) rather than the nanosecond resolution
    # (datetime64[ns]) pandas has historically always defaulted to.
    # Comparing a datetime64[us] Series directly against a bare
    # pd.Timestamp scalar can then raise, because the two sides disagree
    # on resolution, not because either value is actually invalid. Never
    # reproduced against this engine's own synthetic test fixtures (which
    # happened to parse at the historical 'ns' resolution) - only against
    # the client's real environment. Forcing every Timestamp this
    # function returns to nanosecond resolution - the one resolution
    # every pandas version compares cleanly against - sidesteps the
    # mismatch regardless of which resolution the OTHER side (a
    # to_naive_datetime_series()-produced column, below) happens to
    # already be in.
    if pd.notna(ts) and hasattr(ts, "as_unit"):
        ts = ts.as_unit("ns")
    return ts


def to_naive_datetime_series(series):
    """Series/column version of to_naive_timestamp() above - for date
    columns parsed in bulk (e.g. the bank statement's own date column).
    Once any single date column in a comparison chain is tz-aware while
    another is naive, every later subtraction against it crashes - so
    dates are normalized to naive right where they're first parsed,
    rather than patched at every place they're later compared.

    dayfirst=True: this module only ever parses Indian bank statement date
    columns. When the source is a real Excel workbook, dates already arrive
    as proper datetime objects and dayfirst is irrelevant - but a CSV
    export (verified against the client's real Kotak CSV export) carries
    dates as plain DD-MM-YYYY text, and pandas defaults to a US-style
    month-first read of ambiguous strings. Without dayfirst=True, any date
    with day <= 12 (roughly half of every month) silently swaps day and
    month - e.g. "03-07-2025" (3 July) read back as 3 March - which doesn't
    error, it just quietly shifts ~half the bank statement's dates by
    months, breaking the settlement-batch amount+date matching for exactly
    those rows. Caught by this engine's own settlement register showing a
    dramatically worse match rate from the CSV bank export than the xlsx
    export of the SAME underlying bank account for the SAME period.
    """
    s = pd.to_datetime(series, errors="coerce", dayfirst=True)
    try:
        if s.dt.tz is not None:
            s = s.dt.tz_localize(None)
    except (AttributeError, TypeError):
        pass
    # 2026-08-31 fix - see to_naive_timestamp()'s own docstring above for
    # the full "Invalid comparison between dtype=datetime64[us] and
    # Timestamp" story (client-reported). Force this column to
    # nanosecond resolution too, so it always compares cleanly against a
    # to_naive_timestamp()-produced scalar regardless of which resolution
    # the source data happened to parse at.
    try:
        s = s.astype("datetime64[ns]")
    except (TypeError, ValueError):
        pass
    return s


def normalize_utr(series):
    """UTRs are alphanumeric references - just need consistent case/whitespace
    for matching, no digit-only stripping like order IDs."""
    return series.astype(str).str.strip().str.upper()


def extract_utr_from_narration(narration, require_rail_prefix=False):
    """
    Most real Indian bank statement exports don't have a dedicated "UTR"
    column at all - the reference is embedded in the free-text Narration,
    in whatever format that bank uses. This was the #1 known gap flagged
    in this tool's README ("Bank statement column mapping is my best
    guess, not yet validated against your real bank export").

    Validated against the client's actual bank statement export: every
    NEFT gateway settlement credit in it is formatted
    "NEFT <REFERENCE> <REMITTER NAME>..." - e.g.
    "NEFT AXISP00785504798 DELHIVERY  LIMITED UTIB0000...". The token
    right after "NEFT" is exactly the UTR/reference the gateway files
    quote (AXISP.../AXISCN.../IN226... etc.).

    Broadened from the original NEFT-only version, which silently dropped
    any settlement credit that arrived via a different rail (RTGS/IMPS/UPI)
    or with the bank's own prefix punctuation ("NEFT-CR", "NEFT IN ...")
    from the bank ledger entirely - a real, already-credited settlement then
    looked "not identified" no matter what the gateway side checked, simply
    because it was never even loaded into the candidate UTR list. Two-step
    rule now:
      1. A recognised rail prefix (NEFT/RTGS/IMPS/UPI/INFT, case-insensitive,
         optionally followed by "-CR"/"CR"/"IN" before the reference, with
         "/" treated as a separator alongside space/hyphen so slash-
         delimited UPI/IMPS narrations - e.g.
         "UPI/402912345678/username/YESB0000/paytm" - are handled, not just
         space-delimited NEFT/RTGS ones) - still exactly the original NEFT
         behaviour when the text is "NEFT <ref> ...", plus the same idea for
         the other rails. The capture itself stops at the next separator
         (letters/digits only) - EXCEPT some banks print the reference
         running directly into the remitter's name with NO separator at all
         (e.g. "NEFT AXISCN1024139802RAZORPAY SOFTWARE PRIVATE L..." - no
         space between the UTR and "RAZORPAY"), so a plain greedy alnum
         match would swallow the remitter's name straight into the "UTR"
         and never match the gateway's own recorded reference. Since
         NEFT/RTGS/INFT UTRs and IMPS/UPI RRNs are each a fixed, nationally
         standardised length (16 alphanumeric characters for NEFT/RTGS/INFT,
         12 digits for IMPS/UPI - not a per-bank convention), the captured
         token is capped at its own rail's known length (see
         RAIL_UTR_LENGTH below) - a safe no-op whenever a real delimiter
         was already present (the captured token is already exactly that
         length), and the fix whenever one wasn't.
      2. Otherwise, for narrations that don't start with a recognised prefix
         at all (e.g. a bank that puts the remitter's name before the
         reference): prefer a long all-digit token (10+ digits) first - the
         typical UPI RRN / IMPS reference number is purely numeric - before
         falling back to a mixed letters+digits token, 8+ characters, so a
         genuine all-digit reference isn't skipped in favour of some
         unrelated mixed-format token elsewhere in the narration. This can
         only ever produce a false match if some gateway's own recorded UTR
         happens to collide with an unrelated token in someone else's
         narration, which is vanishingly unlikely - so it's safe to leave
         broad.

    Still just a best-effort read of free-text bank narration - see this
    module's docstring: treat any row this helps match as a lead for the
    reconciling accountant to confirm, not a final answer. If a bank/client
    turns out to use a format this still misses, add another branch here
    rather than editing the gateway configs - this is the one place that
    should ever need to know about narration formats.

    require_rail_prefix (added 2026-08-21, client-reported): when True,
    ONLY the rail-prefix rule above (step 1) is attempted - the two weaker
    fallback heuristics (step 2: a bare long-digit or mixed alnum token,
    used when there's no recognised rail keyword at all) are skipped,
    returning None instead. Used by load_bank_statement() below to tell
    apart a HIGH-CONFIDENCE narration read (a real "NEFT/RTGS/IMPS/UPI
    <reference>..." match - this is a fixed, standardised bank-narration
    format, not a guess) from a low-confidence one, so a clean rail-tagged
    reference in the narration always wins over whatever a file's own
    auxiliary "Reference"/"Chq No" column happens to contain - that column
    was found, twice, to sometimes hold something other than the real
    bank UTR (a placeholder for one row, an unrelated internal reference
    number for another) even when it's consistently non-blank across the
    file.
    """
    if narration is None:
        return None
    text = str(narration).strip()
    if not text:
        return None

    prefix_match = re.match(
        r"^(NEFT|RTGS|IMPS|UPI|INFT)[\s\-/]*(?:CR|IN)?[\s\-/]+([A-Za-z0-9]+)",
        text, flags=re.IGNORECASE,
    )
    if prefix_match:
        rail = prefix_match.group(1).upper()
        ref = prefix_match.group(2).upper()
        expected_len = RAIL_UTR_LENGTH.get(rail)
        # Only ever trims, never pads - a reference genuinely shorter than
        # the standard length (a bank-specific quirk, or the narration
        # simply ended there) is left exactly as captured.
        if expected_len and len(ref) > expected_len:
            ref = ref[:expected_len]
        return ref

    if require_rail_prefix:
        return None

    digit_tokens = re.findall(r"\d{10,}", text)
    if digit_tokens:
        return digit_tokens[0]

    for token in re.findall(r"[A-Za-z0-9]{8,}", text):
        if re.search(r"\d", token) and re.search(r"[A-Za-z]", token):
            return token.upper()
    return None


def _looks_like_utr(value):
    """
    Client-reported (2026-08-21): a bank export's explicit "Reference"/
    "Chq/Ref No" column is sometimes non-blank for a row without actually
    containing a usable reference (e.g. a bank-side placeholder like "-",
    "0", or "NA" for a NEFT credit that has no cheque number) - and
    load_bank_statement() below prefers that column, when it's populated
    often enough across the file, over the Narration-extracted UTR for any
    row where it's non-blank. A placeholder value being merely "non-blank"
    was enough to win that preference and silently block a perfectly good
    narration-extracted UTR (e.g. AXISCN1301353011, cleanly extractable
    from "NEFT AXISCN1301353011 RAZORPAY PAYMENTS PVT LTD P") from ever
    being used for that row.

    A crude but effective filter: a real UTR/reference, however the issuing
    bank formats it, is never this short or purely non-numeric - requiring
    a minimum length and at least one digit rejects "-", "0", "NA", "NIL"
    etc. without rejecting any real reference format seen so far.
    """
    return len(value) >= 6 and any(ch.isdigit() for ch in value)


def _utr_suffix_lookup(utrs, suffix_len=UTR_SUFFIX_MATCH_LEN):
    """
    utr (already normalized) -> {last `suffix_len` chars: [full utrs sharing
    that suffix]}, built once per bank ledger and reused for every gateway
    UTR looked up against it (see match_utr_against_bank below), rather
    than rescanning the whole bank ledger per row.
    """
    lookup = {}
    for u in utrs:
        if u and len(u) >= suffix_len:
            lookup.setdefault(u[-suffix_len:], []).append(u)
    return lookup


def match_utr_against_bank(utr, bank_utr_set, suffix_lookup, suffix_len=UTR_SUFFIX_MATCH_LEN):
    """
    Matches one gateway-side UTR against the bank ledger's known UTRs
    (bank_utr_set - normalized values from load_bank_statement()), two ways:

      1. Exact match (normalized) - the common case.
      2. Fallback: the trailing `suffix_len` characters match exactly ONE
         bank UTR - see UTR_SUFFIX_MATCH_LEN's docstring above for why this
         is a safe, client-motivated fallback rather than a loose guess.
         Only applied when exactly one bank UTR shares that suffix - if
         more than one candidate does, that's genuinely ambiguous and left
         unmatched (the whole point of this fallback is to add confidence,
         not remove it by guessing between candidates).

    Returns (matched_bank_utr, match_type) - match_type is "exact",
    "suffix", or None (matched_bank_utr is also None when match_type is
    None).
    """
    if not utr:
        return None, None
    if utr in bank_utr_set:
        return utr, "exact"
    if len(utr) >= suffix_len:
        candidates = suffix_lookup.get(utr[-suffix_len:])
        if candidates and len(candidates) == 1:
            return candidates[0], "suffix"
    return None, None


def _parse_indian_amount(series):
    """
    Some bank exports (this module's original target format included, but
    also e.g. the client's Kotak CSV export) print larger amounts with
    Indian-style comma grouping and wrapped in quotes, e.g. "28,000.00" -
    pd.to_numeric() returns NaN for a string containing commas, silently
    dropping the transaction to 0 rather than erroring. Stripping thousands
    separators (and any stray currency symbol/whitespace) before the
    numeric parse is safe for every format this module handles, since a
    real amount never legitimately contains a comma or rupee sign.
    """
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def _resolve_bank_credit_amount(bank_df, bank_cfg, label):
    """
    Resolves the CREDIT-only amount column for a bank statement export,
    supporting two different shapes seen across real exports:

      1. Separate Withdrawal/Deposit columns (e.g. "Withdrawal (Dr)" /
         "Deposit") - bank_cfg["amount_col"] already points at the
         deposit-only column, so its value IS the credit amount as-is.

      2. One combined "Amount" column plus a separate Dr/Cr indicator
         column (e.g. the client's Kotak CSV export: "Amount" + "Dr / Cr")
         - here bank_cfg["amount_col"] resolves to a column that holds
         BOTH debits and credits, so it must be filtered down to only the
         rows the indicator column marks as a credit, or every withdrawal
         would be wrongly counted as money received.

      Distinguishing the two: if a configured "debit_col" resolves to an
      actual (and different) column in this file, shape 1 applies -
      amount_col is already credit-only. Otherwise, if a
      "credit_indicator_col" is configured and present, shape 2 applies.
    """
    amount_col = resolve_col_or_raise(bank_df, bank_cfg["amount_col"], label)
    indicator_col_spec = bank_cfg.get("credit_indicator_col")
    resolved_indicator_col = resolve_col(bank_df, indicator_col_spec) if indicator_col_spec else None

    debit_col_spec = bank_cfg.get("debit_col")
    resolved_debit_col = resolve_col(bank_df, debit_col_spec) if debit_col_spec else None
    has_separate_debit_col = resolved_debit_col is not None and resolved_debit_col != amount_col

    amt = _parse_indian_amount(bank_df[amount_col])

    if has_separate_debit_col or resolved_indicator_col is None:
        return amt

    credit_values = {v.strip().upper() for v in bank_cfg.get("credit_indicator_values", ["CR", "C", "CREDIT"])}
    indicator_text = bank_df[resolved_indicator_col].astype(str).str.strip().str.upper()
    return amt.where(indicator_text.isin(credit_values), 0.0)


def load_bank_statement(bank_df, bank_cfg):
    """
    Returns a clean bank ledger: utr | bank_amount | bank_date

    UTR resolution order (client-reported, 2026-08-21 - revised after TWO
    separate real cases where a file's own "Reference"/"Chq No" column beat
    a perfectly good, cleanly-extractable narration UTR: once with a bare
    placeholder value like "-", and again with what's presumably some
    other, unrelated reference the bank puts in that column):
      1. A HIGH-CONFIDENCE narration read: the narration matches the
         standardised "NEFT/RTGS/IMPS/UPI <reference> ..." bank format
         (extract_utr_from_narration(..., require_rail_prefix=True)) -
         this is a fixed national format, not a per-bank guess, so when
         it's present it's trusted over any auxiliary column the file
         happens to also carry.
      2. Otherwise, an explicit UTR-like column, if the file actually has
         one and it's not mostly blank. A value in that column that
         doesn't actually look like a reference (see _looks_like_utr()
         above - e.g. a bare "-"/"0"/"NA" placeholder) is treated as blank
         here, so it can't block step 3 below for that row either.
      3. Otherwise, the weaker narration fallback heuristics (a bare long
         digit run, or a mixed alnum token - see
         extract_utr_from_narration()'s own docstring) - covers a raw bank
         export with no dedicated UTR field AND no recognised rail prefix
         in its narration either.
    """
    label = bank_cfg.get("label", "Bank Statement")
    date_col = bank_cfg.get("date_col")

    out = pd.DataFrame(index=bank_df.index)

    utr_col_spec = bank_cfg.get("utr_col")
    resolved_utr_col = resolve_col(bank_df, utr_col_spec) if utr_col_spec else None
    if resolved_utr_col:
        explicit_utr_raw = normalize_utr(bank_df[resolved_utr_col].fillna(""))
        explicit_utr = explicit_utr_raw.where(explicit_utr_raw.apply(_looks_like_utr), "")
    else:
        explicit_utr = pd.Series([""] * len(bank_df), index=bank_df.index)

    narration_col_spec = bank_cfg.get("narration_col")
    resolved_narration_col = resolve_col(bank_df, narration_col_spec) if narration_col_spec else None
    if resolved_narration_col:
        narration_col_raw = bank_df[resolved_narration_col]
        narration_utr_strict = narration_col_raw.apply(lambda n: extract_utr_from_narration(n, require_rail_prefix=True))
        narration_utr_strict = narration_utr_strict.fillna("").astype(str).str.upper()
        narration_utr_any = narration_col_raw.apply(extract_utr_from_narration)
        narration_utr_any = narration_utr_any.fillna("").astype(str).str.upper()
    else:
        narration_utr_strict = pd.Series([""] * len(bank_df), index=bank_df.index)
        narration_utr_any = pd.Series([""] * len(bank_df), index=bank_df.index)

    # Step 1 (rail-prefixed narration) first; step 2 (explicit column) only
    # for rows step 1 didn't resolve; step 3 (weaker narration fallback)
    # only for rows neither of the above resolved.
    out["utr"] = narration_utr_strict.where(
        narration_utr_strict.str.len() > 0,
        explicit_utr.where(explicit_utr.str.len() > 0, narration_utr_any),
    )

    out["bank_amount"] = _resolve_bank_credit_amount(bank_df, bank_cfg, label)

    resolved_date_col = resolve_col_or_raise(bank_df, date_col, label) if date_col else None
    # Normalized to tz-naive here, at the one place bank dates are parsed -
    # see to_naive_datetime_series()'s docstring above. A bank export's own
    # date column is rarely tz-aware, but this keeps bank_date guaranteed
    # comparable against every other date this module handles, rather than
    # relying on every downstream consumer to remember to check.
    out["bank_date"] = to_naive_datetime_series(bank_df[resolved_date_col]) if resolved_date_col else pd.NaT

    out = out[out["utr"].str.len() > 0]
    out = out[out["utr"] != "NAN"]
    return out.reset_index(drop=True)


def match_consolidated_to_bank(consolidated_df, bank_ledger_df):
    """
    Matches each gateway transaction's UTR against the bank statement.
    Adds bank_matched (bool), bank_date, matched_bank_utr, and
    bank_match_type ("exact"/"suffix"/None) columns to the consolidated
    receipt ledger. A gateway row with no UTR captured, or a UTR not found
    in the bank statement (by either method), is left unmatched - which is
    itself useful information ("bank receipt not identified").

    Matching is exact-first, falling back to match_utr_against_bank()'s
    UTR-suffix heuristic when no exact match exists - see that function's
    docstring for why (a gateway's own recorded UTR and what actually
    prints in the bank narration for the same credit can be genuinely
    different formats around the same core reference).

    matched_bank_utr is exposed (rather than just the bool) because it can
    legitimately differ from this row's own "utr" value on a suffix match -
    anything that needs to exclude an already-matched bank credit (e.g.
    matched_order_level_utrs() below) must exclude by THIS column, not by
    the gateway-side "utr" column, since the latter may not appear in the
    bank ledger's own utr values at all.
    """
    df = consolidated_df.copy()
    if "utr" not in df.columns or bank_ledger_df is None or bank_ledger_df.empty:
        df["bank_matched"] = False
        df["bank_date"] = pd.NaT
        df["matched_bank_utr"] = None
        df["bank_match_type"] = None
        return df

    df["utr_norm"] = normalize_utr(df["utr"].fillna(""))
    bank_lookup = bank_ledger_df.drop_duplicates(subset="utr").set_index("utr")
    bank_utr_set = set(bank_lookup.index)
    suffix_lookup = _utr_suffix_lookup(bank_utr_set)

    match_results = df["utr_norm"].apply(lambda u: match_utr_against_bank(u, bank_utr_set, suffix_lookup))
    df["matched_bank_utr"] = match_results.apply(lambda m: m[0])
    df["bank_match_type"] = match_results.apply(lambda m: m[1])
    df["bank_matched"] = df["matched_bank_utr"].notna()
    df["bank_date"] = df["matched_bank_utr"].map(bank_lookup["bank_date"].to_dict())
    df = df.drop(columns=["utr_norm"])
    return df


def resolve_split_payment_leg_status(reco_df, consolidated_df, bank_ledger_df, gateway_configs):
    """
    2026-09-06 (round 17) - client-reported direct follow-up to round 16's
    order #30456 Payment Provider fix ("Delhivery COD, PayU"): "query
    update missing... this case delhivery COD setled but payu setlment
    pending, query to be asked 'Delhivery COD Setled, payu Setlment
    Pending'... Settlement Pending & Exceptions (by Payment Gateway)
    should be update[d]".

    Root cause this needs its OWN function rather than reusing
    classify_order_bank_status() above: that function computes
    `bank_matched` as `payments.groupby("order_id")["bank_matched"].any()`
    - an ORDER-LEVEL aggregate across every one of an order's consolidated_
    df rows, regardless of which "leg" (a COD-mode gateway vs a Prepaid-
    mode gateway, per gateway_configs' own payment_mode field) each row
    belongs to. That's the right call for classify_order_bank_status's own
    job (one order needs exactly one of six categories), but for a genuine
    split-payment order it means ONE settled leg (here, Delhivery COD
    already bank-matched) makes the WHOLE order read as fully resolved,
    completely masking that the OTHER leg (PayU) has a confirmed
    transaction that simply hasn't reached the bank yet. Both
    engine.reco.refine_queries_with_settlement_status()'s Query text and
    engine.settlement_pending.settlement_pending_summary_by_gateway()'s
    pending-order filter read that same masked, order-level signal - this
    function is the shared, per-leg-aware alternative both of those need,
    built by reusing match_consolidated_to_bank()'s existing ROW-level
    (non-aggregated) bank_matched output rather than re-deriving it.

    Scope, disclosed: matches each leg's rows against the bank statement
    by DIRECT UTR ONLY (match_consolidated_to_bank()'s exact-then-suffix
    matching) - it deliberately does NOT replicate classify_order_bank_
    status()'s additional COD settlement-BATCH amount/date fallback
    (build_cod_settlement_batches()/match_batches_to_bank()), which groups
    an entire COD gateway's rows by (source, settlement date) across the
    WHOLE ledger and matches the batch total to a bank credit - that
    logic doesn't decompose cleanly per split-payment order without
    materially more risk (a batch match is shared across many unrelated
    orders' rows at once). Practical effect: a split-payment order whose
    COD leg genuinely settled only via that batch-level fallback (no
    direct UTR match of its own) will read `cod_matched: False` here even
    though classify_order_bank_status()'s fuller logic might already
    treat the order's overall Reconciliation Category as COD_BANK_MATCHED
    - a disclosed simplification, not silently guessed around.

    Returns: {order_id: {"cod_matched": bool or None, "prepaid_matched":
    bool or None}} for every order_id in reco_df that has at least one
    consolidated_df row from either a COD-mode or Prepaid-mode gateway.
    None for a leg means "no row for this order from that leg's gateways
    at all" (not the same as False - "row(s) exist, none bank-matched
    yet"). An order with no rows from either leg at all is omitted
    entirely.
    """
    if reco_df is None or reco_df.empty or consolidated_df is None or consolidated_df.empty:
        return {}

    gateway_configs = gateway_configs or []
    cod_labels = {cfg["label"] for cfg in gateway_configs if str(cfg.get("payment_mode", "")).strip().lower() == "cod"}
    prepaid_labels = {cfg["label"] for cfg in gateway_configs if str(cfg.get("payment_mode", "")).strip().lower() == "prepaid"}
    if not cod_labels and not prepaid_labels:
        return {}

    order_ids = set(reco_df["order_id"].astype(str))

    with_bank = match_consolidated_to_bank(consolidated_df, bank_ledger_df)
    payments = with_bank[~with_bank["is_refund"]].copy()
    if payments.empty:
        return {}
    payments["order_id"] = payments["order_id"].astype(str)
    payments = payments[payments["order_id"].isin(order_ids)]
    if payments.empty:
        return {}

    cod_rows = payments[payments["source"].isin(cod_labels)]
    prepaid_rows = payments[payments["source"].isin(prepaid_labels)]

    cod_matched_by_order = cod_rows.groupby("order_id")["bank_matched"].any().to_dict()
    prepaid_matched_by_order = prepaid_rows.groupby("order_id")["bank_matched"].any().to_dict()

    result = {}
    for oid in order_ids:
        cod_matched = cod_matched_by_order.get(oid)
        prepaid_matched = prepaid_matched_by_order.get(oid)
        if cod_matched is None and prepaid_matched is None:
            continue
        result[oid] = {"cod_matched": cod_matched, "prepaid_matched": prepaid_matched}

    return result


def matched_order_level_utrs(consolidated_df, bank_ledger_df):
    """
    Set of normalized UTRs already claimed by a direct order-level match
    (match_consolidated_to_bank above) - passed as `exclude_utrs` to
    match_batches_to_bank() so the settlement-batch amount/date fallback
    can never re-match a bank credit that's already accounted for by a
    real reference-number match (which would otherwise double-count that
    one credit's money across two different orders/gateways). Shared here
    so classify_order_bank_status() and any direct caller building batches
    for display (e.g. the Bank Linking page) use the exact same exclusion
    set rather than two versions that could disagree.
    """
    if consolidated_df is None or consolidated_df.empty or bank_ledger_df is None or bank_ledger_df.empty:
        return set()
    matched = match_consolidated_to_bank(consolidated_df, bank_ledger_df)
    payments = matched[~matched["is_refund"]]
    if "matched_bank_utr" not in payments.columns or not payments["bank_matched"].any():
        return set()
    # Excludes by the BANK ledger's own UTR value, not the gateway's "utr"
    # column - match_batches_to_bank() filters bank_ledger_df["utr"] by this
    # set, so it must contain values that actually appear there. On a
    # suffix match (see match_consolidated_to_bank above) those two can
    # legitimately differ.
    return set(payments.loc[payments["bank_matched"], "matched_bank_utr"].dropna())


def bank_status_by_order(consolidated_with_bank_df):
    """
    Per-order rollup: has the receiving bank credit actually been traced
    for this order, and when. Feeds the Order Lookup's "Bank credit date"
    field and the exception report's "Bank receipt not identified" flag.
    """
    if consolidated_with_bank_df.empty:
        return pd.DataFrame(columns=["order_id", "bank_credit_date", "bank_status"])

    payments = consolidated_with_bank_df[~consolidated_with_bank_df["is_refund"]].copy()

    grouped = payments.groupby("order_id").agg(
        bank_credit_date=("bank_date", "min"),
        bank_status=("bank_matched", lambda s: "Matched" if s.any() else "Bank receipt not identified"),
    ).reset_index()

    return grouped


def build_settlement_ledger(consolidated_df, period_by_order_id):
    """
    One row per (order_id, UTR) pair that carries a real reference, shaped
    for bank_reconciliation_by_utr() below:
        order_id | utr | final_payment | period_bucket

    consolidated_df: engine.consolidator.build_consolidated_receipt()
        output - every individual gateway transaction row (order_id |
        source | amount | deduction | is_refund | utr | ...), NOT
        pre-collapsed to one row per order. That matters specifically
        here: an order settled across more than one genuinely distinct
        UTR (e.g. a COD order remitted in two separate courier batches,
        each its own bank credit) needs one ledger row PER UTR, each
        carrying only the amount actually settled under THAT reference -
        not the order's combined total attached to a single comma-joined
        "UTR1, UTR2" string that can never match any one real bank
        credit. Grouping straight off this raw per-transaction table
        (rather than off engine.consolidator.summarize_receipts_by_order/
        receipt_detail_by_order's already-collapsed one-row-per-order
        views) is what makes that possible - see this module's history
        (client-reported: "multiple UTRs reflecting in a single row",
        and separately, a single genuinely-correct UTR still not
        matching because a blank UTR on another row for the same order
        had been silently joined in alongside it).
    period_by_order_id: dict order_id (str) -> "This period" /
        "Previous period" / "Subsequent period" / "Order not found"
        (see engine/period.py classify_order_periods()).

    final_payment mirrors the sample workbook's "Final Payment" column:
    receipt minus deduction minus refund - i.e. this engine's existing
    settlement_amount concept (see engine/reco.py attach_receipts_and_diff),
    computed here per (order_id, utr) group. A refund row that happens to
    carry the SAME utr as its original payment nets against it correctly;
    a refund with no utr of its own (common - refunds aren't always
    re-settled via their own bank reference) simply doesn't reduce any
    UTR group's total here, which is correct for a per-UTR reconciliation
    view - it still reduces the order's own net receipt everywhere else
    in this engine (see engine.consolidator.summarize_receipts_by_order),
    it just was never actually paid out against this particular reference.
    """
    cols = ["order_id", "utr", "final_payment", "period_bucket"]
    if consolidated_df is None or consolidated_df.empty or "utr" not in consolidated_df.columns:
        return pd.DataFrame(columns=cols)

    df = consolidated_df.copy()
    # fillna("") BEFORE astype(str) - not just .astype(str) alone. A gateway
    # row with no resolvable order_id (see engine.consolidator.normalize_gateway_df's
    # "blank order id but real money" case - e.g. a Razorpay settlement row
    # this tool's own order-linking step couldn't trace to a Shopify order)
    # has a genuinely missing order_id. Since pandas 3.0, plain
    # `.astype(str)` on a "string"-dtype column no longer stringifies a
    # missing value to the literal text "nan" the way it used to (and the
    # way an object-dtype column's .astype(str) still does) - the missing
    # value survives as a real NaN, and this column's own dtype stays
    # nullable. That NaN then makes it all the way into the
    # groupby(["order_id", "utr"]) calls below, which - like every pandas
    # groupby - drops NaN-keyed groups by default. The row's UTR (a real,
    # matchable bank reference) and its settled amount would then simply
    # vanish before ever reaching bank_reconciliation_by_utr(), which is
    # exactly the client-reported symptom (2026-08-22): a UTR confirmed
    # present in both the bank statement and the payment gateway report
    # not appearing AT ALL in the "Bank Reco (UTR-wise)" sheet, rather than
    # appearing with a wrong Remark. fillna("") first turns that missing
    # value into a real, groupable empty-string key (consistent with how a
    # blank order_id is represented everywhere else in this engine - see
    # engine.consolidator._blank_order_id_mask), so the row's UTR still
    # gets its own row in the per-UTR rollup, just with an empty order_id.
    df["order_id"] = df["order_id"].fillna("").astype(str)
    df["utr"] = normalize_utr(df["utr"].fillna(""))
    df = df[(df["utr"].str.len() > 0) & (~df["utr"].isin(["NAN", "NONE", "NAT"]))]
    if df.empty:
        return pd.DataFrame(columns=cols)

    payments = df[~df["is_refund"]]
    refunds = df[df["is_refund"]]

    receipt = payments.groupby(["order_id", "utr"])["amount"].sum().rename("receipt_amount")
    deduction = payments.groupby(["order_id", "utr"])["deduction"].sum().rename("total_deduction")
    refund = refunds.groupby(["order_id", "utr"])["amount"].sum().rename("refund_amount")

    grouped = pd.concat([receipt, deduction, refund], axis=1).fillna(0.0).reset_index()
    grouped["final_payment"] = grouped["receipt_amount"] - grouped["total_deduction"] - grouped["refund_amount"]

    period_by_order_id = period_by_order_id or {}
    grouped["period_bucket"] = grouped["order_id"].map(period_by_order_id).fillna("Order not found")

    return grouped[cols]


_REMARK_PHRASES = [
    # (bucket key, phrase) - order here is also the display order joined
    # by " + " in a composed Remarks string. Matches the client's own
    # corrected working (2026-08-23) exactly, phrase for phrase - see
    # bank_reconciliation_by_utr's docstring for how each bucket is
    # computed and _classify_utr_remark for how they're composed.
    ("this_period_now", "Same-period txn settled same period"),
    ("this_period_subsequent", "This-period txn settled next period"),
    ("subsequent_order", "Next-period txn settled next period"),
    ("previous_order", "Prior-period txn settled this period"),
    ("not_found_now", "Order ID not found - settled in reco period"),
    ("not_found_subsequent", "Order ID not found - settled after reco period"),
]


def _classify_utr_remark(buckets, bank_credit_total, difference,
                          order_ids=None, consolidated_df=None, match_type=None, matched_bank_utr=None, utr=None):
    """
    Deterministic, disclosed rules approximating the Remarks categories
    used in the sample workbook. See this module's docstring: the
    sample's own Remarks column mixed formula-driven and manually-typed
    values, so treat any row this labels for review as exactly that - a
    prompt for the reconciling accountant to check by hand, not a final
    answer.

    buckets (revised 2026-08-23, client-reported: compared against her own
    manually corrected "Reco working" sheet line by line): a dict with six
    keys, each this UTR's claimed amount broken down by BOTH which period
    the underlying order belongs to AND, for the two buckets where it's
    possible to tell, whether the bank credit for it posted within the
    selected report window or after it closed:
      - this_period_now: a THIS-period order, its own settlement posted
        within the selected report window.
      - this_period_subsequent: a THIS-period order, its own settlement
        only posted AFTER the report window closed (the "Settled
        Subsequent Period" case).
      - previous_order: a PREVIOUS-period order settled during the current
        window (the existing "Settled (other period)" scenario) - shown as
        its own claimed amount regardless of exactly which date it posted,
        same as this column has always behaved.
      - subsequent_order: a SUBSEQUENT-period order (belongs to a period
        AFTER the one selected) whose own settlement is - as it must be,
        an order can't be paid before it's placed - also dated after the
        window. Genuinely unrelated to the current report except that it
        shares a UTR/settlement batch with something that IS related;
        client-reported: conflating this with "Settled (other period)"
        wrongly implied it was previous-period money.
      - not_found_now / not_found_subsequent: a settlement ledger row
        whose order_id could never be resolved to any known order at all
        (engine.period.NOT_FOUND - blank order id, or a real-looking one
        that was never uploaded and has no historical date on record),
        split the same in-window/after-window way as this_period's own
        two buckets. This is the reinstated "Order ID Not Found" concept
        (client-reported 2026-08-23: "preserve... Order ID Not Found
        classifications" - the single combined column removed on
        2026-08-21 undercounted this: it only caught a bank credit
        EXCEEDING every known order's claim, never a real ledger row whose
        own order simply couldn't be dated/resolved).
    Every bucket a caller doesn't have a period end to split by (e.g. the
    live, single-period Reconciliation page - see
    bank_reconciliation_by_utr's `period_end_date` parameter) collapses
    into its "_now" half - this_period_subsequent and not_found_subsequent
    are always 0.0 there, exactly matching this function's behaviour
    before this six-way split existed.

    bank_credit_total: this UTR's TOTAL bank credit found anywhere in the
    available data (in-window plus after-window combined) - client's own
    working shows "Bank Credit" as this plain total regardless of timing,
    even for a fully-subsequent-period UTR; the six buckets above are
    what carries the in-window-vs-after-window story, not this column.

    Client-reported (2026-08-21): a bare "Bank amount does not tie out -
    review manually" wasn't good enough for two very identifiable causes:
      1. A REFUND for one of this UTR's own order(s) is sitting right
         there in the uploaded payment gateway report (consolidated_df),
         just not netted against this specific UTR - refunds commonly
         carry no settlement UTR of their own (or a different one, from
         whichever LATER batch actually deducted it), so
         build_settlement_ledger's own per-(order_id, utr) grouping never
         attaches it to the original payment's UTR. Searched for here by
         ORDER ID instead (order_ids: every order settled under this UTR,
         from settlement_ledger_df - see bank_reconciliation_by_utr),
         which finds it regardless of what UTR the refund itself carries.
      2. The bank credited MORE than this UTR's own known orders claim -
         almost always money collected through the SAME payment gateway
         for something that was never part of the Shopify order/gateway
         reports at all, and with no row for it anywhere in
         consolidated_df either (contrast this with not_found_now/
         not_found_subsequent above, which DO have a real ledger row,
         just an unresolvable order_id) - a classic case: offline/POS
         sales processed through the same gateway account, batched into
         the same settlement.

    order_ids / consolidated_df: optional - when not supplied (e.g. an
    older caller), falls back to the original bare "review manually"
    behaviour for an untied-out UTR, same as before this cross-check
    existed.

    match_type / matched_bank_utr / utr (added 2026-08-21, client-reported):
    when the bank credit was found via match_utr_against_bank()'s UTR-suffix
    fallback rather than an exact match (see UTR_SUFFIX_MATCH_LEN's
    docstring in this module - e.g. gateway-recorded "ICICN22025102703875831"
    vs the bank narration's own "ICIN230003875831", same core reference,
    different bank-generated formatting), every remark this function would
    otherwise return gets a bracketed note naming the actual bank-side
    reference and flagging the format mismatch for a quick manual glance -
    a genuine match, surfaced honestly as the heuristic it is, exactly like
    every other cross-check in this module.

    Returns a single remark string.
    """
    total_claimed = sum(buckets.values())
    bank_found_anywhere = abs(bank_credit_total) > TOLERANCE or abs(total_claimed) <= TOLERANCE
    tied_out = abs(difference) <= TOLERANCE
    has_other_period = abs(buckets["previous_order"]) > TOLERANCE or abs(buckets["subsequent_order"]) > TOLERANCE

    suffix_note = (
        f" [Matched via UTR core reference: bank narration shows '{matched_bank_utr}', payment gateway report "
        f"recorded '{utr}' - same trailing reference digits, different bank-generated prefix/date formatting. "
        "Confirm manually.]"
        if match_type == "suffix" and matched_bank_utr else ""
    )

    # Genuinely not found ANYWHERE in the available data (not just outside
    # the selected window). Checked first and unconditionally, so this can
    # never be reached for a UTR that IS actually backed by a real bank
    # credit somewhere, however it later gets bucketed.
    if not bank_found_anywhere:
        base = "Bank statement not found for this UTR - needs manual check"
        return f"{base} (includes other-period settlement)" if has_other_period else base

    if tied_out:
        phrases = [phrase for key, phrase in _REMARK_PHRASES if abs(buckets[key]) > TOLERANCE]
        if not phrases:
            # Degenerate case (e.g. every bucket nets to ~0 - a UTR whose
            # claimed total and bank credit are both effectively zero) -
            # bank_found_anywhere's own "or abs(total_claimed) <= TOLERANCE"
            # branch already covers "nothing claimed, nothing credited" as
            # "not applicable" territory upstream in most callers, but kept
            # here too so this function never returns an empty string.
            phrases = ["Matched"]
        return " + ".join(phrases) + " | Ties to bank credit" + suffix_note

    # difference = total claimed - bank credit found anywhere, so
    # difference > 0 means we claimed MORE than the bank has paid so far
    # (a refund is the classic cause); difference < 0 means the bank paid
    # MORE than we claimed (an unattributed extra collection through the
    # same gateway/UTR, with no row for it anywhere in our own data).
    if order_ids and difference > TOLERANCE and consolidated_df is not None and not consolidated_df.empty \
            and "is_refund" in consolidated_df.columns:
        refund_rows = consolidated_df[
            consolidated_df["is_refund"] & consolidated_df["order_id"].astype(str).isin(order_ids)
        ]
        refund_total = round(float(refund_rows["amount"].sum()), 2) if len(refund_rows) else 0.0
        if refund_total > TOLERANCE:
            if abs(difference - refund_total) <= TOLERANCE:
                return (
                    f"Bank amount is lower by Rs.{refund_total:,.2f} than the total claimed for this UTR - "
                    "matches a customer refund already present in the payment gateway report for the "
                    "order(s) settled under this UTR. Ties out once the refund is accounted for."
                ) + suffix_note
            return (
                f"Bank amount does not fully tie out - a refund of Rs.{refund_total:,.2f} was found in the "
                f"payment gateway report for the order(s) under this UTR, but it doesn't fully explain the "
                f"Rs.{difference:,.2f} difference. Review manually."
            ) + suffix_note

    if difference < -TOLERANCE:
        not_found_amt = round(abs(difference), 2)
        return (
            f"Bank credit includes an amount of Rs.{not_found_amt:,.2f} not matched to any uploaded order - "
            "likely a settlement for an order not present in the uploaded Sales/Payment Gateway reports "
            f"(e.g. an offline/POS sale through the same gateway account).{suffix_note}"
        )

    return "Bank amount does not tie out - review manually" + suffix_note


def bank_reconciliation_by_utr(settlement_ledger_df, bank_ledger_df, consolidated_df=None,
                                period_start_date=None, period_end_date=None,
                                utr_gateway_lookup=None):
    """
    settlement_ledger_df: build_settlement_ledger() output above.
    bank_ledger_df: load_bank_statement() output above.
    consolidated_df: engine.consolidator.build_consolidated_receipt()
        output (every individual gateway transaction row) - optional, but
        strongly recommended: lets the Remarks column cross-check a
        difference against refunds already present in the payment
        gateway report for this UTR's own order(s), rather than always
        falling back to a bare "review manually" (see
        _classify_utr_remark's docstring above for exactly what this
        catches and what it can't).
    period_start_date (added 2026-08-31, client-reported - Executive
        Summary Point 7 "Receivable Collection Period Analysis"): the
        START of the reconciliation period this call is reporting on -
        the Reports page's selected "From date" (or the whole financial
        year's first day when left blank - see views/page_reports.py's
        _render_dtc). Paired with period_end_date below to give the
        in-window bank-date test a genuine LOWER bound, not just an
        upper one - see the "THE BUG THIS FIXES" note below for exactly
        why this was missing and what it broke. None (the default, and
        what the live single-period Reconciliation page still passes)
        means "no lower bound" - nothing is ever classified as predating
        the window, same as before this parameter existed.
    period_end_date: the END of the reconciliation period this call is
        reporting on - for the Reports page, the selected "To date" (or
        the whole financial year's own last day when "To date" was left
        blank - see views/page_reports.py's _render_dtc). Used ONLY to
        decide, for a THIS-period order or an unresolvable ("not found")
        ledger row, whether ITS OWN bank credit posted within the
        selected window or after it closed - see the six buckets
        documented in _classify_utr_remark's docstring above. Passing
        None (the live, single-period Reconciliation page's call - see
        views/page_reconciliation.py) means everything found is treated
        as posted within the window - the "_subsequent" half of every
        split bucket is always 0.0.

    Each UTR's bank credit is looked up via match_utr_against_bank() -
    exact match first, falling back to the UTR-suffix heuristic (see
    UTR_SUFFIX_MATCH_LEN above) when the payment gateway's own recorded UTR
    and the bank statement's own reference for the same credit are
    genuinely different formats around the same core reference. A
    suffix-based match is flagged inline in the Remarks column rather than
    silently presented as identical to an exact match.

    utr_gateway_lookup: optional Series/dict, normalized UTR -> payment
        gateway label (see engine.attribution.build_payment_gateway_lookups) -
        populates the "Payment Gateway" column (2026-08-25 client
        request). A UTR with no entry shows "Unknown" - not left blank,
        so it's visibly a gap rather than looking like a genuine "no
        gateway involved" case.

    Returns one row per distinct UTR (revised 2026-08-25, client-reported -
    compared line by line against her own manually corrected working):
        UTR | Bank Date | Payment Gateway | Amount (this period settled this period) |
        Amount (this period transaction Settled Subsequent Period) |
        Settled (subsequent period transaction subsequent period) |
        Settled (previous period transaction settled this period) |
        Order ID Not Found - Settled During Reco Period |
        Order ID Not Found - Settled After Reco Period |
        Total | Bank Credit | Difference | Remarks |
        Settled (previous period transaction settled before this period - already reported) |
        Settled (previous period transaction not yet settled as of this period's close) |
        Order ID Not Found - Settled Before Reco Period (Already Reported)

    This replaces the previous ("Amount (this period)" / "Settled (other
    period)" / "Settled Subsequent Period") three-way split, which - the
    client's own corrected numbers showed - conflated three genuinely
    different situations into one "Settled (other period)" figure: a
    PREVIOUS-period order settled this period (the original, still-valid
    meaning of that column), a SUBSEQUENT-period order that happens to
    share a UTR with something relevant to this report (new: "Settled
    (subsequent period transaction subsequent period)" - completely
    unrelated to the current window otherwise), and a settlement ledger
    row whose order could never be resolved at all (new: the reinstated,
    now two-way-split "Order ID Not Found" columns - see
    _classify_utr_remark's docstring). "Bank Date" (new) is the earliest
    bank-ledger date found for this UTR, for a quick eyeball without
    cross-referencing the bank statement sheet separately. "Bank Credit"
    keeps its original meaning: the UTR's full matched bank credit,
    regardless of timing - the two "Settled Subsequent"/"...After Reco
    Period" buckets are a breakdown of what's ALREADY inside that total,
    not a subtraction from it.

    THE BUG THIS FIXES (2026-08-31, client-reported - Executive Summary
    Point 7 "Receivable Collection Period Analysis"): "Settled (previous
    period transaction settled this period)" - the column that section's
    "Previous Period Amount Received This Period" is built from - used to
    show a PREVIOUS-period order's full claimed settlement regardless of
    WHEN the bank actually credited it, because this function never split
    that bucket by timing at all (unlike this_period/not_found, which at
    least had a two-way now/subsequent split). Worse, the in-window test
    those two buckets DID have was only ever one-sided
    (`bank_date <= period_end`, no lower bound) - so even they would
    wrongly re-count a bank credit from many months ago as "settled this
    period" every single time a later month's report was generated. Both
    problems share the same root cause: this function was never told when
    the selected period actually STARTS, only when it ends. Reproduced
    exactly against the client's own stated example: filtering "July"
    showed the ENTIRE June receipt total under "Previous Period Amount
    Received This Period" instead of only the June orders whose bank
    credit genuinely posted in July.

    Fixed by adding the period_start_date parameter above and using it,
    together with period_end_date, to classify every UTR's bank credit
    into exactly one of three buckets - "before" (posted before
    period_start - i.e. already reported as this same UTR's "now" amount
    in an EARLIER period's own report; must not be recounted here),
    "now" (posted within [period_start, period_end] - genuinely this
    report's own money), or "after" (posted after period_end - not yet
    received as of this period's close, a roll-forward candidate for a
    later period's report) - then applying that three-way split to
    previous_order and not_found the same way this_period already had a
    (now only two-way) split. A UTR with no parseable bank date at all
    still can't be judged against either bound and is treated as "now",
    exactly the same "no evidence, don't manufacture a split" convention
    this module already used for the upper bound alone.

    Deliberately NOT changed by this fix: "Total"/"Bank Credit"/
    "Difference"/"Remarks" stay full-history figures (every rupee ever
    claimed vs found for this UTR, regardless of which period's report
    is being generated) - unaffected by the fix above and still exactly
    what they were before it. Those columns answer "does this UTR
    reconcile AT ALL", a genuinely different question from "which
    period's report should this money be shown against" that the new
    split answers - conflating the two would have silently changed the
    tie-out logic multiple already-verified reports depend on.
    """
    cols = [
        "UTR", "Bank Date", "Payment Gateway",
        "Amount (this period settled this period)",
        "Amount (this period transaction Settled Subsequent Period)",
        "Settled (subsequent period transaction subsequent period)",
        "Settled (previous period transaction settled this period)",
        "Order ID Not Found - Settled During Reco Period",
        "Order ID Not Found - Settled After Reco Period",
        "Total", "Bank Credit", "Difference", "Remarks",
        "Settled (previous period transaction settled before this period - already reported)",
        "Settled (previous period transaction not yet settled as of this period's close)",
        "Order ID Not Found - Settled Before Reco Period (Already Reported)",
    ]
    if settlement_ledger_df is None or settlement_ledger_df.empty:
        return pd.DataFrame(columns=cols)

    df = settlement_ledger_df.copy()
    df["utr"] = normalize_utr(df["utr"].fillna(""))
    df = df[df["utr"].str.len() > 0]
    if df.empty:
        return pd.DataFrame(columns=cols)

    df["_this_period"] = df["final_payment"].where(df["period_bucket"] == THIS_PERIOD, 0.0)
    df["_previous_order"] = df["final_payment"].where(df["period_bucket"] == PREVIOUS_PERIOD, 0.0)
    df["_subsequent_order"] = df["final_payment"].where(df["period_bucket"] == SUBSEQUENT_PERIOD, 0.0)
    df["_not_found"] = df["final_payment"].where(df["period_bucket"] == NOT_FOUND, 0.0)

    grouped = df.groupby("utr").agg(
        this_period=("_this_period", "sum"),
        previous_order=("_previous_order", "sum"),
        subsequent_order=("_subsequent_order", "sum"),
        not_found=("_not_found", "sum"),
        # Every order settled under this UTR - passed to
        # _classify_utr_remark so an unrelated difference can be
        # cross-checked against refunds for THESE SPECIFIC orders in the
        # payment gateway report, not just a bare amount comparison.
        order_ids=("order_id", lambda s: sorted(set(s.astype(str)))),
    ).reset_index()
    grouped["total"] = grouped[["this_period", "previous_order", "subsequent_order", "not_found"]].sum(axis=1)

    if bank_ledger_df is not None and not bank_ledger_df.empty:
        bank_ledger_df = bank_ledger_df.copy()
        bank_ledger_df["utr"] = normalize_utr(bank_ledger_df["utr"].fillna(""))
        bank_dates = to_naive_datetime_series(bank_ledger_df["bank_date"])
        # THE FIX (2026-08-31, client-reported - see this function's own
        # docstring "THE BUG THIS FIXES"): the in-window test used to be
        # one-sided (<= period_end only), so ANY bank credit ever recorded
        # - even one already reported as "this period" money in a much
        # earlier month's own run - counted as "in window" again every
        # time a later period's report was generated. Now genuinely
        # two-sided: "before" (predates period_start - already reported,
        # must not be recounted), "in window" (between period_start and
        # period_end inclusive - genuinely this report's own money), and
        # "after" (postdates period_end - not yet received as of this
        # period's close). A bank row with no parseable date at all can't
        # be judged against either bound - treated as in-window, same
        # "no evidence, don't manufacture a split" convention this module
        # already used for the upper bound alone.
        if period_start_date is not None:
            period_start_ts = to_naive_timestamp(period_start_date)
            before_start_mask = bank_dates.notna() & (bank_dates < period_start_ts)
        else:
            before_start_mask = pd.Series(False, index=bank_ledger_df.index)
        if period_end_date is not None:
            period_end_ts = to_naive_timestamp(period_end_date)
            after_end_mask = bank_dates.notna() & (bank_dates > period_end_ts)
        else:
            after_end_mask = pd.Series(False, index=bank_ledger_df.index)
        in_window_mask = ~before_start_mask & ~after_end_mask
        bank_by_utr_window = bank_ledger_df[in_window_mask].groupby("utr")["bank_amount"].sum()
        bank_by_utr_subsequent = bank_ledger_df[after_end_mask].groupby("utr")["bank_amount"].sum()
        bank_by_utr_before = bank_ledger_df[before_start_mask].groupby("utr")["bank_amount"].sum()
        bank_date_by_utr = bank_ledger_df.groupby("utr")["bank_date"].min()
        # Matching itself must consider a UTR found ANYWHERE (any of the
        # three buckets) - the window split only decides which column the
        # this-period/previous-period/not-found money lands in, never
        # whether the UTR counts as "found" at all.
        bank_utr_set = set(bank_ledger_df["utr"])
    else:
        bank_by_utr_window = pd.Series(dtype=float)
        bank_by_utr_subsequent = pd.Series(dtype=float)
        bank_by_utr_before = pd.Series(dtype=float)
        bank_date_by_utr = pd.Series(dtype="datetime64[ns]")
        bank_utr_set = set()

    suffix_lookup = _utr_suffix_lookup(bank_utr_set)
    match_results = grouped["utr"].apply(lambda u: match_utr_against_bank(u, bank_utr_set, suffix_lookup))
    grouped["_matched_bank_utr"] = match_results.apply(lambda m: m[0])
    grouped["_match_type"] = match_results.apply(lambda m: m[1])
    grouped["_bank_in_window"] = grouped["_matched_bank_utr"].map(bank_by_utr_window).fillna(0.0)
    grouped["_bank_after_window"] = grouped["_matched_bank_utr"].map(bank_by_utr_subsequent).fillna(0.0)
    grouped["_bank_before_window"] = grouped["_matched_bank_utr"].map(bank_by_utr_before).fillna(0.0)
    grouped["Bank Date"] = grouped["_matched_bank_utr"].map(bank_date_by_utr)
    # "Bank Credit" keeps its original, full-history meaning (see this
    # function's docstring) - now the sum of all THREE timing buckets
    # rather than two, so this total is exactly unchanged by the fix
    # above; only which column each rupee is ATTRIBUTED to (below) changes.
    grouped["bank_credit"] = grouped["_bank_in_window"] + grouped["_bank_after_window"] + grouped["_bank_before_window"]
    grouped["difference"] = (grouped["total"] - grouped["bank_credit"]).round(2)

    # Whether THIS UTR's bank credit (whichever bucket(s) it explains)
    # posted after the report window closed - a UTR essentially always
    # has one bank-ledger row (one date), so this is a per-UTR verdict,
    # not something split further per order. A this-period order's own
    # money, or an unresolvable ("not found") ledger row's money, is only
    # ever shown as "settled subsequent" when NOTHING for this UTR posted
    # within the window at all - if some of it already posted in-window,
    # it stays in the "_now" bucket (matches the client's own worked
    # examples where a partially-later-settling UTR still shows its
    # this-period share under "settled this period"). this_period keeps
    # exactly this original two-way split, unchanged by the fix below -
    # a THIS-period order is, by construction, dated inside the current
    # window itself, so its own bank credit predating period_start isn't
    # a real-world case this engine needs to guard against.
    is_after_window = (grouped["_bank_in_window"].abs() <= TOLERANCE) & (grouped["_bank_after_window"].abs() > TOLERANCE)

    grouped["this_period_now"] = grouped["this_period"].where(~is_after_window, 0.0)
    grouped["this_period_subsequent"] = grouped["this_period"].where(is_after_window, 0.0)

    # THE FIX (2026-08-31, client-reported - see this function's own
    # "THE BUG THIS FIXES" docstring note): previous_order and not_found
    # CAN genuinely predate period_start (a previous-period order or an
    # unresolvable ledger row can carry a bank credit from any earlier
    # month), so both now get the full three-way now/subsequent/before
    # split - "before" takes priority only when NEITHER in-window NOR
    # after-window money exists for this UTR, same "only when nothing
    # else claims it" precedence is_after_window above already uses, one
    # level further.
    is_before_window = (
        (grouped["_bank_in_window"].abs() <= TOLERANCE)
        & (grouped["_bank_after_window"].abs() <= TOLERANCE)
        & (grouped["_bank_before_window"].abs() > TOLERANCE)
    )

    grouped["not_found_now"] = grouped["not_found"].where(~is_after_window & ~is_before_window, 0.0)
    grouped["not_found_subsequent"] = grouped["not_found"].where(is_after_window, 0.0)
    grouped["not_found_before"] = grouped["not_found"].where(is_before_window, 0.0)

    grouped["previous_order_now"] = grouped["previous_order"].where(~is_after_window & ~is_before_window, 0.0)
    grouped["previous_order_subsequent"] = grouped["previous_order"].where(is_after_window, 0.0)
    grouped["previous_order_before"] = grouped["previous_order"].where(is_before_window, 0.0)

    # bucket_cols intentionally still reads the FULL (pre-split)
    # "previous_order"/"not_found" sums, not their new *_now/*_before/
    # *_subsequent halves - Remarks/Total/Bank Credit/Difference are a
    # full-history "does this UTR reconcile at all" view, deliberately
    # untouched by the period-attribution fix above (see this function's
    # docstring).
    bucket_cols = [
        "this_period_now", "this_period_subsequent", "subsequent_order", "previous_order",
        "not_found_now", "not_found_subsequent",
    ]

    def _remark_for(row):
        buckets = {k: row[k] for k in bucket_cols}
        return _classify_utr_remark(
            buckets, row["bank_credit"], row["difference"],
            order_ids=row["order_ids"], consolidated_df=consolidated_df,
            match_type=row["_match_type"], matched_bank_utr=row["_matched_bank_utr"], utr=row["utr"],
        )

    grouped["remarks"] = grouped.apply(_remark_for, axis=1)

    for c in bucket_cols + [
        "total", "bank_credit",
        "previous_order_now", "previous_order_subsequent", "previous_order_before", "not_found_before",
    ]:
        grouped[c] = grouped[c].round(2)

    if utr_gateway_lookup is not None:
        grouped["payment_gateway"] = grouped["utr"].map(utr_gateway_lookup).fillna("Unknown")
    else:
        grouped["payment_gateway"] = "Unknown"

    grouped = grouped.rename(columns={
        "utr": "UTR",
        "payment_gateway": "Payment Gateway",
        "this_period_now": "Amount (this period settled this period)",
        "this_period_subsequent": "Amount (this period transaction Settled Subsequent Period)",
        "subsequent_order": "Settled (subsequent period transaction subsequent period)",
        # THE FIX: this display column - the source for Executive
        # Summary Section 7's "Previous Period Amount Received This
        # Period" - is now the window-limited previous_order_now, not the
        # old unsplit previous_order. The plain "previous_order" column
        # itself is intentionally left unrenamed (and so dropped by the
        # grouped[cols] selection below) - it only ever existed to feed
        # bucket_cols/_classify_utr_remark above and the two new
        # before/subsequent columns, never meant to be shown on its own.
        "previous_order_now": "Settled (previous period transaction settled this period)",
        "not_found_now": "Order ID Not Found - Settled During Reco Period",
        "not_found_subsequent": "Order ID Not Found - Settled After Reco Period",
        "total": "Total",
        "bank_credit": "Bank Credit",
        "difference": "Difference",
        "remarks": "Remarks",
        "previous_order_before": "Settled (previous period transaction settled before this period - already reported)",
        "previous_order_subsequent": "Settled (previous period transaction not yet settled as of this period's close)",
        "not_found_before": "Order ID Not Found - Settled Before Reco Period (Already Reported)",
    })

    return grouped[cols].sort_values("UTR").reset_index(drop=True)


def build_refund_utr_detail(consolidated_df):
    """
    Client-reported 2026-08-30 (item 1): "Refund UTR Number and Refund
    Date are not appearing" - build_order_level_utr_detail() below used
    to hardcode both to the literal "NA" for every order, since no
    refund-side matching existed anywhere in this engine. This is that
    missing piece: a per-order refund reference/date table, built the
    same way as the payment-side settlement_ledger (engine.bank.
    build_settlement_ledger) but from the is_refund=True rows instead of
    the payment rows - and, unlike the payment side, this is NOT matched
    against the bank statement at all. The client's own reference
    workbook's Refund UTR values are the gateway/COD partner's OWN
    reported refund reference (e.g. Razorpay's refund settlement_utr,
    Delhivery/Shiprocket COD's own reversal reference), not a value
    independently confirmed against a bank credit line - and it couldn't
    be: load_bank_statement's own amount_col is configured to read CREDIT
    entries only (money coming IN), so a refund (money going back OUT to
    the customer) would never appear as a bank statement row to match
    against in the first place. Surfacing the gateway's own reported
    reference/date is therefore the correct, honest equivalent here -
    not a guessed-at "bank match" that the data can't actually support.

    consolidated_df: engine.consolidator.build_consolidated_receipt()
    output (order_id | source | amount | deduction | is_refund | utr |
    receipt_date | ...) - the same raw per-transaction table
    build_settlement_ledger() itself groups from.

    An order refunded across more than one distinct UTR keeps only its
    LARGEST-refund-amount one, matching the payment side's own "largest
    UTR wins" convention in build_order_level_utr_detail() below.

    Returns a dataframe with columns [order_id, "Refund UTR", "Refund
    Date"] - empty (never None) when consolidated_df has no refund rows
    carrying a resolvable UTR of their own (most refunds don't - that
    remains a disclosed gap, not something this function guesses around).
    """
    cols = ["order_id", "Refund UTR", "Refund Date"]
    if consolidated_df is None or consolidated_df.empty or "utr" not in consolidated_df.columns:
        return pd.DataFrame(columns=cols)

    df = consolidated_df.copy()
    df["order_id"] = df["order_id"].fillna("").astype(str)
    df = df[df["order_id"].str.len() > 0]
    if "is_refund" not in df.columns:
        return pd.DataFrame(columns=cols)
    df = df[df["is_refund"].fillna(False)]
    if df.empty:
        return pd.DataFrame(columns=cols)

    df["utr"] = normalize_utr(df["utr"].fillna(""))
    df = df[(df["utr"].str.len() > 0) & (~df["utr"].isin(["NAN", "NONE", "NAT"]))]
    if df.empty:
        return pd.DataFrame(columns=cols)

    if "receipt_date" not in df.columns:
        df["receipt_date"] = pd.NaT

    grouped = df.groupby(["order_id", "utr"]).agg(
        refund_amount=("amount", "sum"), refund_date=("receipt_date", "max"),
    ).reset_index()

    # Client-reported 2026-08-31 (round 4, point 4): "Refund UTR is
    # available, but the corresponding Refund Date is blank for some
    # UTRs." Root cause identified in the code, not guessed: an order
    # refunded under more than one UTR keeps only its single LARGEST-
    # refund-amount UTR (see docstring above and the dedup two lines
    # below) - but which UTR has the largest amount has no bearing at all
    # on which UTR happens to carry a usable date. If an order's largest
    # refund line has no Settlement Date recorded yet for that specific
    # transaction, while a SMALLER refund line for the very same order
    # does have one, the old code discarded that perfectly good date
    # along with the smaller-amount row - the UTR shown was still the
    # genuinely correct largest one, but its date came out blank even
    # though real refund-date information for that same order existed
    # elsewhere in the data. Fixed by first computing each order's own
    # best-known refund date (the max real date across every UTR that
    # order was refunded under) and backfilling any UTR-level group whose
    # own date is missing with it, BEFORE the largest-amount UTR is
    # picked below - which UTR gets shown is completely unchanged, only
    # the date attached to it can now come from a sibling row.
    #
    # Checked directly against the client's own July data before landing
    # on this fix: every Refund UTR that repeats across multiple orders in
    # this dataset is consistently either always-dated or always-blank
    # (no single UTR shows a date on one order and a blank on another) -
    # so most of the ~68 currently-blank Refund Date rows are not this
    # per-order dedup issue at all, but genuinely reflect the underlying
    # Gokwik settlement file itself having no "Settlement Date" recorded
    # yet for that refund transaction (the refund reference/UTR has been
    # generated, but Gokwik hasn't reported a completion date for it) -
    # the same "reference exists before completion is dated" pattern as
    # point 1's COD remittance-pending fix above, just on the refund side.
    # This fix still closes the specific multi-UTR-per-order gap it can
    # actually see; the remaining single-UTR, genuinely-undated rows are a
    # disclosed data gap, not a matching bug this function can fix without
    # a different raw column to read a date from (see engine/consolidator.py's
    # normalize_gateway_df() - there is currently no way to configure a
    # refund-specific date column separate from a gateway's one shared
    # "date_col").
    order_best_date = grouped.groupby("order_id")["refund_date"].transform("max")
    grouped["refund_date"] = grouped["refund_date"].fillna(order_best_date)

    # Keep only each order's largest-refund-amount UTR - see docstring above.
    grouped = grouped.sort_values("refund_amount", ascending=False)
    grouped = grouped.drop_duplicates(subset="order_id", keep="first")

    return pd.DataFrame({
        "order_id": grouped["order_id"],
        "Refund UTR": grouped["utr"],
        "Refund Date": grouped["refund_date"],
    })


def build_order_level_utr_detail(settlement_ledger_df, utr_bank_reco_df, refund_utr_detail_df=None):
    """
    Client's own "Bank UTR Detail" section (2026-08-27 request) - lets a
    user trace any order/payment transaction on Reco working straight
    through to the bank credit that paid it, without opening a separate
    sheet. Built entirely from data this engine already computes -
    settlement_ledger_df (build_settlement_ledger() above - one row per
    (order_id, UTR) pair an order was actually settled under) joined
    against utr_bank_reco_df (bank_reconciliation_by_utr() above - the
    per-UTR bank match/date/remark, already computed once per run and
    passed straight through here, never recomputed).

    Adds, per order (column names match the client's own "Bank UTR Detail"
    section headers exactly - see the Reco working sheet's "BANK MATCHING"
    group, including their own "Setlment Remarks" spelling):
        Payment UTR | Payment Date (Bank Date) | Setlment Remarks |
        Refund UTR | Refund Date

    An order settled under more than one UTR (e.g. a COD order remitted
    across two separate courier batches, each its own bank credit - see
    build_settlement_ledger's own docstring) keeps only its LARGEST
    (final_payment) UTR here, so a per-order display column stays one
    value per order rather than an ambiguous list - every individual UTR
    for that order is still fully visible, UTR-by-UTR, on the "Bank Reco
    (UTR-wise)" sheet itself; this is a convenience trace-through, not a
    replacement for that sheet.

    Refund UTR / Refund Date (client-reported 2026-08-30, item 1): filled
    in from refund_utr_detail_df (build_refund_utr_detail() above) when
    supplied - see that function's own docstring for what these values
    actually represent (the gateway/courier's own reported refund
    reference, not a bank-matched one). Every order this function's own
    "Bank UTR Detail" table already covers gets a Refund UTR/Date row
    from refund_utr_detail_df if one exists for it; still the literal
    "NA" for any order with no refund, or when refund_utr_detail_df isn't
    supplied at all (keeps this function's own signature backward-
    compatible with any older caller that doesn't pass it).

    Returns a dataframe with "order_id" as a real column (not the index) -
    empty (never None) when either input is empty/missing, so callers can
    always safely .merge() against it without a None-check first.
    """
    cols = ["order_id", "Payment Date (Bank Date)", "Payment UTR",
            "Setlment Remarks", "Refund UTR", "Refund Date"]
    if settlement_ledger_df is None or settlement_ledger_df.empty \
            or utr_bank_reco_df is None or utr_bank_reco_df.empty:
        return pd.DataFrame(columns=cols)

    ledger = settlement_ledger_df.copy()
    ledger["utr"] = normalize_utr(ledger["utr"].fillna(""))
    ledger = ledger[ledger["utr"].str.len() > 0]
    if ledger.empty:
        return pd.DataFrame(columns=cols)

    # Keep only each order's largest-amount UTR - see docstring above.
    ledger = ledger.sort_values("final_payment", ascending=False)
    ledger = ledger.drop_duplicates(subset="order_id", keep="first")

    utr_detail = utr_bank_reco_df[["UTR", "Bank Date", "Remarks"]].copy()
    utr_detail["UTR"] = normalize_utr(utr_detail["UTR"].fillna(""))

    merged = ledger.merge(utr_detail, left_on="utr", right_on="UTR", how="left")
    merged["order_id"] = merged["order_id"].astype(str)

    out = pd.DataFrame({
        "order_id": merged["order_id"],
        "Payment Date (Bank Date)": merged["Bank Date"],
        "Payment UTR": merged["utr"],
        "Setlment Remarks": merged["Remarks"],
        # object dtype, not the default inferred string dtype - "Refund
        # Date" mixes the literal text "NA" with real Timestamps once
        # refund_utr_detail_df is merged in below, which a plain
        # string-inferred column can't hold (pandas 3.0's default str
        # dtype rejects a non-string value on assignment).
        "Refund UTR": pd.Series(["NA"] * len(merged), dtype=object),
        "Refund Date": pd.Series(["NA"] * len(merged), dtype=object),
    })

    if refund_utr_detail_df is not None and not refund_utr_detail_df.empty:
        rlookup = refund_utr_detail_df.drop_duplicates(subset="order_id", keep="first").set_index("order_id")
        has_refund = out["order_id"].isin(rlookup.index)
        out.loc[has_refund, "Refund UTR"] = out.loc[has_refund, "order_id"].map(rlookup["Refund UTR"].to_dict())
        out.loc[has_refund, "Refund Date"] = out.loc[has_refund, "order_id"].map(rlookup["Refund Date"].to_dict())

    return out


# ---------------------------------------------------------------------------
# Order-level reconciliation status (COD-aware, settlement-batch-aware)
# ---------------------------------------------------------------------------
# Added to fix two real misclassifications the client flagged in Order
# Lookup / Bank Reconciliation:
#
#   1. COD orders that were never delivered (RTO / Cancelled / In Transit /
#      Status Undefined / Undelivered / Lost) were showing "Not checked - no
#      bank statement uploaded". That's misleading twice over: no COD
#      collection is ever expected for an order that was never delivered
#      (nothing was collected at the door, so there's nothing for a bank
#      statement to have received), AND the text is simply wrong whenever a
#      bank statement WAS uploaded - what it actually meant was "this order
#      never had a gateway receipt row to check against the bank statement",
#      not "no bank statement uploaded".
#
#   2. Some delivered COD orders that a Shiprocket/Delhivery COD remittance
#      genuinely already settled and that already landed in the bank were
#      still showing "Bank receipt not identified". COD couriers remit
#      collected cash as ONE lump bank credit covering many orders at once;
#      their AWB/settlement export is order-level, but the reference/UTR
#      column on each individual row is frequently blank or an internal
#      reference rather than the bank's own UTR. Matching only order-by-
#      order (match_consolidated_to_bank above) misses these. See
#      build_cod_settlement_batches()/match_batches_to_bank() below for the
#      fallback that catches this by grouping same-gateway,
#      same-settlement-date rows into one batch and matching the BATCH
#      total to a bank credit by amount + date proximity instead of by UTR.
#
# classify_order_bank_status() below folds both fixes, plus the equivalent
# Prepaid-side logic, into six reconciliation categories (confirmed with
# the client):
COD_NOT_DELIVERED = "COD - Not Delivered (No Receipt Expected)"
COD_SETTLEMENT_PENDING = "COD - Delivered & Settlement Pending"
COD_BANK_MATCHED = "COD - Settlement Received & Bank Matched"
PREPAID_SETTLEMENT_PENDING = "Prepaid - Payment Received & Settlement Pending"
PREPAID_BANK_MATCHED = "Prepaid - Settlement Received & Bank Matched"
EXCEPTION_MANUAL_REVIEW = "Exception / Manual Reconciliation Required"
# Deliberately kept OUTSIDE the six categories above (rare, clearly
# separate): a prepaid order where Shopify's own Financial Status says
# payment was never actually received (pending/voided) and no gateway
# settlement row exists either - genuinely nothing to chase yet, unlike a
# real settlement-pending case where the gateway did receive the money.
PREPAID_PAYMENT_NOT_RECEIVED = "Prepaid - Payment Not Yet Received (per order Financial Status)"

DELIVERED_STATUS = "Delivered"
COD_KEYWORDS = ("cod", "cash on delivery", "cash-on-delivery", "cashondelivery")
PAID_FINANCIAL_STATUSES = {"paid", "partially_paid", "partially refunded", "partially_refunded"}
UNPAID_FINANCIAL_STATUSES = {"pending", "voided", "authorized", ""}

# Best-effort defaults, confirmed as reasonable starting points, not a
# guarantee of any specific gateway's/courier's actual SLA - override per
# gateway via "expected_settlement_days" in the client config if the firm's
# actual agreement with that gateway/courier is known to differ (see
# configs/esca_shopify.json). Days pending beyond these windows moves an
# order from "still normal, settlement/bank credit pending" into
# "Exception / Manual Reconciliation Required" - i.e. these thresholds
# decide when the tool stops assuming "probably still in transit" and asks
# a human to look, not a claim about when money legally must arrive.
DEFAULT_SETTLEMENT_GENERATION_GRACE_DAYS = 20  # delivered COD -> appears in settlement file
DEFAULT_BANK_CREDIT_GRACE_DAYS = 10            # COD settlement generated -> bank credit
DEFAULT_PREPAID_BANK_CREDIT_GRACE_DAYS = 7     # prepaid gateway settled -> bank credit


def classify_payment_type(payment_method_text, financial_status_text=None, has_cod_settlement_row=False):
    """
    COD / Prepaid / Unknown, from Shopify's own Payment Method text.

    has_cod_settlement_row (2026-08-25 client-reported fix): whether this
    order already has a receipt row from a COD-mode gateway (Delhivery
    COD / Shiprocket COD / Prozo COD - i.e. a real courier remittance
    file actually reports this order). That's a FACTUAL signal straight
    from the gateway config, not a guess, so it wins outright over
    anything the Payment Method text says.

    Root cause this exists at all: the client's real July export leaves
    Payment Method BLANK for every genuine COD order (2,231 of 7,289
    rows) - this store's own Shopify integration apparently never
    populates it for COD, only for prepaid gateways ("Gokwik UPI",
    "01 Cards, UPI, NB, Wallets by Razorpay", ...). Before this fix, a
    blank/unrecognized Payment Method fell through to "Unknown", which
    is "handled like Prepaid downstream" - so ~1,700+ genuinely-COD
    orders that already had a Delhivery/Shiprocket COD settlement row
    were being classified as "Prepaid - ..." instead of "COD - ...",
    which is why the Gateway Settlement sheet's COD "Pending for
    Settlement" showed 0.00 despite the client's own Settlement Pending
    Summary correctly showing real Delhivery COD / Shiprocket COD
    pending amounts - the two sheets were silently using two different
    payment-type signals. Verified against the client's own July data:
    100% of the 2,231 blank-Payment-Method rows have Financial Status
    "pending" (vs. only 19 of the other 5,058 rows), so - failing the
    factual settlement-row check above - a blank/unrecognized Payment
    Method paired with an unpaid-looking Financial Status is now treated
    as COD too, rather than "Unknown"-as-Prepaid. Still called out via
    the "Unknown" path if financial_status doesn't back that up either -
    a fallback guess, not a confirmed fact.
    """
    if has_cod_settlement_row:
        return "COD"
    text = str(payment_method_text or "").strip().lower()
    if text and text not in ("nan", "none"):
        if any(k in text for k in COD_KEYWORDS):
            return "COD"
        return "Prepaid"
    fin = str(financial_status_text or "").strip().lower()
    if fin in UNPAID_FINANCIAL_STATUSES:
        return "COD"
    return "Unknown"


def build_cod_settlement_batches(consolidated_df, gateway_configs, bank_ledger_df=None):
    """
    Groups each COD gateway's settlement rows into "settlement batches" -
    same gateway (source) + same settlement/remittance date - and totals
    the net payout for that batch (amount minus the courier's own
    freight/fee deductions - the same "final payment" concept used
    elsewhere in this engine).

    Only COD gateways are grouped this way (payment_mode == "COD" in the
    client config). Prepaid gateways generally already carry a clean
    per-transaction settlement UTR, so batching them by date would risk
    merging together settlements that just happen to land on the same day -
    a much weaker signal for a rail that already settles cleanly order by
    order.

    When bank_ledger_df is supplied, rows whose OWN order-level UTR already
    matched a bank credit (match_consolidated_to_bank) are excluded before
    batching - otherwise an order that's already been explained by a direct
    UTR match would also get counted into a batch total, which could either
    double-count that money or make a real batch match fail because the
    total no longer equals any single bank credit (a courier's remittance
    for one date is one bank credit for the WHOLE batch, not the sum of a
    resolved order plus an unresolved one). Pass None to batch every COD
    row regardless of UTR match status (used where the caller hasn't
    already resolved bank matching, or explicitly wants every row batched).

    Returns: batch_id | source | settlement_date | order_ids | batch_amount
    """
    cols = ["batch_id", "source", "settlement_date", "order_ids", "batch_amount"]
    if consolidated_df is None or consolidated_df.empty:
        return pd.DataFrame(columns=cols)

    cod_labels = {cfg["label"] for cfg in gateway_configs if str(cfg.get("payment_mode", "")).lower() == "cod"}
    if not cod_labels:
        return pd.DataFrame(columns=cols)

    df = consolidated_df[(~consolidated_df["is_refund"]) & (consolidated_df["source"].isin(cod_labels))].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    if bank_ledger_df is not None and not bank_ledger_df.empty:
        with_bank = match_consolidated_to_bank(df, bank_ledger_df)
        df = with_bank[~with_bank["bank_matched"]].copy()
        if df.empty:
            return pd.DataFrame(columns=cols)

    df["_net"] = df["amount"] - df.get("deduction", 0.0)
    df["_settle_date"] = pd.to_datetime(df.get("receipt_date"), errors="coerce").dt.date

    grouped = df.groupby(["source", "_settle_date"], dropna=False).agg(
        batch_amount=("_net", "sum"),
        order_ids=("order_id", lambda s: sorted(set(s))),
    ).reset_index()
    grouped = grouped.rename(columns={"_settle_date": "settlement_date"})
    grouped["batch_id"] = grouped["source"].astype(str) + " | " + grouped["settlement_date"].astype(str)
    grouped["batch_amount"] = grouped["batch_amount"].round(2)

    return grouped[cols]


def match_batches_to_bank(batches_df, bank_ledger_df, date_window_days=5, exclude_utrs=None):
    """
    Matches each COD settlement batch (build_cod_settlement_batches() above)
    to a bank credit by amount + date proximity, since the batch has no UTR
    of its own to look up directly.

    A batch matches a bank credit when the bank credit's amount is within
    TOLERANCE (₹1) of the batch total AND its date falls within
    `date_window_days` of the batch's own settlement/remittance date
    (remittances commonly land in the bank a few business days after the
    courier's settlement-date cutoff, so requiring the exact same day is
    too strict). If more than one bank credit in the window matches the
    amount, the closest by date wins, and that ambiguity is called out in
    the match note so it still gets a human look.

    exclude_utrs: an optional set of normalized UTRs to remove from the
    bank ledger before searching - pass the UTRs already claimed by a
    direct order-level match (see classify_order_bank_status's
    `already_matched_utrs`) so this amount/date fallback can never re-match
    a bank credit that's already accounted for by a real reference-number
    match. Without this, the SAME bank credit could satisfy both a direct
    UTR match and an unrelated batch's amount/date match, double-counting
    that one credit's money as if it were received twice.

    This is inherently a heuristic (amount + date, not a hard reference) -
    exactly like the rest of this module's disclosed-rule approach (see
    module docstring): a strong lead for the reconciling accountant, not a
    substitute for tracing every rupee by reference number.

    A batch with NO settlement date at all (client-reported 2026-08-31,
    round 4, point 1 - see the inline comment in the loop below) is always
    left unmatched, never matched by amount alone with no date window - see
    that comment for the real order this was caught against.

    Returns: batch_id | matched | bank_amount | bank_date | match_note
    """
    cols = ["batch_id", "matched", "bank_amount", "bank_date", "match_note"]
    if batches_df is None or batches_df.empty:
        return pd.DataFrame(columns=cols)
    if bank_ledger_df is None or bank_ledger_df.empty:
        out = batches_df[["batch_id"]].copy()
        out["matched"] = False
        out["bank_amount"] = None
        out["bank_date"] = pd.NaT
        out["match_note"] = "No bank statement uploaded - settlement-batch matching not attempted"
        return out[cols]

    bank = bank_ledger_df.copy()
    # 2026-08-31 fix: use the same hardened helpers
    # bank_reconciliation_by_utr() relies on (see their own docstrings)
    # rather than a bare pd.to_datetime()/implicit comparison - a
    # resolution mismatch between this column and settle_date below (both
    # parsed independently) can otherwise raise "Invalid comparison
    # between dtype=datetime64[us] and Timestamp" on some pandas builds,
    # exactly as client-reported against views/page_reports.py's own new
    # Point 7 code this same day.
    bank["bank_date"] = to_naive_datetime_series(bank["bank_date"])
    if exclude_utrs:
        bank = bank[~bank["utr"].isin(exclude_utrs)]
    used_bank_rows = set()

    rows = []
    for _, batch in batches_df.iterrows():
        settle_date = to_naive_timestamp(batch["settlement_date"])
        # Client-reported 2026-08-31 (round 4, point 1): Shiprocket order
        # 33160 was showing as Bank Credit even though the client confirmed
        # its remittance is still genuinely pending. Root cause: when a
        # batch's own settlement_date is NaT (build_cod_settlement_batches()
        # groups by the courier's OWN remittance/settlement date - a CRF/
        # remittance batch that hasn't actually been remitted yet has no
        # such date to report), the `if pd.notna(settle_date):` guard below
        # used to skip the date-window filter ENTIRELY rather than skipping
        # the match - leaving every not-yet-dated batch to be matched by
        # AMOUNT ALONE against the whole bank ledger, with no date anchor
        # at all. That's exactly backwards: a batch with no settlement date
        # is the clearest possible signal that nothing has been remitted
        # yet, not a reason to relax the matching criteria. A batch with no
        # settlement date is now always left unmatched here (falls through
        # to COD_SETTLEMENT_PENDING/EXCEPTION_MANUAL_REVIEW in
        # classify_order_bank_status(), never a false COD_BANK_MATCHED) -
        # batches that DO carry a real settlement date are completely
        # unaffected, still matched by amount + date proximity exactly as
        # before.
        if pd.isna(settle_date):
            rows.append({
                "batch_id": batch["batch_id"], "matched": False,
                "bank_amount": None, "bank_date": pd.NaT,
                "match_note": (
                    "No settlement/remittance date recorded for this batch yet (the courier "
                    "hasn't reported one - a strong sign the remittance itself hasn't happened) - "
                    "matching by amount alone with no date anchor is unreliable, so this batch is "
                    "left unmatched rather than guessed."
                ),
            })
            continue

        candidates = bank[(bank["bank_amount"] - batch["batch_amount"]).abs() <= TOLERANCE]
        candidates = candidates[(candidates["bank_date"] - settle_date).abs() <= pd.Timedelta(days=date_window_days)]
        candidates = candidates[~candidates.index.isin(used_bank_rows)]

        if candidates.empty:
            rows.append({
                "batch_id": batch["batch_id"], "matched": False,
                "bank_amount": None, "bank_date": pd.NaT,
                "match_note": "No bank credit found matching this batch's amount within "
                              f"±{date_window_days} days of its settlement date - needs manual check",
            })
            continue

        if pd.notna(settle_date):
            gaps = (candidates["bank_date"] - settle_date).abs()
            best_idx = gaps.idxmin()
        else:
            best_idx = candidates.index[0]

        used_bank_rows.add(best_idx)
        note = "Matched via settlement-batch amount/date (not by UTR) - confirm manually"
        if len(candidates) > 1:
            note += "; more than one same-amount bank credit fell in this window, closest date picked"
        rows.append({
            "batch_id": batch["batch_id"], "matched": True,
            "bank_amount": bank.loc[best_idx, "bank_amount"],
            "bank_date": bank.loc[best_idx, "bank_date"],
            "match_note": note,
        })

    return pd.DataFrame(rows)[cols]


def classify_order_bank_status(
    reco_df, consolidated_df, bank_ledger_df, gateway_configs, bank_statement_uploaded,
    as_of_date=None,
    settlement_generation_grace_days=DEFAULT_SETTLEMENT_GENERATION_GRACE_DAYS,
    bank_credit_grace_days=DEFAULT_BANK_CREDIT_GRACE_DAYS,
    prepaid_bank_credit_grace_days=DEFAULT_PREPAID_BANK_CREDIT_GRACE_DAYS,
):
    """
    The full per-order reconciliation-category classifier described in the
    module-section docstring above. One row per order in reco_df, with:
        order_id | payment_type | has_settlement_row | receipt_amount |
        receipt_date | bank_matched | bank_match_method | bank_credit_date |
        days_pending | Reconciliation Category | Category Reason

    bank_statement_uploaded must be passed explicitly (rather than inferred
    from bank_ledger_df being empty) so a bank statement that was uploaded
    but happened to extract zero usable UTRs is still treated as "checked,
    found nothing" - not silently equivalent to "never uploaded".
    """
    cols = [
        "order_id", "payment_type", "has_settlement_row", "receipt_amount",
        "receipt_date", "bank_matched", "bank_match_method", "bank_credit_date",
        "days_pending", "Reconciliation Category", "Category Reason",
    ]
    if reco_df is None or reco_df.empty:
        return pd.DataFrame(columns=cols)

    as_of_date = to_naive_timestamp(as_of_date if as_of_date is not None else pd.Timestamp.now())
    gateway_configs = gateway_configs or []

    if consolidated_df is not None and not consolidated_df.empty:
        consolidated_with_bank = match_consolidated_to_bank(consolidated_df, bank_ledger_df)
        payments = consolidated_with_bank[~consolidated_with_bank["is_refund"]].copy()
    else:
        payments = pd.DataFrame(columns=["order_id", "amount", "bank_matched", "bank_date", "source", "receipt_date"])

    if len(payments):
        # .to_dict() up front, not left as Series: mapping an EMPTY
        # datetime64-typed Series via .map() crashes under some pandas
        # versions/dtype backends (TypeError: Cannot cast DatetimeArray to
        # dtype float64) - hit for real whenever no gateway/COD file has
        # been uploaded yet at all (consolidated_df empty), which is a
        # normal, non-blocked partial run, not an edge case. Plain dicts
        # sidestep that pandas dtype-inference path entirely.
        has_settlement = payments.groupby("order_id").size().to_dict()
        order_utr_matched = payments.groupby("order_id")["bank_matched"].any().to_dict()
        # Whether any of this order's matched rows relied on the UTR-suffix
        # fallback (see match_consolidated_to_bank/match_utr_against_bank
        # above) rather than an exact reference match - surfaced in
        # bank_match_method below so the "confirm manually" caveat that
        # already accompanies every heuristic match in this module travels
        # with it here too.
        order_match_type = (
            payments[payments["bank_matched"]].groupby("order_id")["bank_match_type"]
            .agg(lambda s: "suffix" if (s == "suffix").any() else "exact").to_dict()
            if "bank_match_type" in payments.columns else {}
        )
        order_bank_date = payments.groupby("order_id")["bank_date"].min().to_dict()
        order_receipt_amt = payments.groupby("order_id")["amount"].sum().to_dict()
        order_source = payments.groupby("order_id")["source"].agg(lambda s: ", ".join(sorted(set(s)))).to_dict()
        order_receipt_date = (
            payments.groupby("order_id")["receipt_date"].min().to_dict()
            if "receipt_date" in payments.columns else {}
        )
    else:
        has_settlement = {}
        order_utr_matched = {}
        order_match_type = {}
        order_bank_date = {}
        order_receipt_amt = {}
        order_source = {}
        order_receipt_date = {}

    # Settlement-batch fallback match (COD only) - see functions above.
    # Bank credits already claimed by a direct order-level UTR match are
    # excluded from the batch search via matched_order_level_utrs() -
    # otherwise the SAME bank credit could satisfy both a direct UTR match
    # and an unrelated batch's amount/date match, double-counting that one
    # credit's money across two different orders/gateways.
    already_matched_utrs = matched_order_level_utrs(consolidated_df, bank_ledger_df)
    batches = build_cod_settlement_batches(consolidated_df, gateway_configs, bank_ledger_df)
    batch_match = match_batches_to_bank(batches, bank_ledger_df, exclude_utrs=already_matched_utrs)
    order_to_batch_matched, order_to_batch_note, order_to_batch_date = {}, {}, {}
    if not batches.empty and not batch_match.empty:
        batch_lookup = batch_match.set_index("batch_id")
        for _, b in batches.iterrows():
            info = batch_lookup.loc[b["batch_id"]]
            if bool(info["matched"]):
                for oid in b["order_ids"]:
                    order_to_batch_matched[oid] = True
                    order_to_batch_note[oid] = info["match_note"]
                    order_to_batch_date[oid] = info["bank_date"]

    gateway_grace = {
        cfg["label"]: cfg["expected_settlement_days"]
        for cfg in gateway_configs if cfg.get("expected_settlement_days") is not None
    }

    df = pd.DataFrame({
        "order_id": reco_df["order_id"].astype(str),
        "final_delivery_status": reco_df.get("final_delivery_status"),
        "created_at": reco_df.get("created_at"),
        "delivered_date": reco_df.get("delivered_date"),
        "payment_method": reco_df.get("payment_method") if "payment_method" in reco_df.columns else None,
        "financial_status": reco_df.get("financial_status") if "financial_status" in reco_df.columns else None,
    })

    # The factual "does this order already have a receipt row from a
    # COD-mode gateway" signal classify_payment_type needs (see its own
    # docstring) - built from order_source (already computed above) plus
    # each gateway config's own declared payment_mode, so it reflects
    # exactly what gateway_configs says COD means, not a second guess.
    gateway_payment_mode = {
        cfg["label"]: str(cfg.get("payment_mode", "")).strip().lower() for cfg in gateway_configs
    }

    def _has_cod_settlement_row(source_text):
        return any(
            gateway_payment_mode.get(s.strip()) == "cod"
            for s in str(source_text or "").split(",") if s.strip()
        )

    df["_source"] = df["order_id"].map(order_source)
    df["_has_cod_settlement_row"] = df["_source"].apply(_has_cod_settlement_row)
    df["payment_type"] = df.apply(
        lambda r: classify_payment_type(r["payment_method"], r["financial_status"], r["_has_cod_settlement_row"]),
        axis=1,
    )
    df["has_settlement_row"] = df["order_id"].map(has_settlement).fillna(0).gt(0)
    df["receipt_amount"] = df["order_id"].map(order_receipt_amt).fillna(0.0)
    df["receipt_date"] = df["order_id"].map(order_receipt_date)
    df["_order_utr_matched"] = df["order_id"].map(order_utr_matched).fillna(False)
    df["_batch_matched"] = df["order_id"].map(order_to_batch_matched).fillna(False)
    df["bank_matched"] = df["_order_utr_matched"] | df["_batch_matched"]
    df["_utr_bank_date"] = df["order_id"].map(order_bank_date)
    df["_batch_bank_date"] = df["order_id"].map(order_to_batch_date)
    df["bank_credit_date"] = df["_utr_bank_date"].where(df["_order_utr_matched"], df["_batch_bank_date"])
    df["_order_match_type"] = df["order_id"].map(order_match_type)
    df["bank_match_method"] = None
    df.loc[df["_order_utr_matched"], "bank_match_method"] = "Matched by order UTR"
    df.loc[df["_order_utr_matched"] & (df["_order_match_type"] == "suffix"), "bank_match_method"] = (
        "Matched by order UTR (bank narration reference format differs from the gateway's UTR - confirm manually)"
    )
    df.loc[(~df["_order_utr_matched"]) & df["_batch_matched"], "bank_match_method"] = df["order_id"].map(order_to_batch_note)
    # _source was already computed above (needed early for payment_type's
    # has_cod_settlement_row signal) - not reassigned here again.

    def _grace_for(source_text, default):
        sources = [s.strip() for s in str(source_text or "").split(",") if s.strip()]
        graces = [gateway_grace[s] for s in sources if s in gateway_grace]
        return min(graces) if graces else default

    def classify(row):
        # Both parsed fresh per-row here (created_at/delivered_date are
        # still raw/unparsed at this point - see engine/reco.py), then
        # immediately stripped of any timezone via to_naive_timestamp() -
        # Shopify's own "Created at" export commonly carries a timezone
        # offset, which would otherwise crash every subtraction below
        # against the naive as_of_date/pd.Timestamp.now() used throughout
        # this function.
        ref_date = row["delivered_date"] if pd.notna(row["delivered_date"]) else row["created_at"]
        ref_date = to_naive_timestamp(pd.to_datetime(ref_date, errors="coerce"))
        settle_ref_date = to_naive_timestamp(pd.to_datetime(row["receipt_date"], errors="coerce"))
        days_since_ref = (as_of_date - ref_date).days if pd.notna(ref_date) else None
        days_since_settlement = (as_of_date - settle_ref_date).days if pd.notna(settle_ref_date) else None
        pending_clock = days_since_settlement if row["has_settlement_row"] else days_since_ref

        if row["payment_type"] == "COD":
            # Client-reported 2026-08-30 (item 14, the residual ~₹1,657
            # bank-matching gap): this used to fire for EVERY non-Delivered
            # COD order, before even checking has_settlement_row - so an
            # order the courier's own COD settlement/remittance file
            # already lists (real money genuinely collected/reported,
            # confirmed by a non-zero receipt_amount) got waved away as
            # "no receipt expected" purely because final_delivery_status
            # said Cancelled/RTO/etc, even though the courier's own report
            # said otherwise. Root-caused against the client's own July
            # data: order 32006 (Cancelled per Delhivery's status feed,
            # but Delhivery's own COD settlement file already shows
            # receipt_amount == order value) - the client's own workbook
            # correctly still calls this "Delhivery COD Setlment pending",
            # not "no receipt expected". Fixed by requiring has_settlement_
            # row to ALSO be false before taking this shortcut - a
            # genuinely not-yet-collected not-delivered order (the
            # overwhelmingly common case) is unaffected; only the rarer
            # case of real, already-reported money on a not-yet-"Delivered"
            # order now correctly falls through to the ordinary has_
            # settlement_row/bank_matched logic below instead.
            if row["final_delivery_status"] != DELIVERED_STATUS and not row["has_settlement_row"]:
                return {
                    "Reconciliation Category": COD_NOT_DELIVERED,
                    "Category Reason": (
                        f"Delivery status is '{row['final_delivery_status']}' - no COD collection happens "
                        "unless the order is delivered, so no bank receipt is expected for this order."
                    ),
                    "days_pending": None,
                }

            if not row["has_settlement_row"]:
                overdue = ""
                if days_since_ref is not None and days_since_ref > settlement_generation_grace_days:
                    overdue = (
                        f" - {days_since_ref} days since delivery, beyond the usual "
                        f"{settlement_generation_grace_days}-day settlement window, may need follow-up with the courier"
                    )
                return {
                    "Reconciliation Category": COD_SETTLEMENT_PENDING,
                    "Category Reason": (
                        "Delivered, but this order hasn't appeared in the COD settlement/remittance "
                        f"file yet{overdue}."
                    ),
                    "days_pending": days_since_ref,
                }

            if row["bank_matched"]:
                return {
                    "Reconciliation Category": COD_BANK_MATCHED,
                    "Category Reason": f"COD settlement traced to a bank credit ({row['bank_match_method']}).",
                    "days_pending": pending_clock,
                }

            if not bank_statement_uploaded:
                return {
                    "Reconciliation Category": COD_SETTLEMENT_PENDING,
                    "Category Reason": (
                        "COD settlement/remittance recorded, but no bank statement was uploaded this run "
                        "so bank-matching wasn't attempted."
                    ),
                    "days_pending": pending_clock,
                }

            grace = _grace_for(row["_source"], bank_credit_grace_days)
            if days_since_settlement is not None and days_since_settlement > grace:
                settle_txt = settle_ref_date.date() if pd.notna(settle_ref_date) else "an unknown date"
                return {
                    "Reconciliation Category": EXCEPTION_MANUAL_REVIEW,
                    "Category Reason": (
                        f"COD settlement recorded on {settle_txt}, but no matching bank credit found after "
                        f"{days_since_settlement} days (expected within ~{grace}) - not matched by order UTR "
                        "or by settlement-batch amount/date either. Needs manual reconciliation."
                    ),
                    "days_pending": pending_clock,
                }
            return {
                "Reconciliation Category": COD_SETTLEMENT_PENDING,
                "Category Reason": (
                    f"COD settlement recorded; bank credit not yet identified, still within the usual "
                    f"~{grace}-day window."
                ),
                "days_pending": pending_clock,
            }

        # Prepaid / Unknown payment method
        note_prefix = "" if row["payment_type"] == "Prepaid" else (
            "(Payment method not available on the order - treated as Prepaid for this check.) "
        )
        fin_status = str(row["financial_status"] or "").strip().lower()
        payment_received_signal = row["has_settlement_row"] or fin_status in PAID_FINANCIAL_STATUSES

        if not payment_received_signal and fin_status in UNPAID_FINANCIAL_STATUSES:
            return {
                "Reconciliation Category": PREPAID_PAYMENT_NOT_RECEIVED,
                "Category Reason": (
                    note_prefix + f"Order Financial Status is '{row['financial_status']}' and no gateway "
                    "settlement row exists yet - nothing to reconcile until payment is actually received."
                ),
                "days_pending": None,
            }

        if not row["has_settlement_row"]:
            return {
                "Reconciliation Category": PREPAID_SETTLEMENT_PENDING,
                "Category Reason": (
                    note_prefix + f"Financial Status is '{row['financial_status']}' (payment received by the "
                    "gateway), but it hasn't appeared in the gateway's settlement file yet."
                ),
                "days_pending": days_since_ref,
            }

        if row["bank_matched"]:
            return {
                "Reconciliation Category": PREPAID_BANK_MATCHED,
                "Category Reason": note_prefix + f"Gateway settlement traced to a bank credit ({row['bank_match_method']}).",
                "days_pending": pending_clock,
            }

        if not bank_statement_uploaded:
            return {
                "Reconciliation Category": PREPAID_SETTLEMENT_PENDING,
                "Category Reason": (
                    note_prefix + "Gateway settlement recorded, but no bank statement was uploaded this run "
                    "so bank-matching wasn't attempted."
                ),
                "days_pending": pending_clock,
            }

        grace = _grace_for(row["_source"], prepaid_bank_credit_grace_days)
        if days_since_settlement is not None and days_since_settlement > grace:
            settle_txt = settle_ref_date.date() if pd.notna(settle_ref_date) else "an unknown date"
            return {
                "Reconciliation Category": EXCEPTION_MANUAL_REVIEW,
                "Category Reason": (
                    note_prefix + f"Gateway settlement recorded on {settle_txt}, but no matching bank credit "
                    f"found after {days_since_settlement} days (expected within ~{grace}). Needs manual reconciliation."
                ),
                "days_pending": pending_clock,
            }
        return {
            "Reconciliation Category": PREPAID_SETTLEMENT_PENDING,
            "Category Reason": (
                note_prefix + "Gateway settlement recorded; bank credit not yet identified, still within the "
                f"usual ~{grace}-day window."
            ),
            "days_pending": pending_clock,
        }

    classified = pd.DataFrame(df.apply(classify, axis=1).tolist(), index=df.index)
    df["Reconciliation Category"] = classified["Reconciliation Category"]
    df["Category Reason"] = classified["Category Reason"]
    df["days_pending"] = classified["days_pending"]

    return df[cols]
