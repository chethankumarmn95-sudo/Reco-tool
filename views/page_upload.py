"""
page_upload.py
--------------
Upload Data page. Step 1 is always "which platform is this data for" -
Shopify, Amazon, Flipkart, Tata 1mg, First Cry - before any file uploaders
show up. That answer picks the matching client/channel config automatically
(same one Settings switches between), so there's no separate trip to
Settings needed just to get started. Running the actual reconciliation
happens on the separate Reconciliation page - splitting these two keeps
this page focused on "get the right files in."
"""

import io
import pandas as pd
import streamlit as st

from engine.loaders import load_table_smart, resolve_col
from engine.validation import validate_file_matches_source, friendly_mismatch_message, _signature_columns
from engine import dedup, storage
from engine.dedup import AWB_COLUMN_ALIASES as _AWB_COLUMN_ALIASES
from engine.raw_transforms import RAW_TRANSFORMS, TransformPrerequisiteError
from engine.report_detection import build_report_registry, wrong_report_message
from views.state_init import set_active_config

# Platforms offered on the "which platform is this?" step. A platform shows
# up here even before a config exists for it - it just shows a "not set up
# yet" notice instead of upload boxes until someone adds a configs/*.json
# for it (see Settings page docstring for how).
PLATFORMS = ["Shopify", "Amazon", "Flipkart", "Tata 1mg", "First Cry"]

_UPLOAD_STATE_KEYS = (
    "reco_df", "lookup_df", "totals", "consolidated", "bank_status_df",
    "orders_df", "delivery_frames", "gateway_frames", "bank_df",
    # attribution_frames - optional, informational-only Payment Gateway
    # Attribution uploads (Gokwik Order Report / Gokwik Transaction
    # Report - 2026-08-25 client request, see engine/attribution.py) -
    # dict-shaped like delivery_frames/gateway_frames (keyed by each
    # attribution_sources config entry's own label), reset/restored the
    # same way.
    "attribution_frames",
    "gateway_settlement_df", "utr_bank_reco_df", "bank_ledger_df",
    "recon_status_df", "settlement_pending_df", "settlement_pending_summary_df",
    "gateway_health_df", "cod_batches_df", "cod_batch_match_df", "bank_statement_uploaded",
    # Marketplace-channel (Amazon and future Flipkart/Tata1mg/First Cry)
    # upload slots - separate from the DTC/Shopify-shaped keys above since
    # a marketplace channel's raw inputs (MTR reports, settlement flat
    # files) don't map onto "orders/delivery_partners/gateways" at all.
    "mtr_files", "settlement_files",
    # Marketplace-channel LIVE reconciliation results (set by
    # views/page_reconciliation.py's _render_marketplace / restored by
    # views/page_data_management.py's _render_amazon "Load selected") -
    # cleared here too so switching platforms/clients mid-session can't
    # leak one channel's Dashboard/Reports/Bank Linking figures into
    # another's, same reasoning as every other key in this tuple.
    "amazon_order_reco_df", "amazon_waterfall_df", "amazon_expense_ledger_df",
    "amazon_settlement_summary_df", "amazon_settlement_register_df",
    "amazon_tie_out_df", "amazon_non_mtr_df",
    "amazon_cutoff_date", "amazon_subsequent_settlements_df",
)


@st.cache_data(show_spinner=False)
def _load_bytes_smart(file_bytes, filename, expected_signature):
    buf = io.BytesIO(file_bytes)
    buf.name = filename
    df, header_row_used = load_table_smart(buf, expected_signature=expected_signature)
    return df, header_row_used


def load_and_validate(uploaded_file, source_cfg, registry=None):
    if uploaded_file is None:
        return None
    expected_signature = _signature_columns(source_cfg)
    df, header_row_used = _load_bytes_smart(
        uploaded_file.getvalue(), uploaded_file.name, expected_signature
    )
    if header_row_used is None:
        is_valid, missing = validate_file_matches_source(df, source_cfg)
        # Data-driven "is this actually some OTHER report" check (see
        # engine.report_detection) - runs before falling back to a plain
        # "missing columns" message, so a Shiprocket file dropped into a
        # Delhivery box gets told exactly that, by name, instead of a
        # confusing list of missing Delhivery columns. registry is None
        # for callers that haven't built one (defensive default - the
        # cross-report check simply doesn't run, this source's own
        # missing-columns message still does).
        if registry:
            st.error(wrong_report_message(source_cfg["label"], df, registry, missing))
        else:
            st.error(
                friendly_mismatch_message(source_cfg)
                + f" Missing expected column(s): {', '.join(str(m) for m in missing)}. "
            )
        st.caption(
            "I also tried reading with the header a few rows further down in case an "
            "extra row was inserted above it, but couldn't find a match either way."
        )
        return None
    if header_row_used > 0:
        st.info(
            f"Heads up: {source_cfg['label']} had {header_row_used} extra row(s) "
            f"above the actual column headers - skipped past them automatically."
        )
    return df


@st.cache_data(show_spinner=False)
def _read_all_sheets(file_bytes, filename):
    """Reads EVERY sheet of an uploaded workbook (not just sheet 0) into a
    {sheet_name: DataFrame} dict - needed for sources like Shiprocket COD
    whose raw export has to be auto-mapped by combining more than one
    sheet (see engine.shiprocket_cod). A .csv upload has no concept of
    multiple sheets, so it's wrapped as a single-entry dict keyed "Sheet1"
    - the transform function itself decides whether that's enough (a
    source that genuinely needs a second sheet will raise a specific,
    named error rather than silently proceeding).

    2026-08-31 fix (client-reported: Shiprocket's own COD Remittance
    export only ever comes as a genuine legacy "Excel 97-2003" .xls file,
    which the Upload Data page previously couldn't accept at all - the
    file type simply wasn't offered). Two distinct problems needed fixing,
    not just the file picker:
      1. openpyxl (used for .xlsx) can't read .xls at all - a different
         reader is required for the legacy binary format.
      2. pandas' own default reader for .xls is xlrd, but a REAL client
         file failed even there with a low-level "Workbook corruption"
         error from xlrd's own compound-document parser - not a corrupt
         file (LibreOffice opens it fine), just a legacy export written by
         whatever backend Shiprocket's dashboard uses, apparently not
         byte-for-byte identical to what Microsoft Excel itself would
         produce, in a way xlrd's stricter parser rejects.
    python-calamine (a Rust-based reader, deliberately more tolerant of
    exactly this kind of real-world quirky-but-valid file) reads the same
    client file correctly - confirmed directly against it - so .xls now
    tries calamine first and only falls back to xlrd (pandas' original
    default) if calamine can't be used for some reason, rather than
    failing outright. .xlsx/.csv are completely unaffected - unchanged
    from before this fix.
    """
    buf = io.BytesIO(file_bytes)
    if filename.lower().endswith(".csv"):
        return {"Sheet1": pd.read_csv(buf)}
    if filename.lower().endswith(".xls"):
        try:
            return pd.read_excel(buf, sheet_name=None, engine="calamine")
        except Exception:
            buf.seek(0)
            return pd.read_excel(buf, sheet_name=None, engine="xlrd")
    return pd.read_excel(buf, sheet_name=None)


def _load_gateway_file(f, g_cfg, registry=None):
    """
    Like load_and_validate() below, but for gateway configs that declare
    a "raw_transform" (currently just Shiprocket COD - see
    engine.shiprocket_cod) - i.e. sources where the file the user actually
    has (Shiprocket's own raw COD Remittance export) isn't just a
    differently-named version of what the engine expects, it's genuinely
    missing columns that have to be COMPUTED from a second sheet before
    the file can be used at all.

    The client's own framing: the "mapped" file the tool used to require
    is the reference/sample of the TARGET shape, not the file the user
    should have to build by hand and upload - this auto-builds that same
    target shape from whatever raw file Shiprocket actually hands out.
    Which sheet plays which role is identified purely by column headers
    (see engine.shiprocket_cod's AWB_SIGNATURE/CRF_SIGNATURE) - sheet
    NAME, capitalisation/spacing, and sheet ORDER are never used to
    decide, so a platform renaming a sheet (or swapping the two sheets'
    order) never breaks this on its own.

    Falls back to the plain load_and_validate() path for every gateway
    that doesn't declare a raw_transform, so nothing changes for any
    other source.

    Never surfaces a bare "Invalid File" for a raw_transform source -
    every failure here (no sheet with the right columns, missing column
    within the sheet that WAS found, still-missing signature column even
    after mapping, or the file turning out to be some OTHER known report
    entirely - see engine.report_detection) names exactly what's wrong,
    per the client's explicit ask.
    """
    transform_key = g_cfg.get("raw_transform")
    if not transform_key or f is None:
        return load_and_validate(f, g_cfg, registry=registry)

    transform_fn = RAW_TRANSFORMS.get(transform_key)
    if transform_fn is None:
        # Config references a transform key this version of the app
        # doesn't know about - a config/code mismatch, not a bad upload.
        st.error(
            f"**{g_cfg['label']}** is configured with an unknown raw_transform "
            f"\"{transform_key}\" - please check configs/*.json against the app version."
        )
        return None

    try:
        sheets = _read_all_sheets(f.getvalue(), f.name)
        # context carries whatever OTHER already-uploaded data a transform
        # might need beyond its own file's sheets - today just the Shopify
        # order report, needed by Razorpay's mapping (see
        # engine.razorpay_settlement). A transform that doesn't need any
        # of it just ignores the argument - see engine.raw_transforms.
        df = transform_fn(sheets, context={"orders_df": st.session_state.get("orders_df")})
    except TransformPrerequisiteError as exc:
        # Client-reported (2026-08-21): this failure has NOTHING to do with
        # whether the uploaded file is the right report - e.g. a genuine
        # Razorpay export uploaded before the Shopify order report, which
        # correctly can't be mapped yet (see
        # engine.razorpay_settlement/engine.transform_errors). Shown
        # directly, with none of the "does this look like some OTHER
        # report" checking below - that check was never relevant here, and
        # running it anyway is exactly what previously replaced this
        # file's correct, actionable error with a false "looks like a
        # Gokwik report" one, purely from a few shared generic column
        # names (Order ID, Amount, Settlement UTR).
        st.error(f"**{g_cfg['label']}**: {exc}")
        return None
    except KeyError as exc:
        # Before surfacing this source's own "couldn't find the right
        # data" error, check whether the file actually looks like some
        # OTHER known report entirely (e.g. a Delhivery file dropped into
        # the Shiprocket COD box) - that's a much more useful thing to
        # tell the user than "couldn't find a CRF-level sheet". Checked
        # against whichever sheet in the workbook is largest, since that's
        # the one most likely to carry the file's real identity.
        if registry:
            try:
                sheets_for_check = _read_all_sheets(f.getvalue(), f.name)
                biggest = max(sheets_for_check.values(), key=len, default=None)
            except Exception:
                biggest = None
            other_label, score, total = _best_other_match(biggest, registry, exclude=g_cfg["label"])
            if other_label:
                st.error(
                    f"**Wrong report detected.** You uploaded this file for **{g_cfg['label']}**, but it "
                    f"looks like a **{other_label}** ({score} of {total} expected columns matched). "
                    f"Please upload the correct **{g_cfg['label']}** report here instead."
                )
                return None
        st.error(f"Invalid **{g_cfg['label']}** file — {exc}")
        return None
    except Exception as exc:  # pragma: no cover - defensive, not expected in practice
        st.error(f"Could not read the **{g_cfg['label']}** file: {exc}")
        return None

    df.columns = [str(c).strip() for c in df.columns]

    # Same signature check every other source goes through (order_id_col /
    # amount_col here) - now run against the AUTO-MAPPED shape, so a file
    # that's genuinely missing something even after mapping (e.g. no
    # "Order Value" column at all) still gets a specific, named error
    # instead of silently proceeding with a broken/zeroed-out amount.
    is_valid, missing = validate_file_matches_source(df, g_cfg)
    if not is_valid:
        if registry:
            st.error(wrong_report_message(g_cfg["label"], df, registry, missing))
        else:
            st.error(
                friendly_mismatch_message(g_cfg)
                + " Missing expected column(s) even after auto-mapping the raw file: "
                + ", ".join(str(m) for m in missing) + "."
            )
        return None
    return df


def _best_other_match(df, registry, exclude):
    """Thin wrapper around engine.report_detection.detect_wrong_report -
    same distinctiveness-weighted, comparative-to-its-own-type confidence
    decision used by wrong_report_message, exposed separately here because
    _load_gateway_file's early sheet-detection failure path needs the same
    "does this look like some other report" answer but without an
    already-loaded single DataFrame + missing-columns list to build a full
    wrong_report_message from."""
    from engine.report_detection import detect_wrong_report

    if df is None:
        return None, 0, 0
    return detect_wrong_report(exclude, df, registry)


def _resolve_key_cols(df, col_specs):
    """Turns a list of config column-specs (each a str or list of aliases,
    or None for "this key piece isn't available") into the ACTUAL resolved
    column names present in df, via the same alias-tolerant resolve_col()
    used everywhere else - so the duplicate/updated-record key lines up
    with whatever header this particular export actually used, exactly
    like every other column lookup in this app."""
    return [resolve_col(df, spec) for spec in col_specs if spec]


def render_duplicate_check(df, key_col_specs, client_key, report_key, mode, label, widget_key):
    """
    Shared "have we already seen this before" preview + confirm step for
    the Upload Data page (see engine/dedup.py for the underlying logic;
    engine/storage.py for the ledger these compare against). Two modes:

    mode="skip" - SALES/order-level reports (DTC Orders, Amazon MTR).
    Rows matching a composite key already reconciled in a previous saved
    period are excluded (never double-counted); only genuinely new rows
    get returned once confirmed. Returns None until the user ticks the
    confirmation box - callers must treat None as "not ready yet, don't
    proceed" (mirrors the client's explicit "show a preview first" choice).

    mode="latest" - STATUS/SETTLEMENT reports (delivery partner, payment
    gateway, Amazon Settlement Flat Files). Nothing is ever excluded here
    - the same order/settlement showing a different status or amount than
    last time is exactly the point (Undelivered -> Delivered, a
    settlement correction, etc.), so the newest upload's data is always
    what gets used. This just previews which rows are brand new vs which
    are updating something already recorded, then returns the file
    unchanged once acknowledged.

    Returns the dataframe to actually use (df itself, or the new-only
    subset for mode="skip"), or None if the user hasn't confirmed yet /
    there's a data problem preventing the check from running.
    """
    if df is None or df.empty:
        return df

    key_cols = _resolve_key_cols(df, key_col_specs)
    if not key_cols:
        # Couldn't resolve even one key column (shouldn't normally happen -
        # order_id is always required to have gotten this far) - nothing
        # to check against, just pass the file through as-is.
        return df

    if mode == "skip":
        seen = storage.get_seen_keys(client_key, report_key)
        new_df, dup_df = dedup.split_new_vs_seen(df, key_cols, seen)
        if dup_df.empty:
            return df
        st.warning(
            f"**{label}**: {len(dup_df):,} of {len(df):,} row(s) match data already "
            f"reconciled in a previous upload for this channel - they'll be excluded "
            f"so nothing gets double-counted. **{len(new_df):,} new row(s)** will "
            "actually be processed."
        )
        with st.expander(f"Review the {len(dup_df):,} already-reconciled row(s) being excluded"):
            st.dataframe(dup_df, use_container_width=True, height=250)
        confirmed = st.checkbox(
            f"Looks right — use only the {len(new_df):,} new row(s) for {label}",
            key=widget_key,
        )
        return new_df if confirmed else None

    # mode == "latest"
    seen_records = storage.get_latest_records(client_key, report_key)
    new_df, updated_df = dedup.split_new_vs_updated(df, key_cols, set(seen_records.keys()))
    if updated_df.empty:
        return df
    st.info(
        f"**{label}**: {len(new_df):,} brand-new record(s), and **{len(updated_df):,} "
        "record(s) updating status/amount** from a previous upload - the newest data "
        "will be used for those, nothing is skipped."
    )
    with st.expander(f"Review the {len(updated_df):,} updated record(s)"):
        st.dataframe(updated_df, use_container_width=True, height=250)
    confirmed = st.checkbox(f"Looks right — use the updated data for {label}", key=widget_key)
    return df if confirmed else None


def reset_upload_state():
    """Clear previously uploaded/reconciled data so switching platforms
    mid-session can't accidentally mix one platform's numbers into another's
    report. Called both by this page's own platform/client selectors below
    AND by views/page_settings.py's channel switcher - one shared
    implementation so the two can never quietly drift apart the way they
    once did (see storage.RAW_UPLOAD_SLOTS's docstring: Settings used to
    have its own hand-rolled, incomplete reset that never restored raw
    uploads from disk at all).

    Immediately followed by restoring whatever raw files are already
    persisted on disk (see _restore_raw_uploads below) for the NOW-active
    client_key - the clear above is about isolating one client/channel's
    data from another's, not about forgetting a client's own already-
    confirmed uploads every time this runs. Order matters: this must
    happen AFTER the clear above (and after state_init.set_active_config
    has already updated client_key at every call site of this function),
    or the restore would just get immediately wiped back out."""
    for key in _UPLOAD_STATE_KEYS:
        st.session_state[key] = {} if key in storage.RAW_UPLOAD_DICT_KEYS else None
    _restore_raw_uploads()


def _restore_raw_uploads():
    """
    Re-populates every raw upload slot (see engine.storage.RAW_UPLOAD_SLOTS
    - orders_df, delivery_frames, gateway_frames, attribution_frames,
    bank_df, mtr_files, settlement_files) for the CURRENTLY active
    client_key straight from engine.storage's on-disk raw-uploads store
    (see storage.load_raw_uploads's own docstring for the client-reported
    symptom this fixes: after any app restart or session expiry, a raw
    file that had genuinely already been uploaded and confirmed - possibly
    in a completely different process/session - used to look exactly like
    it had never been uploaded at all, both to Data Management's per-file
    delete list and to any raw_transform (Razorpay) that depends on it
    being available).
    """
    client_key = st.session_state.get("client_key")
    if not client_key:
        return
    restored = storage.load_raw_uploads(client_key)
    for key in storage.RAW_UPLOAD_SESSION_KEYS:
        if key in restored:
            st.session_state[key] = restored[key]


def _select_platform():
    """Step 1: ask which platform this data is for. Returns the matching
    config label once one's been resolved, or None if the page should stop
    here (nothing chosen yet, or no config exists for that platform)."""
    st.subheader("1. Which platform is this data for?")
    placeholder = "— Select a platform —"
    options = [placeholder] + PLATFORMS
    current = st.session_state.get("upload_platform", placeholder)
    index = options.index(current) if current in options else 0
    platform = st.selectbox(
        "Platform", options, index=index, key="upload_platform_select", label_visibility="collapsed"
    )

    if platform != st.session_state.get("upload_platform"):
        st.session_state["upload_platform"] = platform
        reset_upload_state()

    if platform == placeholder:
        st.info("Pick a platform above to continue to file uploads.")
        return None

    config_labels = st.session_state.get("config_labels", {})
    matches = [
        label for label, cfg in config_labels.items()
        if cfg.get("channel_name", "").strip().lower() == platform.strip().lower()
    ]

    if not matches:
        st.warning(
            f"No config is set up for **{platform}** yet. Ask your admin to add a "
            f"configs/*.json file with {platform}'s column mappings (see the "
            f"Settings page for how) - it'll show up here automatically once it exists."
        )
        return None

    if len(matches) == 1:
        chosen_label = matches[0]
    else:
        chosen_label = st.selectbox("Which client account?", matches, key="upload_client_select")

    if st.session_state.get("chosen_label") != chosen_label:
        set_active_config(chosen_label)
        reset_upload_state()

    return chosen_label


def _marketplace_signature(col_spec_list):
    """Builds an expected_signature (see load_table_smart) from a list of
    column specs pulled straight out of the marketplace config's own
    mtr_columns/settlement_columns block - so, like the DTC signature in
    validation.py, there's no separate list to keep in sync by hand."""
    return [spec for spec in col_spec_list if spec]


def load_and_validate_marketplace(uploaded_file, signature, label, registry=None):
    """Marketplace-shaped equivalent of load_and_validate() above - the
    DTC version keys off a "source_cfg" dict shaped like
    {order_id_col, status_col, ...}, which MTR/Flat File configs aren't;
    this instead takes the signature (already resolved from
    mtr_columns/settlement_columns) and the human-readable label directly."""
    if uploaded_file is None:
        return None
    df, header_row_used = _load_bytes_smart(uploaded_file.getvalue(), uploaded_file.name, signature)
    if header_row_used is None:
        missing = [s[0] if isinstance(s, list) else s for s in signature if resolve_col(df, s) is None]
        if registry:
            st.error(wrong_report_message(label, df, registry, missing))
        else:
            st.error(
                f"Invalid file. Please upload the **{label}** report here — "
                f"the file you uploaded doesn't have the columns a {label} export should have. "
                f"Missing expected column(s): {', '.join(str(m) for m in missing)}."
            )
        return None
    if header_row_used > 0:
        st.info(
            f"Heads up: {label} had {header_row_used} extra row(s) above the actual "
            "column headers - skipped past them automatically."
        )
    return df


def _render_dtc_uploads(config, registry):
    """The original Shopify/DTC-shaped upload flow: one order export, N
    delivery-partner files, N payment-gateway files, one bank statement.

    Every slot here also runs through render_duplicate_check() against this
    client's ingestion ledger (see engine/storage.py) - Orders uses "skip"
    mode (Order ID alone identifies a permanent sales record - re-uploading
    a wider date range should only add what's genuinely new), Delivery
    partner and Gateway files use "latest" mode (the same order's status/
    amount can legitimately change in a later file - Order ID + AWB for
    delivery, Order ID + Transaction Type for gateways, per the client's
    own confirmed key choices). Nothing is added to session_state until the
    user has confirmed the preview whenever one was shown - see
    render_duplicate_check's docstring for exactly when a preview appears.

    registry: the app-wide report signature registry (see
    engine.report_detection.build_report_registry, built once in
    render() below) - passed through to every load_and_validate()/
    _load_gateway_file() call so a wrong file dropped in any box here can
    be identified by name, not just rejected as "missing columns"."""
    client_key = st.session_state.get("client_key")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Orders")
        orders_upload = st.file_uploader(
            f"{config['orders']['label']} (order export)", type=["xlsx", "xls", "csv"], key="orders_uploader"
        )
        raw_orders_df = load_and_validate(orders_upload, config["orders"], registry=registry)
        # Client-reported (2026-08-21): session_state["orders_df"] (and
        # every other upload slot below) used to be OVERWRITTEN on every
        # render with whatever this uploader currently returns - including
        # None the moment the widget doesn't currently hold an active file
        # (e.g. after navigating away to a different page and back - a
        # normal Streamlit behaviour, not a bug in the widget). That
        # silently wiped out data the user had already uploaded and
        # confirmed earlier, just because they revisited this page to add
        # ONE more report. Fixed by only ever touching session_state when
        # THIS render actually produced fresh, confirmed data - and even
        # then, merging it into whatever's already accumulated (see
        # engine.dedup.accumulate_df) rather than replacing it outright, so
        # a later, separate upload of e.g. just one more month's orders
        # ADDS to the working dataset instead of replacing it.
        if raw_orders_df is not None:
            confirmed_new_orders = render_duplicate_check(
                raw_orders_df, [config["orders"]["order_id_col"]], client_key, "orders", "skip",
                config["orders"]["label"], f"dup_confirm_orders_{client_key}",
            )
            if confirmed_new_orders is not None:
                key_cols = _resolve_key_cols(confirmed_new_orders, [config["orders"]["order_id_col"]])
                st.session_state["orders_df"] = dedup.accumulate_df(
                    st.session_state.get("orders_df"), confirmed_new_orders, key_cols,
                )
                # Persisted to disk too (see engine.storage.save_raw_upload's
                # own docstring) - so this confirmed upload survives an app
                # restart or session expiry instead of only living in
                # st.session_state, which the 2026-08-22 client report
                # showed gets silently forgotten by both Data Management's
                # per-file delete list and any raw_transform (Razorpay)
                # that depends on this exact data being available.
                storage.save_raw_upload(client_key, "orders_df", None, st.session_state["orders_df"])

        st.subheader("Delivery partners")
        delivery_frames = dict(st.session_state.get("delivery_frames") or {})
        for d_cfg in config["delivery_partners"]:
            f = st.file_uploader(d_cfg["label"], type=["xlsx", "xls", "csv"], key=f"delivery_{d_cfg['label']}")
            raw_df = load_and_validate(f, d_cfg, registry=registry)
            if raw_df is not None:
                checked_df = render_duplicate_check(
                    raw_df, [d_cfg["order_id_col"], _AWB_COLUMN_ALIASES], client_key,
                    f"delivery__{d_cfg['label']}", "latest", d_cfg["label"],
                    f"dup_confirm_delivery_{d_cfg['label']}_{client_key}",
                )
                if checked_df is not None:
                    key_cols = _resolve_key_cols(checked_df, [d_cfg["order_id_col"], _AWB_COLUMN_ALIASES])
                    delivery_frames[d_cfg["label"]] = dedup.accumulate_df(
                        delivery_frames.get(d_cfg["label"]), checked_df, key_cols,
                    )
                    storage.save_raw_upload(client_key, "delivery_frames", d_cfg["label"], delivery_frames[d_cfg["label"]])
        st.session_state["delivery_frames"] = delivery_frames

    with col2:
        st.subheader("Payment / COD gateways")
        gateway_frames = dict(st.session_state.get("gateway_frames") or {})
        for g_cfg in config["gateways"]:
            f = st.file_uploader(g_cfg["label"], type=["xlsx", "xls", "csv"], key=f"gateway_{g_cfg['label']}")
            raw_df = _load_gateway_file(f, g_cfg, registry=registry)
            if raw_df is not None:
                # Client-reported (2026-08-23): a settlement/gateway file
                # whose date range overlaps a previously-uploaded one (e.g.
                # FY2025-26's Razorpay export ran through May, FY2026-27's
                # fresh export starts back in April) caused real settlement
                # money to go missing on re-run, but ONLY when the two
                # uploads overlapped - a single, standalone upload was
                # fine. Root cause: order_id + type_col alone identifies a
                # real order+transaction-TYPE combination, not one
                # specific settlement EVENT - an order can legitimately
                # have more than one row of the same type (a split/partial
                # settlement, a separate adjustment entry, etc). Because
                # dedup.accumulate_df REPLACES every existing row sharing a
                # key with whatever the new file has for that same key,
                # two overlapping exports that don't happen to carry the
                # exact same NUMBER of rows for one order+type key (a
                # normal outcome of two different point-in-time settlement
                # pulls) silently dropped whichever of the old rows didn't
                # have a same-key counterpart in the new file - real,
                # already-reconciled money, gone with no error or warning.
                # Adding the row's own amount into the key (see
                # engine.dedup.build_row_key's numeric-rounding fix, added
                # alongside this) makes two rows collide only when they
                # plausibly ARE the same settlement event (same order, same
                # type, same amount) - a later file correcting only the
                # STATUS of such a row (COD-pending -> COD-realised, say)
                # still matches and updates in place, exactly as before.
                # The tradeoff: if a later export revises the AMOUNT of a
                # genuinely-the-same event too, this key now treats it as a
                # separate row instead of an update - a visible duplicate a
                # human can catch via the existing reconciliation checks,
                # not a silent, invisible loss like the bug this replaces.
                checked_df = render_duplicate_check(
                    raw_df, [g_cfg["order_id_col"], g_cfg.get("type_col"), g_cfg.get("amount_col")], client_key,
                    f"gateway__{g_cfg['label']}", "latest", g_cfg["label"],
                    f"dup_confirm_gateway_{g_cfg['label']}_{client_key}",
                )
                if checked_df is not None:
                    key_cols = _resolve_key_cols(
                        checked_df, [g_cfg["order_id_col"], g_cfg.get("type_col"), g_cfg.get("amount_col")]
                    )
                    gateway_frames[g_cfg["label"]] = dedup.accumulate_df(
                        gateway_frames.get(g_cfg["label"]), checked_df, key_cols,
                    )
                    storage.save_raw_upload(client_key, "gateway_frames", g_cfg["label"], gateway_frames[g_cfg["label"]])
        st.session_state["gateway_frames"] = gateway_frames

        attribution_sources = config.get("attribution_sources", [])
        if attribution_sources:
            st.subheader("Payment Gateway Attribution (optional)")
            st.caption(
                "Optional - only needed to show the specific downstream payment "
                "processor (e.g. easebuzz/payu) instead of the bare checkout-"
                "aggregator name, for a gateway like Gokwik that routes through "
                "one, in the \"Payment Gateway\" column. Upload BOTH reports for "
                "a provider to get that refinement; either one alone, or "
                "neither, is fine - nothing else in the tool depends on these, "
                "and no settlement figures change either way."
            )
            attribution_frames = dict(st.session_state.get("attribution_frames") or {})
            for a_cfg in attribution_sources:
                f = st.file_uploader(a_cfg["label"], type=["xlsx", "xls", "csv"], key=f"attribution_{a_cfg['label']}")
                raw_df = load_and_validate(f, a_cfg, registry=registry)
                if raw_df is not None:
                    key_specs = [spec for spec in (a_cfg.get("order_id_col"), a_cfg.get("payment_id_col")) if spec]
                    checked_df = render_duplicate_check(
                        raw_df, key_specs, client_key,
                        f"attribution__{a_cfg['label']}", "latest", a_cfg["label"],
                        f"dup_confirm_attribution_{a_cfg['label']}_{client_key}",
                    )
                    if checked_df is not None:
                        key_cols = _resolve_key_cols(checked_df, key_specs)
                        attribution_frames[a_cfg["label"]] = dedup.accumulate_df(
                            attribution_frames.get(a_cfg["label"]), checked_df, key_cols,
                        )
                        storage.save_raw_upload(
                            client_key, "attribution_frames", a_cfg["label"], attribution_frames[a_cfg["label"]]
                        )
            st.session_state["attribution_frames"] = attribution_frames

        _render_bank_upload(config, registry)

    st.divider()
    ready = st.session_state.get("orders_df") is not None
    if ready:
        st.success("Orders file uploaded. Head to **Reconciliation** to run it.")
    else:
        st.info("Upload at least the orders file to continue to the Reconciliation page.")


def _render_bank_upload(config, registry):
    """Shared by both upload flows - same engine.bank.load_bank_statement
    consumes whatever comes out of here regardless of channel type.

    Same accumulate-not-overwrite fix as every other slot (see
    _render_dtc_uploads' orders section for the full story) - a bank
    statement export has no natural single ID column, so key_cols=None
    below makes dedup.accumulate_df fall back to a full-row-content hash:
    re-uploading the exact same statement is a safe no-op, a later
    month's statement with no overlapping rows is pure addition, and a
    revisit that doesn't touch this uploader no longer wipes out a bank
    statement uploaded earlier in the session."""
    if "bank_statement" not in config:
        return
    st.subheader("Bank statement (optional)")
    st.caption(
        "Links settlements/receipts to actual bank credits. Accepts either "
        "a plain Excel export (separate Withdrawal/Deposit columns) or a "
        "CSV export with a combined Amount + Dr/Cr column, letterhead rows "
        "and all - the tool detects the real header row and the amount "
        "format automatically."
    )
    bank_cfg = config["bank_statement"]
    f = st.file_uploader(bank_cfg["label"], type=["xlsx", "xls", "csv"], key="bank_statement_uploader")
    raw_df = load_and_validate(f, bank_cfg, registry=registry)
    if raw_df is not None:
        st.session_state["bank_df"] = dedup.accumulate_df(st.session_state.get("bank_df"), raw_df, None)
        storage.save_raw_upload(st.session_state.get("client_key"), "bank_df", None, st.session_state["bank_df"])


def _render_marketplace_uploads(config, registry):
    """
    Amazon (and future Flipkart/Tata 1mg/First Cry) upload flow: MTR-style
    revenue reports (one per segment, e.g. B2C/B2B) and Settlement Flat
    File-style expense reports (one per payment mode, e.g. COD/Online),
    instead of the DTC flow's orders/delivery-partners/gateways shape.

    Every uploader here accepts BOTH .csv and .xlsx, and every column is
    identified by header name (via engine.loaders.resolve_col through the
    config's mtr_columns/settlement_columns block) - not by position - so
    deleting an unused column or re-ordering columns in Amazon's own export
    doesn't break loading. Multiple date-range chunks pasted into one
    "Conso" file (with the header row repeated partway down) are also
    handled - see engine.amazon_loaders.drop_embedded_header_rows.

    Per the client's "Important Warning - Amazon MTR Report" note, MTR
    itself only needs "skip" mode at the FILE-upload-dedup level (Order ID +
    SKU + Transaction Type - a different transaction_type, e.g. Shipment vs
    Refund on the same order, is a genuinely different row and is never
    excluded here as a "duplicate"; which status is the order's CURRENT one
    is a separate concern handled downstream in
    engine.amazon_reco.build_mtr_order_summary's latest-by-date logic, not
    here). Settlement Flat Files use "latest" mode (Order ID + Transaction
    Type) - a settlement/adjustment can legitimately update in a later
    file.

    registry: same app-wide report signature registry threaded through
    _render_dtc_uploads above - see that docstring."""
    client_key = st.session_state.get("client_key")
    mtr_cols_cfg = config.get("mtr_columns", {})
    settlement_cols_cfg = config.get("settlement_columns", {})
    mtr_signature = _marketplace_signature([
        mtr_cols_cfg.get("order_id_col"), mtr_cols_cfg.get("sku_col"), mtr_cols_cfg.get("transaction_type_col"),
    ])
    settlement_signature = _marketplace_signature([
        settlement_cols_cfg.get("settlement_id_col"), settlement_cols_cfg.get("transaction_type_col"),
    ])
    mtr_key_specs = [
        mtr_cols_cfg.get("order_id_col"), mtr_cols_cfg.get("sku_col"), mtr_cols_cfg.get("transaction_type_col"),
    ]
    settlement_key_specs = [
        settlement_cols_cfg.get("order_id_col"), settlement_cols_cfg.get("transaction_type_col"),
    ]

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("MTR Reports (revenue)")
        st.caption("Amazon's Merchant Tax Report - one file per segment.")
        # Accumulate-not-overwrite, same as delivery_frames/gateway_frames in
        # _render_dtc_uploads above: start from whatever's already in
        # session_state (not a fresh {}), and only touch a given segment's
        # entry when this render actually produced a confirmed upload for
        # it - so uploading just one more segment, or a later month's file
        # for a segment already uploaded, never drops what was already
        # accumulated for every OTHER segment (this was silently wiping
        # earlier-uploaded segments before this fix).
        mtr_files = dict(st.session_state.get("mtr_files") or {})
        for r_cfg in config.get("mtr_reports", []):
            f = st.file_uploader(r_cfg["label"], type=["xlsx", "xls", "csv"], key=f"mtr_{r_cfg['segment']}")
            raw_df = load_and_validate_marketplace(f, mtr_signature, r_cfg["label"], registry=registry)
            if raw_df is not None:
                checked_df = render_duplicate_check(
                    raw_df, mtr_key_specs, client_key, f"mtr__{r_cfg['segment']}", "skip",
                    r_cfg["label"], f"dup_confirm_mtr_{r_cfg['segment']}_{client_key}",
                )
                if checked_df is not None:
                    key_cols = _resolve_key_cols(checked_df, mtr_key_specs)
                    mtr_files[r_cfg["segment"]] = dedup.accumulate_df(
                        mtr_files.get(r_cfg["segment"]), checked_df, key_cols,
                    )
                    storage.save_raw_upload(client_key, "mtr_files", r_cfg["segment"], mtr_files[r_cfg["segment"]])
        st.session_state["mtr_files"] = mtr_files

    with col2:
        st.subheader("Settlement Flat Files (expenses)")
        st.caption("Amazon's Date Range Transaction Report - one file per payment mode.")
        # Same accumulate-not-overwrite treatment as mtr_files above.
        settlement_files = dict(st.session_state.get("settlement_files") or {})
        for s_cfg in config.get("settlement_files", []):
            f = st.file_uploader(s_cfg["label"], type=["xlsx", "xls", "csv"], key=f"settlement_{s_cfg['payment_mode']}")
            raw_df = load_and_validate_marketplace(f, settlement_signature, s_cfg["label"], registry=registry)
            if raw_df is not None:
                checked_df = render_duplicate_check(
                    raw_df, settlement_key_specs, client_key, f"settlement__{s_cfg['payment_mode']}", "latest",
                    s_cfg["label"], f"dup_confirm_settlement_{s_cfg['payment_mode']}_{client_key}",
                )
                if checked_df is not None:
                    key_cols = _resolve_key_cols(checked_df, settlement_key_specs)
                    settlement_files[s_cfg["payment_mode"]] = dedup.accumulate_df(
                        settlement_files.get(s_cfg["payment_mode"]), checked_df, key_cols,
                    )
                    storage.save_raw_upload(client_key, "settlement_files", s_cfg["payment_mode"], settlement_files[s_cfg["payment_mode"]])
        st.session_state["settlement_files"] = settlement_files

    st.divider()
    _render_bank_upload(config, registry)

    st.divider()
    ready = bool(st.session_state.get("mtr_files")) and bool(st.session_state.get("settlement_files"))
    if ready:
        st.success("MTR and Settlement Flat File(s) uploaded. Head to **Reconciliation** to run it.")
    else:
        st.info("Upload at least one MTR report and one Settlement Flat File to continue to the Reconciliation page.")


def render():
    st.title("Upload Data")

    chosen_label = _select_platform()
    if not chosen_label:
        return

    config = st.session_state.get("config")
    if not config:
        st.error("No client/channel config selected. Go to Settings first.")
        return

    st.caption(f"{config['client_name']} — {config['channel_name']}")
    st.divider()
    st.subheader("2. Upload files")

    # Built fresh from every loaded config (not just the active one) so
    # "wrong report uploaded here" can recognise ANY of the app's report
    # types - e.g. an Amazon MTR file dropped into a Shopify orders box.
    registry = build_report_registry(st.session_state.get("config_labels", {}))

    if config.get("channel_type") == "marketplace":
        _render_marketplace_uploads(config, registry)
    else:
        _render_dtc_uploads(config, registry)
