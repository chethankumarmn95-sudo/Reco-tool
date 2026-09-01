"""
amazon_loaders.py
------------------
Turns Amazon's raw exports - MTR (Merchant Tax Report, B2C/B2B) and the
Settlement "Flat File" (Date Range Transaction Report, COD/Online) - into
clean pandas DataFrames, and combines multiple files of the same report
type (e.g. MTR B2C + MTR B2B, or Flat File COD + Flat File Online) into one
frame per report type.

Everything here is header-name driven (via engine.loaders.resolve_col), not
position-driven - deleting an unused column or moving columns around in the
source export does not break loading, per the client's explicit requirement.
Both .csv and .xlsx are accepted for every report (engine.loaders.load_table
already handles that).

Why this file exists separately from engine/loaders.py: Amazon's exports
have a quirk .csv/.xlsx exports from other systems don't - when a seller
downloads several date-range chunks and pastes them into one "Conso" file
(exactly what the client did for the Flat File COD/Online consolidated
files), the header row from each chunk often gets pasted in again partway
down the file. If left in, that junk row would be read as a real
transaction with an order-id of "order-id" etc. drop_embedded_header_rows()
below strips these out before anything else touches the data.
"""

import re

import pandas as pd

from .loaders import load_table, resolve_col, resolve_col_or_raise, normalize_order_id


def drop_embedded_header_rows(df, anchor_col):
    """
    Removes rows that are actually a repeated header row pasted into the
    body of the file - identified by the anchor column's own value being
    (case/space-insensitive) equal to the anchor column's own header text.
    Safe/cheap: only ever removes rows that are 100% junk (a real Amazon
    order-id or settlement-id can never literally equal the column name).
    """
    if anchor_col not in df.columns:
        return df
    anchor_text = str(anchor_col).strip().lower()
    is_junk = df[anchor_col].astype(str).str.strip().str.lower() == anchor_text
    return df[~is_junk].reset_index(drop=True)


def _load_one(file):
    """Accepts either an already-loaded DataFrame or a raw file/path and
    always returns a DataFrame with stripped column names.

    Why both: views/page_upload.py already runs each upload through
    load_and_validate_marketplace() (header-shift detection + a friendly
    error message right there on the Upload Data page) and stores the
    resulting DataFrame in session_state["mtr_files"]/["settlement_files"]
    - exactly the same pattern engine/bank.py's load_bank_statement()
    already uses for bank_df. Re-running load_table() on that DataFrame
    would try to hand it to pandas' file readers (pd.read_excel/read_csv),
    which raises "Invalid file path or buffer object type: DataFrame".
    Passing a raw file/path through still works too, so this module can
    also be used standalone/from a script without going through the
    Streamlit upload page."""
    if isinstance(file, pd.DataFrame):
        return file.copy()
    return load_table(file)


def load_mtr_reports(mtr_files, mtr_cols_cfg):
    """
    mtr_files: dict {segment_label ("B2C"/"B2B") -> uploaded file or path}
               Only include segments the user actually uploaded.
    mtr_cols_cfg: the "mtr_columns" block from the client config.

    Returns one combined, normalized MTR ledger - one row per shipment
    item/refund line/cancel line across every uploaded MTR segment - with a
    "Segment" column (B2C/B2B) so downstream reporting can still split by
    segment (matches the client's own "Sales as per MTR - Segment
    Breakdown" table).
    """
    order_col_spec = mtr_cols_cfg["order_id_col"]
    sku_col_spec = mtr_cols_cfg["sku_col"]
    tt_col_spec = mtr_cols_cfg["transaction_type_col"]

    frames = []
    for segment, file in mtr_files.items():
        if file is None:
            continue
        raw = _load_one(file)
        order_col = resolve_col_or_raise(raw, order_col_spec, f"MTR {segment}")
        raw = drop_embedded_header_rows(raw, order_col)
        if raw.empty:
            continue

        out = pd.DataFrame(index=raw.index)
        out["order_id"] = normalize_order_id(raw[order_col])
        out["sku"] = raw[resolve_col_or_raise(raw, sku_col_spec, f"MTR {segment}")].astype(str).str.strip()
        out["segment"] = segment

        tt_col = resolve_col_or_raise(raw, tt_col_spec, f"MTR {segment}")
        out["mtr_transaction_type"] = raw[tt_col].astype(str).str.strip()

        def _num(col_key, default=0.0):
            col_spec = mtr_cols_cfg.get(col_key)
            if not col_spec:
                return pd.Series(default, index=raw.index)
            resolved = resolve_col(raw, col_spec)
            if resolved is None:
                return pd.Series(default, index=raw.index)
            return pd.to_numeric(raw[resolved], errors="coerce").fillna(default)

        def _date(col_key):
            col_spec = mtr_cols_cfg.get(col_key)
            resolved = resolve_col(raw, col_spec) if col_spec else None
            if resolved is None:
                return pd.NaT
            s = pd.to_datetime(raw[resolved], errors="coerce")
            try:
                if s.dt.tz is not None:
                    s = s.dt.tz_localize(None)
            except (AttributeError, TypeError):
                pass
            return s

        def _text(col_key):
            col_spec = mtr_cols_cfg.get(col_key)
            resolved = resolve_col(raw, col_spec) if col_spec else None
            if resolved is None:
                return None
            return raw[resolved]

        out["invoice_number"] = _text("invoice_number_col")
        out["invoice_date"] = _date("invoice_date_col")
        out["order_date"] = _date("order_date_col")
        out["shipment_date"] = _date("shipment_date_col")
        out["quantity"] = _num("quantity_col")
        out["invoice_amount"] = _num("invoice_amount_col")
        out["taxable"] = _num("taxable_col")
        out["tax_amount"] = _num("tax_amount_col")
        out["principal_amount"] = _num("principal_amount_col")
        out["shipping_amount"] = _num("shipping_amount_col")
        out["item_promo_discount"] = _num("item_promo_discount_col")
        out["shipping_promo_discount"] = _num("shipping_promo_discount_col")

        tcs_cols = mtr_cols_cfg.get("tcs_cols", [])
        tcs_total = pd.Series(0.0, index=raw.index)
        for spec in tcs_cols:
            resolved = resolve_col(raw, spec)
            if resolved is not None:
                tcs_total = tcs_total.add(pd.to_numeric(raw[resolved], errors="coerce").fillna(0.0), fill_value=0.0)
        out["tcs_amount"] = tcs_total

        pay_col = resolve_col(raw, mtr_cols_cfg.get("payment_method_col")) if mtr_cols_cfg.get("payment_method_col") else None
        cod_values = {v.strip().upper() for v in mtr_cols_cfg.get("cod_values", ["COD"])}
        if pay_col is not None:
            # fillna("") BEFORE astype(str)/str.upper(): with pandas' newer
            # string dtype, a NaN cell survives astype(str)/.str.upper() as
            # an actual float NaN rather than becoming the text "nan" - and
            # the vectorized str.contains() check below would otherwise
            # crash/mis-flag on a float. fillna("") first sidesteps that
            # entirely (an order with no payment method text just falls
            # through to "Prepaid" below, which is the same "Unknown
            # treated as Prepaid" fallback engine/bank.py already uses).
            #
            # Vectorized via str.contains(regex OR of every cod_value)
            # instead of a per-row .apply(lambda ...) - an MTR export can
            # run into the lakhs of rows for a full financial year, and a
            # Python-level lambda call per row adds up; a single regex
            # pass over the whole column is a C-level operation regardless
            # of row count. cod_values is typically just {"COD"}, so this
            # is functionally identical to the old `any(cv in t for cv in
            # cod_values)` check, just computed for every row at once.
            pay_text = raw[pay_col].fillna("").astype(str).str.upper()
            cod_pattern = "|".join(re.escape(cv) for cv in cod_values) if cod_values else r"(?!)"
            out["payment_type"] = pay_text.str.contains(cod_pattern, regex=True, na=False).map(
                {True: "COD", False: "Prepaid"}
            )
            out["payment_method_raw"] = raw[pay_col]
        else:
            out["payment_type"] = "Unknown"
            out["payment_method_raw"] = None

        out["fulfillment_channel"] = _text("fulfillment_channel_col")
        out["warehouse_id"] = _text("warehouse_id_col")
        out["credit_note_no"] = _text("credit_note_no_col")
        out["credit_note_date"] = _date("credit_note_date_col")
        out["asin"] = _text("asin_col")
        out["item_description"] = _text("item_description_col")

        frames.append(out)

    if not frames:
        return pd.DataFrame(columns=[
            "order_id", "sku", "segment", "mtr_transaction_type", "invoice_number", "invoice_date",
            "order_date", "shipment_date", "quantity", "invoice_amount", "taxable", "tax_amount",
            "principal_amount", "shipping_amount", "item_promo_discount", "shipping_promo_discount",
            "tcs_amount", "payment_type", "payment_method_raw", "fulfillment_channel", "warehouse_id",
            "credit_note_no", "credit_note_date", "asin", "item_description",
        ])

    return pd.concat(frames, ignore_index=True)


def load_settlement_files(settlement_files, settlement_cols_cfg):
    """
    settlement_files: dict {payment_mode_label ("COD"/"Online") -> uploaded
                       file or path}. Only include the ones actually
                       uploaded.
    settlement_cols_cfg: the "settlement_columns" block from the client
                       config.

    Returns one combined RAW settlement frame (embedded header rows
    stripped, a "payment_mode" column added) - NOT yet melted into the
    granular expense ledger. See engine/amazon_consolidator.py for the melt
    step; kept separate so this loader stays a pure "get me a clean
    DataFrame" step, consistent with the rest of this engine's layering.
    """
    settlement_id_col_spec = settlement_cols_cfg["settlement_id_col"]

    frames = []
    for payment_mode, file in settlement_files.items():
        if file is None:
            continue
        raw = _load_one(file)
        sid_col = resolve_col_or_raise(raw, settlement_id_col_spec, f"Settlement Flat File - {payment_mode}")
        raw = drop_embedded_header_rows(raw, sid_col)
        if raw.empty:
            continue
        raw = raw.copy()
        raw["_payment_mode"] = payment_mode
        raw["_settlement_id_resolved"] = raw[sid_col]
        frames.append(raw)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)
