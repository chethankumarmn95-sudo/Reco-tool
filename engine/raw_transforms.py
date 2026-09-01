"""
raw_transforms.py
------------------
The single combined registry of every "raw file needs auto-mapping before
the rest of the engine can use it" transform in the app - consulted by
views/page_upload.py's _load_gateway_file whenever a gateway config sets
"raw_transform": "<key>".

Kept as one small dict, gathered from each source's own module, rather
than a big if/elif chain in page_upload.py or one giant transforms file -
adding a future source's own multi-sheet/cross-file mapping logic is a
one-line registration here (plus writing that source's own module,
following either engine.shiprocket_cod's or engine.razorpay_settlement's
pattern), never a new branch in the upload page itself.

Every registered function shares the same calling convention:

    transform_fn(sheets, context=None) -> DataFrame

sheets: {sheet_name: DataFrame} - every sheet in the uploaded workbook
(see views/page_upload.py's _read_all_sheets).

context: an optional dict of whatever OTHER already-uploaded data a
transform might need beyond the file's own sheets - today just
{"orders_df": ...} (the Shopify order report, needed by Razorpay's
mapping to translate its internal payment token into a real order
number - see engine.razorpay_settlement). A transform that doesn't need
any of it (e.g. Shiprocket's, which only ever combines two sheets from
within its own file) simply accepts and ignores context. This shared
shape is what lets page_upload.py call every transform in RAW_TRANSFORMS
identically, with no per-source special-casing.
"""

from .shiprocket_cod import map_shiprocket_cod_raw_to_mapped
from .razorpay_settlement import map_razorpay_settlement_to_shopify
# Re-exported here so callers (views/page_upload.py) can import it from
# the same place as the transforms themselves - see engine.transform_errors
# for why the class itself lives in its own tiny module (avoiding a
# circular import between this registry and the transform modules it
# imports, which need to raise this same exception type).
from .transform_errors import TransformPrerequisiteError  # noqa: F401

RAW_TRANSFORMS = {
    "shiprocket_cod": map_shiprocket_cod_raw_to_mapped,
    "razorpay_settlement": map_razorpay_settlement_to_shopify,
}
