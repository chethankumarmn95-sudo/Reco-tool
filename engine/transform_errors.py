"""
transform_errors.py
--------------------
Just the one shared exception type raw_transforms (engine.raw_transforms)
can raise to mean "I can't proceed, but not because the uploaded file
itself is the wrong report" - kept in its own tiny module (rather than
inside engine.raw_transforms itself) purely to avoid a circular import:
individual transform modules like engine.razorpay_settlement need to
raise this, but engine.raw_transforms already imports those transform
modules to build its registry, so a transform module importing back from
engine.raw_transforms would be circular.
"""


class TransformPrerequisiteError(Exception):
    """
    Raised by a raw_transform when it can't proceed for a reason that has
    NOTHING to do with whether the uploaded file itself is the right
    report - typically, some OTHER already-uploaded source it depends on
    (e.g. Razorpay's mapping needing the Shopify order report uploaded
    first - see engine.razorpay_settlement) isn't available yet, or is
    itself missing something the transform needs.

    Client-reported (2026-08-21): views/page_upload.py's _load_gateway_file
    used to catch every failure from a raw_transform as a plain KeyError,
    which made it run the "does this file look like some OTHER known
    report" check even for a failure that had nothing to do with the
    file's own identity - e.g. a genuine Razorpay settlement export
    uploaded before the Shopify order report, which correctly can't be
    mapped yet, got its clear, correct, actionable error ("the Shopify
    order report needs to be uploaded first") silently replaced by a
    confusing false "looks like a Gokwik report" message, purely because
    a handful of generic column names (Order ID, Amount, Settlement UTR)
    happen to be shared between the two gateways' raw exports.

    A raw_transform should raise THIS (not KeyError) for that class of
    failure, so _load_gateway_file can show its message directly instead
    of running a wrong-report check that was never relevant in the first
    place. Reserve plain KeyError for "this file's own shape doesn't look
    right" failures (wrong/missing sheet, wrong/missing column within it) -
    those are exactly the cases the wrong-report check is meant to help
    explain.
    """
