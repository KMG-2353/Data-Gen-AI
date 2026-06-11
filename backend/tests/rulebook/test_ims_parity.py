"""Parity tests for the IMS extraction + delegation seam (U7)."""
import copy
import random

from app.rulebook import config as rb_config
from app.policies.ims import _enforce_property_sheet_values


def _fixture_rows():
    return [
        {
            "Number of Stories": "2",
            "Coinsurance": "70",          # invalid -> snapped
            "Cause of Loss": "Comprehensive",  # invalid -> snapped
        },
        {
            "Coinsurance": "80",          # valid -> unchanged
            "Cause of Loss": "Special",   # valid -> unchanged
        },
    ]


def _run(flag, rows):
    original = rb_config.RULEBOOK_ENABLED
    rb_config.RULEBOOK_ENABLED = flag
    try:
        random.seed(99)
        return _enforce_property_sheet_values(copy.deepcopy(rows))
    finally:
        rb_config.RULEBOOK_ENABLED = original


def test_engine_path_matches_handler_path():
    rows = _fixture_rows()
    assert _run(True, rows) == _run(False, rows)


def test_valid_unchanged_invalid_snapped():
    out = _run(True, _fixture_rows())
    assert out[0]["Coinsurance"] in ("80", "90", "100")
    assert out[0]["Cause of Loss"] in ("Basic", "Special", "Broad")
    assert out[1]["Coinsurance"] == "80"      # valid, untouched
    assert out[1]["Cause of Loss"] == "Special"
