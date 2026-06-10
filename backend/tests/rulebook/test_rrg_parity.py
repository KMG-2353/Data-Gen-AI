"""Parity tests for the RRG extraction + delegation seam (U4).

Characterization-first: the RULEBOOK_ENABLED=False path IS the current handler
behavior. We assert the engine path (flag on) reproduces it byte-for-byte under
the same seeded RNG, on fixture rows mixing valid and invalid values.
"""
import copy
import random

import pytest

from app.rulebook import config as rb_config
from app.policies.rrg import RrgHandler


def _fixture_rows():
    return [
        {
            "Test ID": "X",
            "Contact Number": "(212) 555-0147",
            "State": "ZZ",                 # invalid -> snapped
            "State of Operation": "NY",    # valid -> unchanged
            "Rating State": "??",          # invalid -> snapped
            "Org Type": "NotAType",        # invalid -> snapped
            "ZIP Code": "6103",            # padded to 06103
            "New / Renewal": "New Business",
            "GL (Yes/No)": "no",
        },
        {
            "Test ID": "Y",
            "Contact Number": "646 555 0198",  # already canonical
            "State": "TX",                     # valid
            "Org Type": "LLC",                 # valid
            "ZIP Code": "abcd",                # no digits -> left unchanged
        },
    ]


def _run(flag: bool, rows):
    original = rb_config.RULEBOOK_ENABLED
    rb_config.RULEBOOK_ENABLED = flag
    try:
        random.seed(1234)
        return RrgHandler()._fix_policy_info(copy.deepcopy(rows))
    finally:
        rb_config.RULEBOOK_ENABLED = original


def test_engine_path_matches_handler_path():
    rows = _fixture_rows()
    pure = _run(False, rows)
    engine = _run(True, rows)
    assert engine == pure


def test_valid_values_unchanged_and_invalid_snapped():
    out = _run(True, _fixture_rows())
    # invalid state snapped into the allowed set
    assert out[0]["State"] in ("NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ")
    assert out[0]["Rating State"] in ("NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ")
    assert out[0]["State of Operation"] == "NY"  # valid, untouched
    # contact normalized to spaced 10-digit
    assert out[0]["Contact Number"] == "212 555 0147"
    assert out[1]["Contact Number"] == "646 555 0198"
    # zip padded; no-digit zip left as the original string
    assert out[0]["ZIP Code"] == "06103"
    assert out[1]["ZIP Code"] == "abcd"
    # org snapped / preserved
    assert out[1]["Org Type"] == "LLC"


def test_test_id_sequencing_preserved():
    out = _run(True, _fixture_rows())
    assert out[0]["Test ID"] == "DS-01"
    assert out[1]["Test ID"] == "DS-02"
