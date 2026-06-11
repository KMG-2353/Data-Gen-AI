"""End-to-end integration of the rulebook path through the handlers (U8).

The real /api/generate path calls the LLM, so these tests exercise the engine
wiring at the handler seams (profile selection, scenario counts, validators,
flag) which is the deterministic, testable surface.
"""
import copy
import random

import pytest

from app.rulebook import config as rb_config
from app.rulebook.profiles import select_profile
from app.policies import get_handler
from app.policies.ims import _enforce_property_sheet_values
from app.policies.rrg import RrgHandler, _vehicle_counts_for_insureds


@pytest.fixture(autouse=True)
def _engine_on():
    original = rb_config.RULEBOOK_ENABLED
    rb_config.RULEBOOK_ENABLED = True
    yield
    rb_config.RULEBOOK_ENABLED = original


def test_profile_selection_matches_detected_type():
    assert select_profile("RRG").policy_type == "RRG"
    assert select_profile("IMS").policy_type == "IMS"
    assert select_profile("GENERIC").policy_type == "GENERIC"


def test_rrg_with_scenario_drives_cross_sheet_counts():
    handler = RrgHandler()
    si = "1 insured: 20 vehicles, 8 class codes"
    n_policy, _ = handler.build_sheet_context(
        "Policy Information", None, None, original_row_count=10, special_instruction=si,
    )
    assert n_policy == 1  # scenario overrides row_count
    insureds = [{"Named Insured": "A", "Auto (Yes/No)": "Yes", "State": "NY",
                 "Rating State": "NY", "ZIP Code": "10004"}]
    from app.rulebook.scenario import parse_scenarios
    specs = parse_scenarios(si).specs
    assert _vehicle_counts_for_insureds(insureds, insureds, specs) == [20]


def test_rrg_without_scenario_unchanged():
    handler = RrgHandler()
    n_policy, _ = handler.build_sheet_context(
        "Policy Information", None, None, original_row_count=7, special_instruction="",
    )
    assert n_policy == 7  # no scenario -> row_count honored


def test_rrg_post_process_enforces_extracted_enums():
    handler = RrgHandler()
    rows = [{"Test ID": "x", "State": "ZZ", "Contact Number": "2125550147",
             "ZIP Code": "6103"}]
    out = handler.post_process(rows, "Policy Information", special_instruction="")
    assert out[0]["State"] in ("NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ")
    assert out[0]["Contact Number"] == "212 555 0147"
    assert out[0]["ZIP Code"] == "06103"


def test_ims_property_enums_enforced_via_engine():
    rows = [{"Coinsurance": "70", "Cause of Loss": "Comprehensive"}]
    out = _enforce_property_sheet_values(copy.deepcopy(rows))
    assert out[0]["Coinsurance"] in ("80", "90", "100")
    assert out[0]["Cause of Loss"] in ("Basic", "Special", "Broad")


def test_generic_template_is_l0_only_no_errors():
    composed = select_profile("GENERIC").compose()
    assert composed  # has L0 rules
    assert all(r.id.startswith("l0.") for r in composed)


def test_flag_off_falls_back_to_pure_handler():
    rb_config.RULEBOOK_ENABLED = False
    handler = RrgHandler()
    # scenario ignored under flag off
    n_policy, _ = handler.build_sheet_context(
        "Policy Information", None, None, original_row_count=10,
        special_instruction="1 insured: 20 vehicles",
    )
    assert n_policy == 10

    # RRG post-process still snaps (handler inline path), under same seed parity.
    rows = [{"Test ID": "x", "State": "ZZ", "Contact Number": "2125550147"}]
    random.seed(7)
    out = handler.post_process(copy.deepcopy(rows), "Policy Information", "")
    assert out[0]["State"] in ("NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ")
    assert out[0]["Contact Number"] == "212 555 0147"
