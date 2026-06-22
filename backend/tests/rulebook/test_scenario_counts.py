"""Scenario-driven RRG count tests (U6).

Exercises the count seams directly (build_sheet_context row counts + the
per-insured count helpers) so we can assert cross-sheet behavior without an LLM.
"""
import app.policies.rrg as rrg
from app.policies.rrg import (
    RrgHandler,
    _cob_counts_for_insureds,
    _location_counts_for_insureds,
    _vehicle_counts_for_insureds,
)
from app.rulebook.scenario import parse_scenarios


def _insured(name, **flags):
    row = {"Named Insured": name, "State": "NY", "Rating State": "NY", "ZIP Code": "10004"}
    row.update(flags)
    return row


def test_ae1_single_spec_is_a_template_ui_count_wins():
    # "1 insured: 20 vehicles, 8 class codes" is a per-insured template. The UI
    # test-case count (10) decides how many insureds exist, NOT the scenario.
    specs = parse_scenarios("1 insured: 20 vehicles, 8 class codes").specs
    handler = RrgHandler()

    n_policy, _ = handler.build_sheet_context(
        "Policy Information", policy_data=None, driver_data=None,
        original_row_count=10, special_instruction="1 insured: 20 vehicles, 8 class codes",
    )
    assert n_policy == 10  # UI count wins; scenario is a template, not a count

    insureds = [_insured("Acme", **{"Auto (Yes/No)": "Yes", "PL (Yes/No)": "Yes"})]
    veh = _vehicle_counts_for_insureds(insureds, insureds, specs)
    cob = _cob_counts_for_insureds(insureds, insureds, specs)
    assert veh == [20]
    assert cob == [8]


def test_single_spec_broadcasts_to_all_insureds():
    # The exact reported scenario: "One insured with 20 vehicles, 20 locations,
    # 8 classes of business" + UI selected 5 test cases -> all 5 insureds get
    # 20 vehicles / 20 locations / 8 CoB (the template broadcasts to every one).
    specs = parse_scenarios(
        "One insured with 20 vehicles, 20 locations, and 8 classes of business"
    ).specs
    assert len(specs) == 1
    insureds = [
        _insured(f"I{i}", **{"Auto (Yes/No)": "Yes", "PL (Yes/No)": "Yes"})
        for i in range(5)
    ]
    assert _vehicle_counts_for_insureds(insureds, insureds, specs) == [20] * 5
    assert _location_counts_for_insureds(insureds, specs) == [20] * 5
    assert _cob_counts_for_insureds(insureds, insureds, specs) == [8] * 5


def test_ae2_no_scenario_matches_checkpoint_behavior():
    insureds = [_insured(f"I{i}") for i in range(8)]
    # Without specs the helpers return the checkpoint coverage spread, unchanged.
    assert _location_counts_for_insureds(insureds, None) == _location_counts_for_insureds(insureds)
    assert _vehicle_counts_for_insureds(insureds, insureds, None) == _vehicle_counts_for_insureds(insureds)
    # Spread reaches the documented max (20) at high volume.
    assert max(_location_counts_for_insureds(insureds)) == 20


def test_two_scenarios_two_insureds_with_respective_counts():
    text = "1 insured: 20 vehicles, 8 class codes; 1 insured: 3 vehicles, 2 locations"
    specs = parse_scenarios(text).specs
    insureds = [
        _insured("A", **{"Auto (Yes/No)": "Yes", "PL (Yes/No)": "Yes"}),
        _insured("B", **{"Auto (Yes/No)": "Yes", "PL (Yes/No)": "Yes"}),
    ]
    veh = _vehicle_counts_for_insureds(insureds, insureds, specs)
    loc = _location_counts_for_insureds(insureds, specs)
    assert veh == [20, 3]
    # insured A has no location spec -> falls back to coverage; B requested 2.
    assert loc[1] == 2


def test_vehicle_subset_indices_map_to_full_list():
    # Only the second insured is Auto=Yes; its vehicle count must come from spec[1].
    text = "1 insured: 5 vehicles; 1 insured: 12 vehicles"
    specs = parse_scenarios(text).specs
    full = [_insured("A"), _insured("B", **{"Auto (Yes/No)": "Yes"})]
    auto_subset = [full[1]]  # same object identity as in full
    veh = _vehicle_counts_for_insureds(auto_subset, full, specs)
    assert veh == [12]


def test_scenario_disabled_when_flag_off(monkeypatch):
    from app.rulebook import config as rb_config
    monkeypatch.setattr(rb_config, "RULEBOOK_ENABLED", False)
    n_policy, _ = RrgHandler().build_sheet_context(
        "Policy Information", policy_data=None, driver_data=None,
        original_row_count=10, special_instruction="1 insured: 20 vehicles",
    )
    assert n_policy == 10  # flag off -> scenario ignored, row_count honored
