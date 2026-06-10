"""Unit tests for the scenario parser (U5)."""
from app.rulebook.scenario import parse_scenarios


def test_single_insured_vehicles_and_class_codes():
    result = parse_scenarios("1 insured: 20 vehicles, 8 class codes")
    assert len(result) == 1
    assert result.specs[0].counts == {"vehicles": 20, "class_of_business": 8}
    assert result.adjustments == []


def test_multiple_scenarios_produce_ordered_specs():
    text = "1 insured: 20 vehicles, 8 class codes; 1 insured: 3 vehicles, 2 locations"
    result = parse_scenarios(text)
    assert len(result) == 2
    assert result.specs[0].counts == {"vehicles": 20, "class_of_business": 8}
    assert result.specs[1].counts == {"vehicles": 3, "locations": 2}


def test_count_above_max_is_capped_and_surfaced():
    # AE3: 25 vehicles with max 20 -> capped to 20 with an adjustment note.
    result = parse_scenarios("1 insured: 25 vehicles")
    assert result.specs[0].counts["vehicles"] == 20
    assert len(result.adjustments) == 1
    assert "25" in result.adjustments[0] and "20" in result.adjustments[0]


def test_class_of_business_capped_at_eight():
    result = parse_scenarios("1 insured: 12 class of business")
    assert result.specs[0].counts["class_of_business"] == 8
    assert result.adjustments


def test_no_scenario_phrasing_returns_empty():
    assert len(parse_scenarios("effective date 2020-2026, only CA TX NY")) == 0
    assert not parse_scenarios("")
    assert not parse_scenarios(None)


def test_synonyms_and_casing():
    result = parse_scenarios("1 Insured: 2 Vehicles, 3 class code, 5 Locations")
    assert result.specs[0].counts == {
        "vehicles": 2,
        "class_of_business": 3,
        "locations": 5,
    }


def test_n_insureds_replicates_spec():
    result = parse_scenarios("2 insureds: 5 vehicles")
    assert len(result) == 2
    assert all(s.counts == {"vehicles": 5} for s in result.specs)


def test_class_of_business_phrase_not_shadowed_by_class_code():
    result = parse_scenarios("1 insured: 4 class of business")
    assert result.specs[0].counts == {"class_of_business": 4}
