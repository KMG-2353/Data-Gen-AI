"""Unit tests for profile composition and selection (U3)."""
from app.rulebook.l0_base import l0_rules
from app.rulebook.profiles import (
    Profile,
    compose_rules,
    generic_profile,
    select_profile,
)
from app.rulebook.rules import EnumRule


def _ids(rules):
    return [r.id for r in rules]


def test_generic_profile_composes_to_l0_only():
    composed = generic_profile().compose()
    assert _ids(composed) == _ids(l0_rules())  # exact L0 pool, nothing added/dropped


def test_unknown_type_selects_generic():
    composed = compose_rules("SOMETHING_ELSE")
    assert _ids(composed) == _ids(l0_rules())


def test_profile_drop_removes_rule():
    composed = compose_rules("RRG")
    assert "l0.ssn" not in _ids(composed)  # RRG drops the SSN rule
    assert "l0.phone" in _ids(composed)  # but inherits the rest


def test_profile_override_replaces_base_rule():
    override = EnumRule(id="l0.state", field_keywords=("state",), allowed=("NY", "TX"))
    profile = Profile(policy_type="X", overrides={"l0.state": override})
    composed = {r.id: r for r in profile.compose()}
    assert composed["l0.state"] is override  # base rule replaced by override
    assert isinstance(composed["l0.state"], EnumRule)


def test_profile_added_l1_rules_appended():
    extra = EnumRule(id="rrg.org", field_keywords=("org",), allowed=("LLC",))
    profile = Profile(policy_type="X", added=[extra])
    composed = profile.compose()
    assert composed[-1] is extra
    assert len(composed) == len(l0_rules()) + 1


def test_select_profile_maps_detected_type():
    assert select_profile("RRG").policy_type == "RRG"
    assert select_profile("IMS").policy_type == "IMS"
    assert select_profile("GENERIC").policy_type == "GENERIC"
    assert select_profile(None).policy_type == "GENERIC"


def test_dropped_l0_rule_still_applies_to_inheriting_template():
    # AE4: l0.ssn dropped by RRG but present in a template that inherits it.
    rrg_ids = _ids(compose_rules("RRG"))
    generic_ids = _ids(compose_rules("GENERIC"))
    assert "l0.ssn" not in rrg_ids
    assert "l0.ssn" in generic_ids
