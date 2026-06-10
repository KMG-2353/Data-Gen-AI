"""Unit tests for the engine: prompt emit + validator application (U1)."""
from app.rulebook import (
    EnumRule,
    FormatRule,
    RuleContext,
    apply_validators,
    emit_prompt_constraints,
)
from app.rulebook.conditions import sheet_is


def _phone_fmt(v):
    digits = "".join(c for c in str(v or "") if c.isdigit())
    if len(digits) != 10:
        return v
    return f"{digits[:3]} {digits[3:6]} {digits[6:]}"


def test_emit_returns_only_firing_fragments_for_sheet():
    rules = [
        EnumRule(
            id="state",
            field_keywords=("state",),
            allowed=("NY",),
            condition=sheet_is("policy information"),
            prompt_text="- State: one of [NY]",
        ),
        EnumRule(
            id="veh",
            field_keywords=("vehicle type",),
            allowed=("Trailer",),
            condition=sheet_is("sched of vehicle"),
            prompt_text="- Vehicle Type: one of [Trailer]",
        ),
    ]
    out = emit_prompt_constraints(rules, "Policy Information")
    assert "State: one of [NY]" in out
    assert "Vehicle Type" not in out  # vehicle rule does not fire for this sheet


def test_apply_validators_snaps_matching_rows():
    rules = [
        EnumRule(id="state", field_keywords=("state",), allowed=("NY", "TX")),
        FormatRule(id="phone", field_keywords=("contact",), formatter=_phone_fmt),
    ]
    rows = [
        {"Rating State": "ZZ", "Contact Number": "2125550147"},
        {"Rating State": "TX", "Contact Number": "646 555 0198"},
    ]
    apply_validators(rules, rows, sheet="Policy Information")
    assert rows[0]["Rating State"] == "NY"
    assert rows[0]["Contact Number"] == "212 555 0147"
    assert rows[1]["Rating State"] == "TX"
    assert rows[1]["Contact Number"] == "646 555 0198"


def test_apply_validators_is_idempotent():
    rules = [
        EnumRule(id="state", field_keywords=("state",), allowed=("NY", "TX")),
        FormatRule(id="phone", field_keywords=("contact",), formatter=_phone_fmt),
    ]
    rows = [{"Rating State": "ZZ", "Contact Number": "2125550147"}]
    once = apply_validators(rules, [dict(rows[0])], sheet="X")
    twice = apply_validators(rules, [dict(once[0])], sheet="X")
    assert once == twice


def test_apply_validators_respects_sheet_condition():
    rule = EnumRule(
        id="state",
        field_keywords=("state",),
        allowed=("NY",),
        condition=sheet_is("policy information"),
    )
    rows = [{"Rating State": "ZZ"}]
    apply_validators([rule], rows, sheet="Sched of Vehicles")
    assert rows[0]["Rating State"] == "ZZ"  # condition did not fire, no snap
