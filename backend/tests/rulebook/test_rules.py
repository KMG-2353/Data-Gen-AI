"""Unit tests for rule abstractions (U1)."""
from app.rulebook import EnumRule, FormatRule, RuleContext
from app.rulebook.conditions import field_equals
from app.rulebook.rules import CountRule


def _ctx(row=None, sheet="Sheet1"):
    return RuleContext(sheet=sheet, row=row)


def test_format_rule_snaps_malformed_and_leaves_valid():
    def zip5(v):
        digits = "".join(c for c in str(v or "") if c.isdigit())
        return digits.zfill(5)[:5] if digits else v

    rule = FormatRule(id="zip", field_keywords=("zip",), formatter=zip5)
    assert rule.validate("6103", _ctx()) == "06103"
    assert rule.validate("10004", _ctx()) == "10004"  # already canonical, unchanged


def test_enum_rule_snaps_out_of_list_and_leaves_valid():
    rule = EnumRule(id="state", field_keywords=("state",), allowed=("NY", "TX", "FL"))
    assert rule.validate("ZZ", _ctx()) == "NY"  # snapped to first allowed
    assert rule.validate("TX", _ctx()) == "TX"  # valid value untouched


def test_enum_rule_custom_snap():
    rule = EnumRule(
        id="state",
        field_keywords=("state",),
        allowed=("NY", "TX"),
        snap=lambda v, allowed: allowed[-1],
    )
    assert rule.validate("ZZ", _ctx()) == "TX"


def test_apply_to_row_finds_column_by_keywords():
    rule = EnumRule(id="state", field_keywords=("state",), allowed=("NY", "TX"))
    row = {"Rating State": "ZZ", "Other": "x"}
    rule.apply_to_row(row, _ctx(row=row))
    assert row["Rating State"] == "NY"
    assert row["Other"] == "x"  # untouched


def test_apply_to_row_noop_when_column_absent():
    rule = EnumRule(id="state", field_keywords=("state",), allowed=("NY",))
    row = {"Name": "Acme"}
    rule.apply_to_row(row, _ctx(row=row))
    assert row == {"Name": "Acme"}


def test_field_equals_condition_fires_only_on_match():
    rule = EnumRule(
        id="retro",
        field_keywords=("retro",),
        allowed=("X",),
        condition=field_equals(("policy", "basis"), "Claims Made"),
    )
    matching = {"Policy Basis": "Claims Made", "Retro": "ZZ"}
    other = {"Policy Basis": "Occurrence", "Retro": "ZZ"}
    assert rule.fires(_ctx(row=matching)) is True
    assert rule.fires(_ctx(row=other)) is False


def test_count_rule_caps_to_max():
    rule = CountRule(id="vehicles", max_count=20, min_count=1)
    assert rule.cap(25) == (20, True)
    assert rule.cap(8) == (8, False)
    assert rule.cap(0) == (1, False)  # floored to min
