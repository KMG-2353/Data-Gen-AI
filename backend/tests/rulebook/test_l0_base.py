"""Unit tests for the L0 generic base rule pool (U2)."""
from app.rulebook import RuleContext
from app.rulebook.l0_base import l0_by_id, l0_rules


def _ctx():
    return RuleContext(sheet="Policy Information")


def test_l0_pool_has_stable_ids():
    ids = {r.id for r in l0_rules()}
    assert {"l0.phone", "l0.zip", "l0.state", "l0.ssn", "l0.dollar", "l0.date"} <= ids


def test_phone_validator_strips_separators():
    rule = l0_by_id()["l0.phone"]
    assert rule.validate("(212) 555-0147", _ctx()) == "2125550147"
    assert rule.validate("2125550147", _ctx()) == "2125550147"  # idempotent


def test_zip_validator_pads_and_trims():
    rule = l0_by_id()["l0.zip"]
    assert rule.validate("6103", _ctx()) == "06103"  # leading-zero ZIP preserved
    assert rule.validate("100041234", _ctx()) == "10004"
    assert rule.validate("10004", _ctx()) == "10004"


def test_state_validator_uppercases():
    rule = l0_by_id()["l0.state"]
    assert rule.validate(" ny ", _ctx()) == "NY"


def test_ssn_validator_formats_nine_digits():
    rule = l0_by_id()["l0.ssn"]
    assert rule.validate("123456789", _ctx()) == "123-45-6789"
    assert rule.validate("123-45-6789", _ctx()) == "123-45-6789"  # idempotent


def test_date_and_dollar_are_prompt_only():
    rules = l0_by_id()
    # No formatter/enum enforcement — validate is identity (code hook / guidance).
    assert rules["l0.date"].validate("05/28/2026", _ctx()) == "05/28/2026"
    assert rules["l0.dollar"].validate("$1,000", _ctx()) == "$1,000"
    assert rules["l0.date"].prompt_fragment(_ctx())
    assert rules["l0.dollar"].prompt_fragment(_ctx())
