"""End-to-end defect coverage for the SPG Inland Marine (IM) handler.

Each test reproduces a logged QA defect (DF-IM-001 .. DF-IM-008) by feeding the
handler the exact kind of bad row the LLM produced, then asserts the
post-processor repairs it. Detection is verified separately so the IM template
routes to this handler at all.
"""
import random

import pytest

from app.llm_service import detect_policy_type, detect_policy_type_from_headers
from app.policies import get_handler
from app.policies.im import ImHandler, _parse_date


@pytest.fixture(autouse=True)
def _seed():
    random.seed(1234)


@pytest.fixture
def handler():
    return ImHandler()


# ---------------------------------------------------------------------------
# Detection: the IM rater must route to the IM handler
# ---------------------------------------------------------------------------

def test_detect_from_filename():
    assert detect_policy_type("SPG_IM_Rater_Sample_Headers.xlsx") == "IM"


def test_detect_from_headers():
    headers = {
        "Policy Info": ["Test ID", "Scheduled Equipment Coverage"],
        "Equipment Schedule": ["Test ID", "Value ($)"],
        "Misc Articles Schedule": ["Test ID", "Total Value of Miscellaneous Articles ($)"],
    }
    assert detect_policy_type_from_headers(headers) == "IM"


def test_handler_registered():
    assert get_handler("IM").policy_type == "IM"


# ---------------------------------------------------------------------------
# DF-IM-001 / DF-IM-002 — Date of Quote must be <= Effective Date (Rule 9)
# ---------------------------------------------------------------------------

def test_date_of_quote_clamped_after_effective(handler):
    rows = [
        # DF-IM-001: quote 05/29/2026 after effective 05/27/2026
        {"Test ID": "PI-004", "Effective Date": "05/27/2026", "Date of Quote": "05/29/2026",
         "Type of Entity": "LLC"},
        # DF-IM-002: quote 05/31/2026 after effective 05/22/2026
        {"Test ID": "PI-005", "Effective Date": "05/22/2026", "Date of Quote": "05/31/2026",
         "Type of Entity": "Corporation"},
    ]
    out = handler.post_process(rows, "Policy Info", "")
    for row in out:
        eff = _parse_date(row["Effective Date"])
        quote = _parse_date(row["Date of Quote"])
        assert quote is not None and quote <= eff


def test_valid_quote_date_preserved(handler):
    rows = [{"Test ID": "PI-001", "Effective Date": "05/27/2026",
             "Date of Quote": "05/20/2026", "Type of Entity": "LLC"}]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Date of Quote"] == "05/20/2026"


def test_test_ids_stamped_with_global_ts_prefix(handler):
    # GLOBAL RULE: Test IDs default to TS-### (no rulebook mandates PI for IM).
    rows = [
        {"Test ID": "PI-004", "Effective Date": "05/27/2026", "Type of Entity": "LLC"},
        {"Test ID": "whatever", "Effective Date": "05/22/2026", "Type of Entity": "LLC"},
    ]
    out = handler.post_process(rows, "Policy Info", "")
    assert [r["Test ID"] for r in out] == ["TS-001", "TS-002"]


# ---------------------------------------------------------------------------
# DF-IM-003 / DF-IM-004 — no equipment when Scheduled Equipment Coverage = No
# ---------------------------------------------------------------------------

def test_equipment_dropped_for_no_coverage(handler):
    policy = [
        {"Test ID": "PI-001", "Scheduled Equipment Coverage": "Yes"},
        {"Test ID": "PI-002", "Scheduled Equipment Coverage": "No"},   # DF-IM-003
        {"Test ID": "PI-005", "Scheduled Equipment Coverage": "No"},   # DF-IM-004
    ]
    equipment = [
        {"Test ID": "PI-001", "Value ($)": "30000", "Loss Payee?": "No"},
        {"Test ID": "PI-002", "Value ($)": "48750", "Loss Payee?": "No"},
        {"Test ID": "PI-005", "Value ($)": "32900", "Loss Payee?": "No"},
    ]
    out = handler.post_process(
        equipment, "Equipment Schedule", "", {"Policy Info": policy}
    )
    tids = {r["Test ID"] for r in out}
    assert tids == {"PI-001"}


# ---------------------------------------------------------------------------
# DF-IM-006 — Equipment Value ($) must be numeric $25k-$50k (Rule 50)
# ---------------------------------------------------------------------------

def test_equipment_value_coerced_to_number(handler):
    policy = [{"Test ID": "PI-001", "Scheduled Equipment Coverage": "Yes"}]
    equipment = [
        {"Test ID": "PI-001", "Value ($)": "$48,750", "Loss Payee?": "No"},
        {"Test ID": "PI-001", "Value ($)": "$32900", "Loss Payee?": "No"},
    ]
    out = handler.post_process(
        equipment, "Equipment Schedule", "", {"Policy Info": policy}
    )
    for row in out:
        val = row["Value ($)"]
        assert isinstance(val, (int, float)) and not isinstance(val, str)
        assert 25000 < val <= 50000


# ---------------------------------------------------------------------------
# DF-IM-007 — Total Value of Misc Articles must be numeric $0-$10k (Rule 60)
# ---------------------------------------------------------------------------

def test_misc_total_value_coerced_to_number(handler):
    rows = [
        {"Test ID": "PI-001", "Miscellaneous Articles Coverage Selected?": "Yes",
         "Total Value of Miscellaneous Articles ($)": "$4850"},
        {"Test ID": "PI-002", "Miscellaneous Articles Coverage Selected?": "Yes",
         "Total Value of Miscellaneous Articles ($)": "$7,200"},
    ]
    out = handler.post_process(rows, "Misc Articles Schedule", "")
    for row in out:
        val = row["Total Value of Miscellaneous Articles ($)"]
        assert isinstance(val, (int, float)) and not isinstance(val, str)
        assert 0 < val <= 10000


def test_misc_total_blank_when_coverage_no(handler):
    rows = [{"Test ID": "PI-003", "Miscellaneous Articles Coverage Selected?": "No",
             "Total Value of Miscellaneous Articles ($)": "$5000"}]
    out = handler.post_process(rows, "Misc Articles Schedule", "")
    assert out[0]["Total Value of Miscellaneous Articles ($)"] == ""


# ---------------------------------------------------------------------------
# DF-IM-005 — Loss Date within 3 years before Effective Date (Rule 64)
# ---------------------------------------------------------------------------

def test_loss_date_clamped_into_window(handler):
    policy = [{"Test ID": "PI-004", "Effective Date": "05/27/2026"}]
    loss = [
        # DF-IM-005: 11/09/2022 is ~4 years before effective -> out of window
        {"Test ID": "PI-004", "Any Losses in Past 3 Years?": "Yes", "#": 1,
         "Loss Date": "11/09/2022", "Amount ($)": "$18450"},
    ]
    out = handler.post_process(loss, "Loss History", "", {"Policy Info": policy})
    eff = _parse_date("05/27/2026")
    ld = _parse_date(out[0]["Loss Date"])
    assert ld is not None
    assert (eff.toordinal() - ld.toordinal()) <= 3 * 365
    assert ld < eff


# ---------------------------------------------------------------------------
# DF-IM-008 — Loss Amount ($) must be a positive number (Rule 67)
# ---------------------------------------------------------------------------

def test_loss_amount_coerced_to_number(handler):
    policy = [{"Test ID": "PI-001", "Effective Date": "05/27/2026"}]
    loss = [
        {"Test ID": "PI-001", "Any Losses in Past 3 Years?": "Yes", "#": 1,
         "Loss Date": "01/15/2025", "Amount ($)": "$18450"},
        {"Test ID": "PI-001", "Any Losses in Past 3 Years?": "Yes", "#": 2,
         "Loss Date": "03/10/2024", "Amount ($)": "$12,780"},
    ]
    out = handler.post_process(loss, "Loss History", "", {"Policy Info": policy})
    for row in out:
        amt = row["Amount ($)"]
        assert isinstance(amt, (int, float)) and not isinstance(amt, str)
        assert amt > 0


def test_loss_fields_blanked_when_no_losses(handler):
    policy = [{"Test ID": "PI-002", "Effective Date": "05/27/2026"}]
    loss = [{"Test ID": "PI-002", "Any Losses in Past 3 Years?": "No", "#": 1,
             "Loss Date": "01/15/2025", "Amount ($)": "$5000", "Type of Loss": "Fire"}]
    out = handler.post_process(loss, "Loss History", "", {"Policy Info": policy})
    assert out[0]["Loss Date"] == ""
    assert out[0]["Amount ($)"] == ""
    assert out[0]["Type of Loss"] == ""


# ---------------------------------------------------------------------------
# Scenario architecture: Test Scenario Details summarises real counts
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# DF-IM-013 — Type of Entity must exercise the full dropdown, not just Corp/LLC
# ---------------------------------------------------------------------------

def test_type_of_entity_varies_across_records(handler):
    # The LLM collapsed every insured onto Corporation/LLC.
    rows = [
        {"Test ID": "PI-001", "Effective Date": "05/27/2026", "Type of Entity": "Corporation",
         "Insured Full Name": "Acme Co", "Mailing Address State": "VA"}
        for _ in range(8)
    ]
    out = handler.post_process(rows, "Policy Info", "")
    seen = {r["Type of Entity"] for r in out}
    # More than the two values QA observed, and all within the IM dropdown.
    from app.policies.im import _IM_ENTITY_TYPES
    assert seen.issubset(set(_IM_ENTITY_TYPES))
    assert len(seen) >= 4


def test_trust_rows_are_completed_not_left_empty(handler):
    rows = [
        {"Test ID": "PI-001", "Type of Entity": "x", "Insured Full Name": "Bob Smith",
         "Mailing Address City": "Richmond", "Mailing Address State": "VA",
         "Trustee Full Name": "", "Trustee Address City": "", "Trustee Address State": ""}
        for _ in range(8)
    ]
    out = handler.post_process(rows, "Policy Info", "")
    for r in out:
        if str(r["Type of Entity"]).lower() == "trust":
            assert r["Trustee Full Name"].strip()          # populated
        else:
            assert r["Trustee Full Name"] == ""            # blank for non-Trust


# ---------------------------------------------------------------------------
# DF-IM-009 — Override Reason only when Fee Override holds a value
# ---------------------------------------------------------------------------

def test_override_reason_cleared_when_no_fee_override(handler):
    rows = [
        {"Test ID": "PI-001", "Type of Entity": "LLC", "Fee Override": "",
         "Override Reason": "Some stray reason"},
        {"Test ID": "PI-002", "Type of Entity": "LLC", "Fee Override": "150",
         "Override Reason": ""},
    ]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Override Reason"] == ""        # no override -> blank
    assert out[1]["Override Reason"].strip()      # override present -> populated


# ---------------------------------------------------------------------------
# DF-IM-017 — Equipment: multiple rows per insured, capped at 20
# ---------------------------------------------------------------------------

def test_equipment_count_request_is_multiple_per_insured(handler):
    policy = [{"Test ID": f"TS-00{i}"} for i in range(1, 4)]  # 3 insureds
    count, _rules = handler.build_sheet_context("Equipment Schedule", policy, None, 3)
    assert count >= 6                                          # not 1-per-insured


def test_equipment_capped_at_20_per_insured(handler):
    policy = [{"Test ID": "PI-001", "Scheduled Equipment Coverage": "Yes"}]
    equipment = [{"Test ID": "PI-001", "Value ($)": "30000", "Loss Payee?": "No"}
                 for _ in range(25)]
    out = handler.post_process(equipment, "Equipment Schedule", "",
                               {"Policy Info": policy})
    assert len(out) == 20


# ---------------------------------------------------------------------------
# DF-IM-018 — Loss Payees: 1..10 per policy
# ---------------------------------------------------------------------------

def test_loss_payees_capped_at_10(handler):
    rows = [{"Test ID": "PI-001", "State": "VA"} for _ in range(14)]
    out = handler.post_process(rows, "IM LossPayees", "")
    assert len(out) == 10


def test_loss_payees_count_request_is_multiple(handler):
    policy = [{"Test ID": "TS-001"}, {"Test ID": "TS-002"}]
    count, _ = handler.build_sheet_context("IM LossPayees", policy, None, 2)
    assert count >= 4


# ---------------------------------------------------------------------------
# DF-IM-010/011/012/014/015/016 — phone/fax format + address-state variety
# are enforced by the universal SPG variety pass (app/rulebook/variety.py),
# which main.py runs after this handler for every IM sheet.
# ---------------------------------------------------------------------------

def test_universal_pass_fixes_im_phone_and_states():
    from app.rulebook.variety import enforce_variety_fields
    from app.rulebook.primitives import is_us_state
    rows = [
        {"Test ID": "TS-001", "Binding State": "VA", "Agent Address State": "VA",
         "Mailing Address State": "VA", "Coverage Address State": "VA",
         "Agent Phone": "8045551234", "Agent Fax": "8045559999"},
        {"Test ID": "TS-002", "Binding State": "MD", "Agent Address State": "MD",
         "Mailing Address State": "MD", "Coverage Address State": "MD",
         "Agent Phone": "3015550000", "Agent Fax": "3015551111"},
    ]
    enforce_variety_fields(rows)
    for r in rows:
        assert r["Agent Phone"].startswith("(")          # DF-IM-010/014
        assert r["Agent Fax"].startswith("(")            # DF-IM-011
        for c in ("Agent Address State", "Mailing Address State", "Coverage Address State"):
            assert is_us_state(r[c])                      # DF-IM-012/015/016
            assert r[c] != r["Binding State"]            # not collapsed onto binding


def test_scenario_details_built_from_upstream(handler):
    previous = {
        "Policy Info": [
            {"Test ID": "PI-001", "Binding State": "VA", "Type of Entity": "LLC"},
            {"Test ID": "PI-002", "Binding State": "MD", "Type of Entity": "Trust"},
        ],
        "Equipment Schedule": [
            {"Test ID": "PI-001"}, {"Test ID": "PI-001"},
        ],
        "Additional Interests": [{"Test ID": "PI-001"}],
        "Loss History": [
            {"Test ID": "PI-002", "Loss Date": "01/15/2025"},
            {"Test ID": "PI-002", "Loss Date": ""},  # blank -> not counted
        ],
    }
    headers = ["Scenario ID", "State", "Type of Entity",
               "Equipment Count", "Additional Interest Count", "Loss Count"]
    out = handler.pre_generate(
        "Test Scenario Details", headers, None, None, None, previous
    )
    by_id = {r["Scenario ID"]: r for r in out}
    assert by_id["PI-001"]["Equipment Count"] == 2
    assert by_id["PI-001"]["Additional Interest Count"] == 1
    assert by_id["PI-001"]["Loss Count"] == 0
    assert by_id["PI-002"]["Loss Count"] == 1
    assert by_id["PI-002"]["Type of Entity"] == "Trust"
