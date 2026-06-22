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
