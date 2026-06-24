"""End-to-end defect coverage for the SPG Homeowners (HO) handler.

Each test reproduces a logged QA defect (HO-001 .. HO-012) by feeding the
handler the exact kind of bad row the LLM produced, then asserts the
post-processor repairs it. Detection is verified separately so the HO template
routes to this handler at all. Allowed dropdown values come from the rater
template's "15_LKP_Dropdowns" sheet (mirrored as module constants in spg_pl).
"""
import datetime as dt
import random

import pytest

from app.llm_service import detect_policy_type, detect_policy_type_from_headers
from app.policies import get_handler
from app.policies.spg_pl import HoHandler, _HO_ROOFING, _HO_SIDING


@pytest.fixture(autouse=True)
def _seed():
    random.seed(1234)


@pytest.fixture
def handler():
    return HoHandler()


# ---------------------------------------------------------------------------
# Detection: the HO rater must route to the HO handler
# ---------------------------------------------------------------------------

def test_detect_from_filename():
    assert detect_policy_type("SPG_PL_HO_Rater_Sample_Headers.xlsx") == "HO"


def test_detect_from_headers():
    headers = {
        "Policy Info": ["Test ID", "Product Selected"],
        "HO Dwelling": ["Test ID", "Protection Class", "Type of Roofing Material"],
        "HO Loss History": ["Test ID", "Any Open Claims?"],
    }
    assert detect_policy_type_from_headers(headers) == "HO"


def test_handler_registered():
    assert get_handler("HO").policy_type == "HO"


# ---------------------------------------------------------------------------
# HO-001 / HO-002 — Product pinned to "HO-3 Homeowners"; Quote Date = today
# ---------------------------------------------------------------------------

def test_product_pinned(handler):
    rows = [{"Test ID": "x", "Product Selected": "Dwelling Fire DP-3", "Effective Date": "05/27/2026"}]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Product Selected"] == "HO-3 Homeowners"


def test_quote_date_is_today(handler):
    rows = [{"Test ID": "x", "Effective Date": "05/27/2026", "Quote Date": "01/01/2020"}]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Quote Date"] == dt.date.today().strftime("%m/%d/%Y")


# ---------------------------------------------------------------------------
# HO-003 / HO-004 / HO-005 — phone / fax / insured numbers U.S.-formatted
# ---------------------------------------------------------------------------

def test_phone_fax_insured_formatted(handler):
    rows = [{
        "Test ID": "x",
        "Effective Date": "05/27/2026",
        "Phone Number": "8045551234",
        "Agent Fax": "7035559876",
        "Agent Phone": "5715550000",
    }]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Phone Number"] == "(804) 555-1234"
    assert out[0]["Agent Fax"] == "(703) 555-9876"
    assert out[0]["Agent Phone"] == "(571) 555-0000"


# ---------------------------------------------------------------------------
# HO-006 / HO-007 — Water Source / Fire Dept "Yes" only when Prot Class > 8
# ---------------------------------------------------------------------------

def test_protection_class_dependency_low_class(handler):
    rows = [{
        "Test ID": "x", "Protection Class": "3",
        "Within 1,000 ft of Water Source?": "Yes",
        "Fire Dept Response < 15 Min?": "Yes",
    }]
    out = handler.post_process(rows, "HO Dwelling", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Within 1,000 ft of Water Source?"] == "No"
    assert out[0]["Fire Dept Response < 15 Min?"] == "No"


def test_protection_class_dependency_high_class_preserved(handler):
    rows = [{
        "Test ID": "x", "Protection Class": "9",
        "Within 1,000 ft of Water Source?": "Yes",
        "Fire Dept Response < 15 Min?": "Yes",
    }]
    out = handler.post_process(rows, "HO Dwelling", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Within 1,000 ft of Water Source?"] == "Yes"
    assert out[0]["Fire Dept Response < 15 Min?"] == "Yes"


# ---------------------------------------------------------------------------
# HO-008 / HO-009 — Roofing / Siding must be valid dropdown values
# ---------------------------------------------------------------------------

def test_roofing_coerced(handler):
    rows = [{"Test ID": "x", "Protection Class": "5", "Type of Roofing Material": "Clay Tile"}]
    out = handler.post_process(rows, "HO Dwelling", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Type of Roofing Material"] in _HO_ROOFING


def test_siding_coerced(handler):
    rows = [{"Test ID": "x", "Protection Class": "5", "Type of Siding Material": "Brick"}]
    out = handler.post_process(rows, "HO Dwelling", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Type of Siding Material"] in _HO_SIDING


# ---------------------------------------------------------------------------
# HO-010 — Loss Payees capped at 10
# ---------------------------------------------------------------------------

def test_loss_payees_capped_at_10(handler):
    rows = [{"Test ID": "x", "Is Mortgagee?": "Yes", "Loan Number": f"L{i}"} for i in range(15)]
    out = handler.post_process(rows, "HO LossPayees", "", {"Policy Info": [{"Test ID": "x"}]})
    assert len([r for r in out if r["Test ID"] == "x"]) == 10


# ---------------------------------------------------------------------------
# HO-011 — Any Open Claims? must always be "No"
# ---------------------------------------------------------------------------

def test_any_open_claims_forced_no(handler):
    policy = [{"Test ID": "x", "Effective Date": "05/27/2026"}]
    rows = [{
        "Test ID": "x", "Any Open Claims?": "Yes",
        "Any Losses in Past 5 Years?": "No",
    }]
    out = handler.post_process(rows, "HO Loss History", "", {"Policy Info": policy})
    assert out[0]["Any Open Claims?"] == "No"


# ---------------------------------------------------------------------------
# HO-012 — Loss History (Loss Schedule) capped at 10 per policy
# ---------------------------------------------------------------------------

def test_loss_history_capped_at_10(handler):
    policy = [{"Test ID": "x", "Effective Date": "05/27/2026"}]
    rows = [{
        "Test ID": "x", "Any Open Claims?": "No", "Any Losses in Past 5 Years?": "Yes",
        "#": i, "Loss Date": "01/15/2025", "Amount": "$5000",
    } for i in range(15)]
    out = handler.post_process(rows, "HO Loss History", "", {"Policy Info": policy})
    assert len([r for r in out if r["Test ID"] == "x"]) == 10
