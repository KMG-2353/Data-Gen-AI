"""End-to-end defect coverage for the SPG Dwelling Fire (DW) handler.

Each test reproduces a logged QA defect (DEF-001 .. DEF-010) by feeding the
handler the exact kind of bad row the LLM produced, then asserts the
post-processor repairs it. Detection is verified separately so the DW template
routes to this handler at all. Allowed dropdown values come from the rater
template's "15_LKP_Dropdowns" sheet (mirrored as module constants in spg_pl).
"""
import datetime as dt
import random

import pytest

from app.llm_service import detect_policy_type, detect_policy_type_from_headers
from app.policies import get_handler
from app.policies.spg_pl import (
    DwHandler,
    _DW_SIDING,
    _TERM_REASON_NONE,
    _WIND_HAIL_DEDUCTIBLES,
)
from app.rulebook.primitives import parse_date as _parse_date


@pytest.fixture(autouse=True)
def _seed():
    random.seed(1234)


@pytest.fixture
def handler():
    return DwHandler()


# ---------------------------------------------------------------------------
# Detection: the DW rater must route to the DW handler
# ---------------------------------------------------------------------------

def test_detect_from_filename():
    assert detect_policy_type("SPG_PL_DW_Rater_Sample_Headers.xlsx") == "DW"


def test_detect_from_headers():
    headers = {
        "Policy Info": ["Test ID", "Product Selected"],
        "DF Policy": ["Test ID", "Reason for Termination"],
        "DF Locations": ["Test ID", "Type of Siding Material"],
    }
    assert detect_policy_type_from_headers(headers) == "DW"


def test_handler_registered():
    assert get_handler("DW").policy_type == "DW"


# ---------------------------------------------------------------------------
# DEF-001 — Product Selected must always be "Dwelling Fire DP-3"
# ---------------------------------------------------------------------------

def test_product_pinned(handler):
    rows = [
        {"Test ID": "x", "Product Selected": "HO-3 Homeowners", "Effective Date": "05/27/2026"},
        {"Test ID": "y", "Product Selected": "", "Effective Date": "05/22/2026"},
    ]
    out = handler.post_process(rows, "Policy Info", "")
    assert {r["Product Selected"] for r in out} == {"Dwelling Fire DP-3"}


# ---------------------------------------------------------------------------
# DEF-002 — Quote Date must be today's date
# ---------------------------------------------------------------------------

def test_quote_date_is_today(handler):
    rows = [{"Test ID": "x", "Effective Date": "05/27/2026", "Quote Date": "01/01/2020"}]
    out = handler.post_process(rows, "Policy Info", "")
    assert out[0]["Quote Date"] == dt.date.today().strftime("%m/%d/%Y")


# ---------------------------------------------------------------------------
# DEF-004 — Reason for Termination is a mandatory dropdown (never blank
# when prior insurance exists); blank only when there is no prior insurance.
# ---------------------------------------------------------------------------

def test_reason_for_termination_filled_when_not_terminated(handler):
    rows = [{
        "Test ID": "x",
        "Prior Insurance on This Account?": "Yes",
        "Terminated at Company Request?": "No",
        "Reason for Termination": "",
    }]
    out = handler.post_process(rows, "DF Policy", "")
    assert out[0]["Reason for Termination"] == _TERM_REASON_NONE


def test_reason_for_termination_blank_when_no_prior_insurance(handler):
    rows = [{
        "Test ID": "x",
        "Prior Insurance on This Account?": "No",
        "Terminated at Company Request?": "No",
        "Reason for Termination": "Loss History",
    }]
    out = handler.post_process(rows, "DF Policy", "")
    assert out[0]["Reason for Termination"] == ""


# ---------------------------------------------------------------------------
# DEF-005 — Previous Wind/Hail Deductible must be a 0-5 code, not a dollar amount
# ---------------------------------------------------------------------------

def test_wind_hail_deductible_coerced_to_code(handler):
    rows = [{
        "Test ID": "x",
        "Prior Insurance on This Account?": "Yes",
        "Previous Wind/Hail Deductible": "$2,500",
    }]
    out = handler.post_process(rows, "DF Policy", "")
    assert out[0]["Previous Wind/Hail Deductible"] in _WIND_HAIL_DEDUCTIBLES


# ---------------------------------------------------------------------------
# DEF-006 — Management Company Phone must be U.S. formatted
# ---------------------------------------------------------------------------

def test_management_phone_formatted(handler):
    rows = [{
        "Test ID": "x",
        "Prior Insurance on This Account?": "Yes",
        "Management Company Phone": "8045551234",
    }]
    out = handler.post_process(rows, "DF Policy", "")
    assert out[0]["Management Company Phone"] == "(804) 555-1234"


# ---------------------------------------------------------------------------
# DEF-007 — up to 20 locations per insured (never more)
# ---------------------------------------------------------------------------

def test_locations_capped_at_20(handler):
    rows = [{"Test ID": "x", "Loc #": i} for i in range(1, 26)]
    out = handler.post_process(rows, "DF Locations", "", {"Policy Info": [{"Test ID": "x"}]})
    assert len([r for r in out if r["Test ID"] == "x"]) == 20


# ---------------------------------------------------------------------------
# DEF-008 — Siding Material must be a valid dropdown value ("Brick" invalid)
# ---------------------------------------------------------------------------

def test_siding_brick_coerced(handler):
    rows = [{"Test ID": "x", "Type of Siding Material": "Brick"}]
    out = handler.post_process(rows, "DF Locations", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Type of Siding Material"] in _DW_SIDING


def test_siding_valid_value_preserved(handler):
    rows = [{"Test ID": "x", "Type of Siding Material": "Masonry"}]
    out = handler.post_process(rows, "DF Locations", "", {"Policy Info": [{"Test ID": "x"}]})
    assert out[0]["Type of Siding Material"] == "Masonry"


# ---------------------------------------------------------------------------
# DEF-009 — Loss Payees capped at 10 per policy
# ---------------------------------------------------------------------------

def test_loss_payees_capped_at_10(handler):
    rows = [{"Test ID": "x", "Is Mortgagee?": "Yes", "Loan Number": f"L{i}"} for i in range(15)]
    out = handler.post_process(rows, "DF LossPayees", "", {"Policy Info": [{"Test ID": "x"}]})
    assert len([r for r in out if r["Test ID"] == "x"]) == 10


# ---------------------------------------------------------------------------
# DEF-010 — Loss History (Loss Schedule) capped at 10 per policy
# ---------------------------------------------------------------------------

def test_loss_history_capped_at_10(handler):
    policy = [{"Test ID": "x", "Effective Date": "05/27/2026"}]
    rows = [{
        "Test ID": "x", "Any Losses in Past 5 Years?": "Yes", "#": i,
        "Loss Date": "01/15/2025", "Amount": "$5000",
    } for i in range(15)]
    out = handler.post_process(rows, "DF Loss history", "", {"Policy Info": policy})
    assert len([r for r in out if r["Test ID"] == "x"]) == 10
