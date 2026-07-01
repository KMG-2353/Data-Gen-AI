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


# ---------------------------------------------------------------------------
# HO-017 — "Is Dwelling a Manufactured Home?" must exercise Yes, not always No
# ---------------------------------------------------------------------------

def test_manufactured_home_varies(handler):
    rows = [
        {"Test ID": "TS-01", "Is Dwelling a Manufactured Home?": "No",
         "Square Footage": "1800", "Protection Class": "5"}
        for _ in range(9)
    ]
    out = handler.post_process(rows, "HO Dwelling", "")
    seen = {r["Is Dwelling a Manufactured Home?"] for r in out}
    assert "Yes" in seen and "No" in seen           # both values now appear


def test_manufactured_home_yes_meets_min_sqft(handler):
    from app.policies.spg_pl import _MANUFACTURED_MIN_SQFT
    from app.rulebook.primitives import to_number
    rows = [
        {"Test ID": "TS-01", "Is Dwelling a Manufactured Home?": "No",
         "Square Footage": "900", "Protection Class": "5"}
        for _ in range(9)
    ]
    out = handler.post_process(rows, "HO Dwelling", "")
    for r in out:
        if r["Is Dwelling a Manufactured Home?"] == "Yes":
            assert to_number(r["Square Footage"]) >= _MANUFACTURED_MIN_SQFT


# ---------------------------------------------------------------------------
# HO-018 — Coverage F is mandatory (non-blank) when Coverage E > $0
# ---------------------------------------------------------------------------

def test_coverage_f_filled_when_coverage_e_present(handler):
    from app.policies.spg_pl import _HO_COVERAGE_F
    rows = [
        {"Test ID": "TS-01", "Coverage E — Limit of Liability": "$100,000",
         "Coverage F — Increased Medical Payments": ""},        # blank -> must fill
        {"Test ID": "TS-02", "Coverage E — Limit of Liability": "$300,000",
         "Coverage F — Increased Medical Payments": "garbage"},  # invalid -> snap
    ]
    out = handler.post_process(rows, "HO Coverages", "")
    for r in out:
        assert r["Coverage F — Increased Medical Payments"] in _HO_COVERAGE_F


def test_coverage_f_blank_when_coverage_e_excluded(handler):
    rows = [{"Test ID": "TS-01", "Coverage E — Limit of Liability": "Excluded",
             "Coverage F — Increased Medical Payments": "$5,000"}]
    out = handler.post_process(rows, "HO Coverages", "")
    assert out[0]["Coverage F — Increased Medical Payments"] == ""


# ---------------------------------------------------------------------------
# HO-010 / HO-012 — loss payees & loss history: multiple per scenario, capped 10
# ---------------------------------------------------------------------------

def test_ho_loss_payees_capped_at_10(handler):
    rows = [{"Test ID": "TS-01", "State": "VA", "Is Mortgagee?": "No"} for _ in range(15)]
    out = handler.post_process(rows, "HO LossPayees", "")
    assert len(out) == 10


def test_ho_child_counts_request_multiples(handler):
    policy = [{"Test ID": "TS-01"}, {"Test ID": "TS-02"}]
    lp_count, _ = handler.build_sheet_context("HO LossPayees", policy, None, 2)
    lh_count, _ = handler.build_sheet_context("HO LossHistory", policy, None, 2)
    assert lp_count >= 4 and lh_count >= 4


# ---------------------------------------------------------------------------
# HO-013 — large data set: child schedules must cover EVERY test case, not just
# the first few the LLM emitted ("6 loss payees for 50 test cases").
# ---------------------------------------------------------------------------

def _roster(n):
    from app.rulebook.primitives import default_test_id
    # Handler normalizes to zero-padded 2-digit ids (TS-01 … TS-50).
    return [{"Test ID": f"TS-{i:02d}"} for i in range(1, n + 1)]


def test_loss_payees_cover_all_test_cases_on_large_set(handler):
    """LLM produced payees for only the first 6 of 50 insureds; the roster-anchored
    multiplicity pass must synthesise the missing 44 so all 50 are represented."""
    policy = _roster(50)
    rows = [
        {"Test ID": f"TS-{i:02d}", "State": "VA", "Is Mortgagee?": "Yes",
         "Loan Number": f"L{i}-{j}", "Mortgage Current?": "Yes"}
        for i in range(1, 7) for j in range(10)          # only TS-01..TS-06, 10 each
    ]
    out = handler.post_process(rows, "HO LossPayees", "", {"Policy Info": policy})
    covered = {r["Test ID"] for r in out}
    assert covered == {f"TS-{i:02d}" for i in range(1, 51)}   # all 50 present
    for i in range(1, 51):                                    # each within 3–10
        n = len([r for r in out if r["Test ID"] == f"TS-{i:02d}"])
        assert 3 <= n <= 10, f"TS-{i:02d} has {n} payees"
    # Synthesised insureds keep distinct loan numbers (no cross-Test-ID dupes).
    loans = [str(r.get("Loan Number") or "") for r in out
             if str(r.get("Is Mortgagee?", "")).strip().lower() == "yes"
             and str(r.get("Loan Number") or "").strip()]
    assert len(loans) == len(set(loans))


def test_loss_history_cover_all_test_cases_on_large_set(handler):
    policy = [{"Test ID": f"TS-{i:02d}", "Effective Date": "05/27/2026"}
              for i in range(1, 51)]
    rows = [
        {"Test ID": f"TS-{i:02d}", "Any Open Claims?": "No",
         "Any Losses in Past 5 Years?": "Yes", "#": j,
         "Loss Date": "01/15/2025", "Amount": "5000"}
        for i in range(1, 7) for j in range(1, 4)        # only TS-01..TS-06
    ]
    out = handler.post_process(rows, "HO Loss History", "", {"Policy Info": policy})
    covered = {r["Test ID"] for r in out}
    assert covered == {f"TS-{i:02d}" for i in range(1, 51)}
