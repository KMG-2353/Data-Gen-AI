"""End-to-end defect coverage for the SPG Cargo (CARGO) and Auto Physical Damage
(APD) handlers.

Both templates previously fell through to the GenericHandler, so their dropdown
fields were never enforced. These tests prove the templates now route to their
handlers and that the logged defects are repaired: phone/fax format + address-
state variety (via the universal pass main.py runs for the SPG family) and the
Cargo Radius dropdown.
"""
import random

import pytest

from app.llm_service import detect_policy_type, detect_policy_type_from_headers
from app.policies import get_handler
from app.policies.spg_auto import CargoHandler, ApdHandler, _CARGO_RADIUS
from app.rulebook.variety import enforce_variety_fields
from app.rulebook.primitives import is_us_state


@pytest.fixture(autouse=True)
def _seed():
    random.seed(1234)


# ---------------------------------------------------------------------------
# Detection + registration: the raters must route to their handlers
# ---------------------------------------------------------------------------

def test_detect_cargo_from_filename():
    assert detect_policy_type("SPG_Cargo_Rater_V1.xlsx") == "CARGO"


def test_detect_apd_from_filename():
    assert detect_policy_type("SPG_APD_Rater_v1.0.xlsx") == "APD"


def test_detect_cargo_from_headers():
    headers = {
        "Policy Information": ["Test ID", "Rating State", "Radius of Operations"],
        "Sched of Commodities": ["Test ID", "Commodity Name"],
        "Sched of Power Units": ["Test ID", "Garaging State"],
    }
    assert detect_policy_type_from_headers(headers) == "CARGO"


def test_detect_apd_from_headers():
    headers = {
        "Policy Information": ["Test ID", "Rating State"],
        "APD Commodities": ["Test ID", "Commodity Name"],
        "APD Trailers": ["Test ID", "Garaging State"],
    }
    assert detect_policy_type_from_headers(headers) == "APD"


def test_handlers_registered():
    assert get_handler("CARGO").policy_type == "CARGO"
    assert get_handler("APD").policy_type == "APD"


# ---------------------------------------------------------------------------
# CARGO-004 — Radius of Operations must use its dropdown, not arbitrary values
# ---------------------------------------------------------------------------

def test_cargo_radius_snapped_to_dropdown():
    handler = CargoHandler()
    rows = [
        {"Test ID": "TS-001", "Commodity Name": "Steel", "Radius of Operations": "350 miles"},
        {"Test ID": "TS-001", "Commodity Name": "Grain", "Radius of Operations": ""},
        {"Test ID": "TS-001", "Commodity Name": "Autos", "Radius of Operations": "100-500 Miles"},
    ]
    out = handler.post_process(rows, "Sched of Commodities", "")
    for r in out:
        assert r["Radius of Operations"] in _CARGO_RADIUS
    # the already-valid value is preserved
    assert out[2]["Radius of Operations"] == "100-500 Miles"


# ---------------------------------------------------------------------------
# Test IDs stamped on the policy sheet
# ---------------------------------------------------------------------------

def test_cargo_policy_test_ids_stamped():
    handler = CargoHandler()
    rows = [{"Test ID": "x", "Named Insured": "A"},
            {"Test ID": "y", "Named Insured": "B"}]
    out = handler.post_process(rows, "Policy Information", "")
    assert [r["Test ID"] for r in out] == ["TS-001", "TS-002"]


# ---------------------------------------------------------------------------
# APD-016 / CARGO-008(quote) — Date of Quote must default to TODAY across all
# SPG LOBs (data-creation date), Effective clamped to >= today, Exp = Eff + 1yr.
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("handler", [CargoHandler(), ApdHandler()])
def test_quote_date_pinned_to_today(handler):
    from datetime import date
    from app.rulebook.primitives import format_date_slash, add_one_year

    today = date.today()
    # A past effective date is clamped up to today; a future one is preserved.
    rows = [
        {"Test ID": "TS-001", "Effective Date": "05/27/2026",
         "Expiration Date": "", "Quote Date": "05/20/2026"},
        {"Test ID": "TS-002", "Effective Date": "12/01/2027",
         "Expiration Date": "", "Quote Date": "01/01/2026"},
    ]
    out = handler.post_process(rows, "Policy Information", "")
    # Quote is today on every row.
    assert all(r["Quote Date"] == format_date_slash(today) for r in out)
    # Past effective clamped to today; expiration re-derived.
    assert out[0]["Effective Date"] == format_date_slash(today)
    assert out[0]["Expiration Date"] == format_date_slash(add_one_year(today))
    # Future effective preserved.
    assert out[1]["Effective Date"] == "12/01/2027"
    assert out[1]["Expiration Date"] == "12/01/2028"


# ---------------------------------------------------------------------------
# Cargo Loss History data rules: Loss Year within the past 3 years (Rule 130 —
# never the effective/future year), and per-policy "Any Unrepaired Damage from
# Prior Losses?" held consistent across an insured's loss rows.
# ---------------------------------------------------------------------------

def test_cargo_loss_year_clamped_and_unrepaired_consistent():
    handler = CargoHandler()
    prev = {"01_Policy_Info": [{"Test ID": "TS-001", "Effective Date": "07/18/2026"}]}
    rows = [
        {"Test ID": "TS-001", "Any Losses in the Past 3 Years?": "Yes",
         "Any Unrepaired Damage from Prior Losses?": "No", "Loss Year": "2026"},
        {"Test ID": "TS-001", "Any Losses in the Past 3 Years?": "Yes",
         "Any Unrepaired Damage from Prior Losses?": "Yes", "Loss Year": "2030"},
    ]
    out = handler.post_process([dict(r) for r in rows], "07_Cargo_LossHistory", "", prev)
    # Loss Year strictly in 2023..2025 (eff year 2026, past 3 years, not future).
    for r in out:
        assert 2023 <= int(r["Loss Year"]) <= 2025
    # Unrepaired-damage answer is one value for the insured, not both.
    assert len({r["Any Unrepaired Damage from Prior Losses?"] for r in out}) == 1


def test_cargo_no_loss_policy_single_blank_row():
    handler = CargoHandler()
    prev = {"01_Policy_Info": [{"Test ID": "TS-001", "Effective Date": "07/18/2026"}]}
    rows = [{"Test ID": "TS-001", "Any Losses in the Past 3 Years?": "No",
             "Any Unrepaired Damage from Prior Losses?": "Yes", "Loss Year": "2024"}]
    out = handler.post_process(rows, "07_Cargo_LossHistory", "", prev)
    assert len(out) == 1
    assert out[0]["Loss Year"] == ""
    assert out[0]["Any Unrepaired Damage from Prior Losses?"] == ""


# ---------------------------------------------------------------------------
# APD-017/020/022/024 — every insured on the Policy roster must get a child
# schedule, not just the first few the LLM emitted ("for 20 sets, only 07").
# ---------------------------------------------------------------------------

def test_apd_child_schedules_cover_full_roster():
    handler = ApdHandler()
    # 20 policies on the roster …
    roster = [f"TS-{i:03d}" for i in range(1, 21)]
    policy_rows = [{"Test ID": t, "Binding State": "VA"} for t in roster]
    prev = {"01_Policy_Info": policy_rows}

    # … but the LLM only emitted vehicle rows for the first 3 Test IDs.
    vehicles = [{"Test ID": f"TS-{i:03d}", "VIN Number": f"VIN{i}", "#": 1}
                for i in range(1, 4)]
    out = handler.post_process(vehicles, "05_APD_Vehicles", "",
                               previous_sheets_data=prev)
    covered = {r["Test ID"] for r in out}
    assert covered == set(roster)               # all 20 insureds have vehicles
    # Each insured has a multi-row (>=min) schedule, capped at the 20 unit max.
    from collections import Counter
    counts = Counter(r["Test ID"] for r in out)
    assert all(3 <= c <= 20 for c in counts.values())

    # Loss payees / loss history also cover the full roster.
    lp = [{"Test ID": "TS-001", "Have Loss Payees": "Yes",
           "Loss Payee Name": "Bank A", "#": 1}]
    lp_out = handler.post_process(lp, "09_APD_LossPayees", "",
                                  previous_sheets_data=prev)
    assert {r["Test ID"] for r in lp_out} == set(roster)


# ---------------------------------------------------------------------------
# CARGO-001/002/003 + APD-001/002/003 — phone/fax format & state variety are
# enforced by the universal SPG variety pass that main.py runs for these LOBs.
# ---------------------------------------------------------------------------

def test_universal_pass_fixes_cargo_phone_fax_states():
    rows = [
        {"Test ID": "TS-001", "Rating State": "VA", "State": "VA", "Agency State": "VA",
         "Agency Phone Number": "8045551234", "Agency Fax Number": "8045559999"},
        {"Test ID": "TS-002", "Rating State": "TX", "State": "TX", "Agency State": "TX",
         "Agency Phone Number": "2145550000", "Agency Fax Number": "2145551111"},
    ]
    enforce_variety_fields(rows)
    for r in rows:
        assert r["Agency Phone Number"].startswith("(")     # CARGO-001 / APD-001
        assert r["Agency Fax Number"].startswith("(")       # CARGO-002 / APD-002
        for c in ("State", "Agency State"):                 # CARGO-003 / APD-003
            assert is_us_state(r[c])
            assert r[c] != r["Rating State"]                # not collapsed onto rating
