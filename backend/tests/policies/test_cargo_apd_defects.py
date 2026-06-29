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
