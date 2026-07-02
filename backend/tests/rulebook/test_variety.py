"""Unit tests for the universal dropdown-variety pass (app/rulebook/variety.py).

Covers the determinism-by-default contract behind the cross-LOB QA defect class
"agent collapses dropdown values onto the binding state / leaves phones
unformatted": address states must fan across the full US pool (never the binding
state), phones/faxes must be U.S.-formatted, and the pass must be idempotent and
preserve intentional blanks + the binding column.
"""
import copy
import random

import pytest

from app.rulebook.primitives import US_STATES, is_us_state
from app.rulebook.variety import (
    enforce_variety_fields,
    format_phone_fax,
    spread_address_states,
)


@pytest.fixture(autouse=True)
def _seed():
    random.seed(1234)


def _rows():
    # The exact symptom QA logs: every address state == the binding state.
    return [
        {"Test ID": "TS-001", "Binding State": "VA", "Address - State": "VA",
         "Garaging State": "VA", "Mailing State": "VA", "Agent Phone": "8045551234",
         "Agent Fax": "18045559999"},
        {"Test ID": "TS-002", "Binding State": "MD", "Address - State": "MD",
         "Garaging State": "MD", "Mailing State": "", "Agent Phone": "(301) 555-0000",
         "Agent Fax": "abc"},
        {"Test ID": "TS-003", "Binding State": "PA", "Address - State": "PA",
         "Garaging State": "PA", "Mailing State": "PA", "Agent Phone": "2155557777",
         "Agent Fax": ""},
    ]


def test_address_states_leave_binding_subset():
    rows = _rows()
    enforce_variety_fields(rows)
    for r in rows:
        for col in ("Address - State", "Garaging State"):
            assert is_us_state(r[col])
            assert r[col] != r["Binding State"]  # no longer collapsed onto binding


def test_binding_state_is_never_touched():
    rows = _rows()
    enforce_variety_fields(rows)
    assert [r["Binding State"] for r in rows] == ["VA", "MD", "PA"]


def test_blank_state_cells_preserved():
    # Mailing State blanked by a handler dependency rule must stay blank.
    rows = _rows()
    enforce_variety_fields(rows)
    assert rows[1]["Mailing State"] == ""


def test_sibling_address_columns_show_combinations():
    rows = _rows()
    enforce_variety_fields(rows)
    # A single row should not repeat one state across all its address fields.
    assert rows[0]["Address - State"] != rows[0]["Garaging State"]


def test_phone_and_fax_formatting():
    rows = _rows()
    format_phone_fax(rows)
    assert rows[0]["Agent Phone"] == "(804) 555-1234"
    assert rows[0]["Agent Fax"] == "(804) 555-9999"  # 11-digit leading 1 dropped
    assert rows[2]["Agent Fax"] == ""                 # blank preserved


def test_idempotent():
    rows = _rows()
    enforce_variety_fields(rows)
    snapshot = copy.deepcopy(rows)
    enforce_variety_fields(rows)
    assert rows == snapshot


def test_state_selection_overrides_pool():
    # Frontend state filter wins: address states stay inside the selection.
    rows = _rows()
    spread_address_states(rows, state_selection=["TX", "FL"])
    for r in rows:
        for col in ("Address - State", "Garaging State", "Mailing State"):
            if str(r[col]).strip():
                assert r[col] in ("TX", "FL")


def test_full_pool_is_exercised_over_enough_rows():
    rows = [{"State": "VA"} for _ in range(len(US_STATES))]
    spread_address_states(rows)
    seen = {r["State"] for r in rows}
    # With one column and N>=len rows the spread covers the whole pool.
    assert seen == set(US_STATES)


# ---------------------------------------------------------------------------
# City/State/ZIP correspondence (Class B):
# DEF-024 / HO-023 / CARGO-007 / WH-004 / IM generic / APD-015
# ---------------------------------------------------------------------------
from app.rulebook.geo import STATE_GEO

_ZIP_TO_STATE = {z: st for st, pairs in STATE_GEO.items() for _c, z in pairs}
_CITYZIP_TO_STATE = {(c, z): st for st, pairs in STATE_GEO.items() for c, z in pairs}


def _addr_rows(n=8):
    return [
        {"Address – City": "OldCity", "Address – State": "VA", "Address – Zip": "00000",
         "Mailing City": "MOld", "Mailing State": "VA", "Mailing Zip": "99999",
         "Binding State": "VA"}
        for _ in range(n)
    ]


def test_zip_matches_assigned_state():
    rows = _addr_rows()
    spread_address_states(rows)
    for r in rows:
        assert _ZIP_TO_STATE[r["Address – Zip"]] == r["Address – State"]
        assert _ZIP_TO_STATE[r["Mailing Zip"]] == r["Mailing State"]


def test_city_state_zip_triple_is_real():
    rows = _addr_rows()
    spread_address_states(rows)
    for r in rows:
        assert (r["Address – City"], r["Address – Zip"]) in _CITYZIP_TO_STATE
        assert _CITYZIP_TO_STATE[(r["Address – City"], r["Address – Zip"])] == r["Address – State"]


def test_state_selection_keeps_zip_consistent():
    rows = _addr_rows()
    spread_address_states(rows, state_selection=["TX", "FL"])
    for r in rows:
        assert r["Address – State"] in {"TX", "FL"}
        assert _ZIP_TO_STATE[r["Address – Zip"]] == r["Address – State"]


def test_address_consistency_is_idempotent():
    rows = _addr_rows()
    spread_address_states(rows)
    snapshot = copy.deepcopy(rows)
    spread_address_states(rows)
    assert rows == snapshot


def test_blank_block_city_zip_preserved():
    # A dependency-blanked mailing block (State blank) leaves City/Zip untouched.
    rows = [{"Address – City": "C", "Address – State": "VA", "Address – Zip": "00000",
             "Mailing City": "", "Mailing State": "", "Mailing Zip": "",
             "Binding State": "VA"}]
    spread_address_states(rows)
    assert rows[0]["Mailing City"] == "" and rows[0]["Mailing Zip"] == ""


def test_every_state_has_geo():
    # Every state in the spread pool must have a consistent geo triple available.
    for st in US_STATES:
        assert STATE_GEO.get(st), f"missing geo for {st}"


# ---------------------------------------------------------------------------
# HO-023 / WH-004 — City/State/ZIP correspondence must hold on EVERY LOB.
# ---------------------------------------------------------------------------

from app.rulebook.variety import _addr_kind, enforce_address_consistency


def test_stated_amount_not_misclassified_as_state():
    # "Stated Amount or ACV" (APD/Cargo) contains the letters "state" but is a
    # value column — it must never be treated as an address State and clobbered.
    assert _addr_kind("Stated Amount or ACV") is None
    assert _addr_kind("Real Estate Value") is None
    assert _addr_kind("Interstate Miles") is None
    # genuine address columns still classify
    assert _addr_kind("Agency Address State") == "state"
    assert _addr_kind("Insured Address City") == "city"
    assert _addr_kind("Mailing ZIP") == "zip"


def test_spread_preserves_stated_amount_value():
    rows = [{"Test ID": "TS-01", "Stated Amount or ACV": "50000",
             "Address State": "NY", "Address City": "x", "Address ZIP": "0"}]
    spread_address_states(rows)
    assert rows[0]["Stated Amount or ACV"] == "50000"      # not overwritten
    assert _ZIP_TO_STATE[rows[0]["Address ZIP"]] == rows[0]["Address State"]


def test_consistency_repairs_mismatch_keeping_state():
    # The exact HO-023 / WH-004 symptom: valid ZIP but for a different state.
    rows = [{"Test ID": "TS-01", "Insured Address State": "TX",
             "Insured Address City": "Miami", "Insured Address ZIP": "33130"}]
    enforce_address_consistency(rows)
    r = rows[0]
    assert r["Insured Address State"] == "TX"                       # state unchanged
    assert (r["Insured Address City"], r["Insured Address ZIP"]) in _CITYZIP_TO_STATE
    assert _CITYZIP_TO_STATE[(r["Insured Address City"], r["Insured Address ZIP"])] == "TX"


def test_consistency_is_idempotent_and_preserves_valid_triple():
    rows = [{"State": "CA", "City": "San Diego", "ZIP": "92101"}]
    enforce_address_consistency(rows)
    assert rows[0] == {"State": "CA", "City": "San Diego", "ZIP": "92101"}


def test_consistency_preserves_blank_state_block():
    rows = [{"Mailing State": "", "Mailing City": "keep", "Mailing ZIP": "keepz"}]
    enforce_address_consistency(rows)
    assert rows[0]["Mailing City"] == "keep" and rows[0]["Mailing ZIP"] == "keepz"


def test_consistency_covers_generic_multi_block():
    # A generic uploaded template with two independent address blocks, both wrong.
    rows = [{
        "State": "FL", "City": "Denver", "ZIP": "80202",           # CO city/zip under FL
        "Mailing State": "WA", "Mailing City": "Austin", "Mailing ZIP": "78701",  # TX under WA
    }]
    enforce_address_consistency(rows)
    r = rows[0]
    assert _CITYZIP_TO_STATE[(r["City"], r["ZIP"])] == "FL"
    assert _CITYZIP_TO_STATE[(r["Mailing City"], r["Mailing ZIP"])] == "WA"
