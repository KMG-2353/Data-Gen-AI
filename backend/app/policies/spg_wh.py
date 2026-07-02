"""SPG Wind / Hail (WH) Excel Rater handler.

The WH rater is a compact two-schedule template: a ``01_Policy_Info`` sheet and a
``02_WH_Locations`` schedule (plus a deterministic ``Test Scenario Details``
summary). Every sheet joins on ``Test ID``.

The logged defects are the same dropdown-collapse / format family the rest of the
SPG raters hit:

  * WH-001 — Agent phone / fax not in U.S. format.
  * WH-002 — state fields (Policy Info address states + the Locations ``St``
    column) collapsed onto the Binding/Rating state instead of spanning all
    states.

Phone/fax + address-state variety are handled uniformly by
``app/rulebook/variety.py`` once ``main.py`` runs it for the WH LOB (the WH
Locations state column is the bare abbreviation ``St``, which the variety pass
now recognises). This handler adds the template-specific pieces: Test IDs, Type
of Entity variety, Expiration = Effective + 1 year, Quote Date <= Effective Date,
deterministic multi-row Locations, and the Test Scenario Details summary.
"""
from __future__ import annotations

import random
from datetime import timedelta
from typing import Any

from app.rulebook.primitives import (
    add_one_year as _add_one_year,
    column_seed as _col_seed,
    default_test_id as _default_tid,
    find_col as _find_col,
    format_date_slash as _fmt_date,
    normalize_sheet_name as _norm_sheet,
    parse_date as _parse_date,
    pin_quote_effective_expiration as _pin_policy_dates,
    spread_pick as _spread_pick,
    tid_value as _tid,
)
from app.rulebook.variety import ensure_child_row_multiplicity

# Type of Entity dropdown — verbatim from the WH 15_LKP_Dropdowns "Type of
# Entity" column (no Joint Venture, unlike the commercial-auto raters). Single
# source. The LLM collapses onto Corporation/LLC, so the engine fans the value.
_WH_ENTITY_TYPES = (
    "Individual", "Corporation", "LLC", "Partnership", "Trust", "Estate",
)

# A single insured may carry up to 20 locations (ruleset). The per-insured
# minimum guarantees a multi-location sample rather than the single location the
# LLM emits; the maximum caps over-production.
_LOCATIONS_PER_INSURED = 4
_MAX_LOCATIONS = 20


def _policy_rows(previous: dict | None) -> list[dict]:
    if not previous:
        return []
    for name, rows in previous.items():
        if "policy info" in _norm_sheet(name):
            return rows or []
    return []


class WhHandler:
    policy_type = "WH"

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = _norm_sheet(sheet_name)
        if "policy info" in sn or "policy information" in sn:
            return "policy"
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        if "location" in sn:
            return "locations"
        if "loss" in sn:
            return "loss_history"
        return "unknown"

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        st = self.detect_sheet_type(sheet_name)
        if st == "scenario_details":
            return 0, ""
        if st == "policy":
            return original_row_count, self._policy_rules()
        if st == "locations":
            n = len(policy_data) if policy_data else 0
            count = max(original_row_count, n * _LOCATIONS_PER_INSURED) if n else original_row_count
            return count, self._locations_rules()
        return original_row_count, ""

    def _policy_rules(self) -> str:
        return f"""
SPG WIND/HAIL — POLICY INFO RULES (HARD CONSTRAINTS):
- Test ID: sequential TS-001, TS-002, TS-003 … (zero-padded). Reuse the SAME id across every sheet.
- Binding State (auto from Location 1) ONLY is the restricted state; the Agent Address State, Insured Mailing State and every other physical-address state may be ANY valid US state, VARIED and consistent with their own City/ZIP — do NOT collapse them onto the binding state. [WH-002]
- Type of Entity: VARY across records so all of {', '.join(_WH_ENTITY_TYPES)} appear — not just Corporation/LLC.
- Effective Date / Expiration Date: MM/DD/YYYY. Expiration = Effective + 1 year.
- Quote Date: MM/DD/YYYY; on or before the Effective Date.
- Agent Phone / Agent Fax / Phone Number: valid U.S. format, e.g. (123) 456-7890. [WH-001]
"""

    def _locations_rules(self) -> str:
        return """
SPG WIND/HAIL — LOCATIONS RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Info.
- #: sequential per scenario starting at 1. Generate MULTIPLE locations per insured (a realistic mix, up to 20) — never a single location.
- St: the 2-letter state, ANY valid US state varied across rows and consistent with City/ZIP — NOT limited to the binding state. [WH-002]
"""

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]] | None:
        if self.detect_sheet_type(sheet_name) == "scenario_details":
            return self._build_scenario_details(unique_headers, previous_sheets_data or {})
        return None

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        st = self.detect_sheet_type(sheet_name)
        if st == "policy":
            self._fix_policy(rows)
            return rows
        if st == "locations":
            # Anchor to the Policy roster so every insured gets locations and no
            # orphan (LLM-invented) Test ID leaks into the schedule.
            roster = [_tid(r) for r in _policy_rows(previous_sheets_data) if _tid(r)]
            return ensure_child_row_multiplicity(
                rows, min_per_tid=_LOCATIONS_PER_INSURED, max_per_tid=_MAX_LOCATIONS,
                all_test_ids=roster or None,
            )
        return rows

    def _fix_policy(self, rows: list[dict[str, Any]]) -> None:
        if not rows:
            return
        entity_key = _find_col(rows[0], "type of entity") or _find_col(rows[0], "org type")
        ent_seed = _col_seed(entity_key or "type of entity")
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower().strip() == "test id"), None)
            if tid_key is not None:
                row[tid_key] = _default_tid(idx + 1)

            if entity_key:
                row[entity_key] = _spread_pick(idx, _WH_ENTITY_TYPES, seed=ent_seed)

            # WH-005: Quote Date pinned to today (data-creation date), Effective
            # clamped to >= today, Expiration = Effective + 1 year.
            _pin_policy_dates(row)

    def _build_scenario_details(
        self,
        unique_headers: list[str],
        previous_sheets_data: dict[str, list[dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        pi_rows = _policy_rows(previous_sheets_data)
        if not pi_rows:
            return []

        loc_rows: list[dict] = []
        for name, data in previous_sheets_data.items():
            if "location" in _norm_sheet(name):
                loc_rows = data or []
                break
        loc_count: dict[str, int] = {}
        for r in loc_rows:
            t = _tid(r)
            if t:
                loc_count[t] = loc_count.get(t, 0) + 1

        def _hdr(*frags: str) -> str | None:
            for h in unique_headers:
                hl = h.lower()
                if all(f in hl for f in frags):
                    return h
            return None

        sid_key = _hdr("scenario")
        state_key = _hdr("state")
        entity_key = _hdr("type of entity") or _hdr("entity")
        locc_key = _hdr("location", "count")

        pi_state_key = (_find_col(pi_rows[0], "binding state")
                        or _find_col(pi_rows[0], "rating state")
                        or _find_col(pi_rows[0], "state"))
        pi_entity_key = _find_col(pi_rows[0], "type of entity")

        out: list[dict[str, Any]] = []
        for pi in pi_rows:
            tid = _tid(pi)
            row: dict[str, Any] = {}
            if sid_key:
                row[sid_key] = tid
            if state_key:
                row[state_key] = pi.get(pi_state_key, "") if pi_state_key else ""
            if entity_key:
                row[entity_key] = pi.get(pi_entity_key, "") if pi_entity_key else ""
            if locc_key:
                row[locc_key] = loc_count.get(tid, 0)
            out.append(row)
        return out
