"""SPG commercial-auto Excel Rater handlers — Motor Truck Cargo (CARGO) and
Auto Physical Damage (APD).

Both raters share the same shape: a flat ``Policy Information`` sheet, an
additional-insured / UW sheet, driver + vehicle (power unit) schedules, a
trailer schedule (APD), a commodities schedule, coverages, loss history, loss
payees, additional info, and a ``Test Scenario Details`` summary — every sheet
joining on ``Test ID``.

The numbered blank templates (``01_Policy_Info``, ``03_APD_Drivers``,
``09_APD_LossPayees`` …) renamed the sheets, so detection is done on the
normalised sheet name. The logged defects are dropdown-variety / multiplicity /
format failures:

  * CARGO-001/002, APD-001/002 — Agent phone/fax not in U.S. format.
  * CARGO-003, APD-003 — address state fields collapsed onto the binding/rating
    state instead of spanning all states.
  * CARGO-004 — Radius of Operations populated with arbitrary values.
  * CARGO-005, APD-004 — Type of Entity collapsed onto Corporation/LLC.
  * APD-005 — Type of Company / Carrier set to an out-of-dropdown value.
  * APD-006/007/008/009 — Drivers / Vehicles / Trailers / Commodities generated
    one row per insured instead of the multiple the schedule allows (up to 20).
  * APD-010 / APD-011 — Loss Payees / Loss History generated a single row.
  * APD-012 — the Test Scenario Details summary sheet was missing.

Phone/fax + address-state variety are handled uniformly for the whole SPG family
by ``app/rulebook/variety.py`` (which ``main.py`` runs after this handler). This
handler adds the template-specific pieces: Test IDs, Type of Entity + Type of
Company variety, the Cargo Radius dropdown, deterministic child-row multiplicity,
and the deterministic Test Scenario Details summary.
"""
from __future__ import annotations

import random
from datetime import timedelta
from typing import Any

from app.rulebook.primitives import (
    coerce_to_allowed as _coerce,
    column_seed as _col_seed,
    default_test_id as _default_tid,
    find_col as _find_col,
    format_date_slash as _fmt_date,
    is_no as _is_no,
    normalize_sheet_name as _norm_sheet,
    parse_date as _parse_date,
    spread_pick as _spread_pick,
    tid_value as _tid,
)
from app.rulebook.variety import ensure_child_row_multiplicity

# Radius of Operations dropdown — sourced verbatim from the SPG Cargo rater's
# Commodities data validation. CARGO-004: arbitrary values are snapped back onto
# this set (and varied across rows when invalid). Single source.
_CARGO_RADIUS = ("0-100 Miles", "100-500 Miles", "Over 500 Miles")

# Type of Entity dropdown — verbatim from the APD/Cargo 13/14_LKP_Dropdowns
# "Type of Entity" column. CARGO-005 / APD-004: the LLM only emits Corporation/
# LLC, so the engine fans the value across the full set. Single source.
_AUTO_ENTITY_TYPES = (
    "Individual", "Corporation", "LLC", "Partnership",
    "Joint Venture", "Trust", "Estate",
)

# Type of Company / Carrier dropdown — verbatim from the same LKP sheet.
# APD-005: out-of-list values (e.g. "tow trucks") are snapped onto this set.
_AUTO_COMPANY_TYPES = (
    "Common Carriers", "Private Carrier", "Contract Carriers",
    "Owner of Cargo", "Other",
)

# Child-schedule multiplicity: ruleset upper bound is 20 units/insured; loss
# payees and loss history are 1-10. The per-insured minimums guarantee a
# multi-row sample (APD-006..011) rather than the single row the LLM emits.
_UNIT_MAX = 20
_DRIVERS_PER_INSURED = 3
_VEHICLES_PER_INSURED = 3
_TRAILERS_PER_INSURED = 2
_COMMODITIES_PER_INSURED = 3
_LOSS_PAYEES_PER_INSURED = 2
_MAX_LOSS_PAYEES = 10
_LOSS_HISTORY_PER_INSURED = 2
_MAX_LOSS_HISTORY = 10


def _policy_rows(previous: dict | None) -> list[dict]:
    if not previous:
        return []
    for name, rows in previous.items():
        if "policy info" in _norm_sheet(name):
            return rows or []
    return []


class _SpgAutoHandler:
    """Shared CARGO/APD behaviour. Subclasses set ``policy_type`` and may add a
    ``_coerce_dropdowns`` hook for template-specific dropdown snapping."""

    policy_type = "SPG_AUTO"

    # ------------------------------------------------------------------
    # Sheet type detection (normalised: matches legacy + numbered names)
    # ------------------------------------------------------------------

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = _norm_sheet(sheet_name)
        if "policy info" in sn or "policy information" in sn:
            return "policy"
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        if "additional insured" in sn:
            return "uw"
        if "additinfo" in sn or "additional info" in sn:
            return "addit_info"
        if "driver" in sn:
            return "drivers"
        if "trailer" in sn:
            return "trailers"
        if "commodit" in sn:
            return "commodities"
        if "vehicle" in sn or "power unit" in sn:
            return "vehicles"
        if "coverage" in sn:
            return "coverages"
        if "loss history" in sn or "losshistory" in sn:
            return "loss_history"
        if "loss payee" in sn or "losspayee" in sn:
            return "loss_payees"
        return "unknown"

    # ------------------------------------------------------------------
    # Sheet context: row count + per-sheet hard-constraint prompt
    # ------------------------------------------------------------------

    _CHILD_TARGETS = {
        "drivers": _DRIVERS_PER_INSURED,
        "vehicles": _VEHICLES_PER_INSURED,
        "trailers": _TRAILERS_PER_INSURED,
        "commodities": _COMMODITIES_PER_INSURED,
        "loss_payees": _LOSS_PAYEES_PER_INSURED,
        "loss_history": _LOSS_HISTORY_PER_INSURED,
    }

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
        if st == "commodities":
            return self._child_count(original_row_count, policy_data,
                                     _COMMODITIES_PER_INSURED), self._commodities_rules()
        per = self._CHILD_TARGETS.get(st)
        if per:
            return self._child_count(original_row_count, policy_data, per), ""
        return original_row_count, ""

    @staticmethod
    def _child_count(original: int, policy_data: list[dict] | None, per_insured: int) -> int:
        """Scale a child sheet's requested row count to ``insureds * per_insured``
        so the LLM is asked for a multi-row sample; the deterministic expansion in
        post-processing guarantees it regardless."""
        n = len(policy_data) if policy_data else 0
        return max(original, n * per_insured) if n else original

    def _policy_rules(self) -> str:
        return f"""
SPG COMMERCIAL AUTO — POLICY INFORMATION RULES (HARD CONSTRAINTS):
- Test ID: sequential TS-001, TS-002, TS-003 … (zero-padded). Reuse the SAME id across every sheet for the same scenario.
- Binding State / Rating State ONLY: restricted to the binding dropdown (VA, MD, DC, PA, NC, WV, DE, CA, TX, GA, NV, SC, OH, AZ).
- Insured Address State / Agency State / Garaging State and every other physical-address state: ANY valid US state, VARIED across records and consistent with its own City/ZIP — do NOT force them to equal the binding state.
- Type of Entity: VARY across records so all of {', '.join(_AUTO_ENTITY_TYPES)} appear — not just Corporation/LLC. [CARGO-005 / APD-004]
- Type of Company / Carrier: one of {', '.join(_AUTO_COMPANY_TYPES)} — never an out-of-list value. [APD-005]
- Effective Date / Quote Date: MM/DD/YYYY; Quote Date MUST be on or before the Effective Date.
- Agency Phone Number / Agency Fax Number: valid U.S. format, e.g. (123) 456-7890.
"""

    def _commodities_rules(self) -> str:
        return ""

    # ------------------------------------------------------------------
    # Deterministic pre-generation: Test Scenario Details summary
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Post-processing
    # ------------------------------------------------------------------

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
            self._coerce_dropdowns(rows)
            return rows

        self._coerce_dropdowns(rows)

        if st in ("drivers", "vehicles", "trailers", "commodities"):
            unique = {
                "drivers": ("license number",),
                "vehicles": ("vin number",),
                "trailers": ("vin number",),
                "commodities": (),
            }[st]
            return ensure_child_row_multiplicity(
                rows, min_per_tid=self._CHILD_TARGETS[st], max_per_tid=_UNIT_MAX,
                unique_frags=unique,
            )
        if st == "loss_payees":
            return self._fix_loss_payees(rows)
        if st == "loss_history":
            return self._fix_loss_history(rows)
        return rows

    # ------------------------------------------------------------------
    # Policy Info fixers
    # ------------------------------------------------------------------

    def _fix_policy(self, rows: list[dict[str, Any]]) -> None:
        """Stamp Test IDs; fan Type of Entity across the full dropdown
        (CARGO-005 / APD-004); snap Type of Company / Carrier onto its dropdown
        (APD-005); clamp Quote Date <= Effective Date."""
        if not rows:
            return
        entity_cols = [
            k for k in rows[0]
            if "type of entity" in k.lower() or "org type" in k.lower()
        ]
        company_key = _find_col(rows[0], "type of company")
        ent_seeds = {k: _col_seed(k) for k in entity_cols}
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower().strip() == "test id"), None)
            if tid_key is not None:
                row[tid_key] = _default_tid(idx + 1)

            for k in entity_cols:
                row[k] = _spread_pick(idx, _AUTO_ENTITY_TYPES, seed=ent_seeds[k])

            if company_key:
                row[company_key] = _coerce(
                    row.get(company_key), _AUTO_COMPANY_TYPES,
                    _spread_pick(idx, _AUTO_COMPANY_TYPES, seed=_col_seed(company_key)),
                    fill_blank=True,
                )

            eff_key = _find_col(row, "effective date")
            quote_key = _find_col(row, "quote date") or _find_col(row, "date of quote")
            if eff_key and quote_key:
                eff = _parse_date(row.get(eff_key))
                quote = _parse_date(row.get(quote_key))
                if eff and (quote is None or quote > eff):
                    row[quote_key] = _fmt_date(eff - timedelta(days=random.randint(1, 14)))

    # ------------------------------------------------------------------
    # Child-schedule fixers
    # ------------------------------------------------------------------

    def _fix_loss_payees(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """APD-010: blank loss-payee detail when the policy has none, then
        guarantee a 1–10 multi-row schedule per policy."""
        for row in rows:
            flag_key = _find_col(row, "have loss payees") or _find_col(row, "has loss payees")
            if flag_key and _is_no(row.get(flag_key)):
                for frag in ("loss payee name", "address street", "city", "state",
                             "zip", "assigned units"):
                    k = _find_col(row, frag)
                    if k:
                        row[k] = ""
        return ensure_child_row_multiplicity(
            rows, min_per_tid=_LOSS_PAYEES_PER_INSURED, max_per_tid=_MAX_LOSS_PAYEES,
            skip_predicate=self._no_loss_payees,
        )

    @staticmethod
    def _no_loss_payees(row: dict[str, Any]) -> bool:
        k = _find_col(row, "have loss payees") or _find_col(row, "has loss payees")
        return bool(k) and _is_no(row.get(k))

    def _fix_loss_history(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """APD-011: blank loss detail for no-loss policies, then guarantee a
        multi-row loss schedule (1–10) for policies that DID have losses."""
        for row in rows:
            flag_key = _find_col(row, "losses in the past") or _find_col(row, "any losses")
            if flag_key and _is_no(row.get(flag_key)):
                for frag in ("loss year", "type of loss", "loss description",
                             "amount paid", "amount outstanding", "premium at time"):
                    k = _find_col(row, frag)
                    if k:
                        row[k] = ""
        return ensure_child_row_multiplicity(
            rows, min_per_tid=_LOSS_HISTORY_PER_INSURED, max_per_tid=_MAX_LOSS_HISTORY,
            skip_predicate=self._no_losses,
        )

    @staticmethod
    def _no_losses(row: dict[str, Any]) -> bool:
        k = _find_col(row, "losses in the past") or _find_col(row, "any losses")
        return bool(k) and _is_no(row.get(k))

    # Overridable per-LOB dropdown snapping (CARGO snaps Radius).
    def _coerce_dropdowns(self, rows: list[dict[str, Any]]) -> None:
        return None

    # ------------------------------------------------------------------
    # Test Scenario Details (deterministic per-insured summary) — APD-012
    # ------------------------------------------------------------------

    def _build_scenario_details(
        self,
        unique_headers: list[str],
        previous_sheets_data: dict[str, list[dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        pi_rows = _policy_rows(previous_sheets_data)
        if not pi_rows:
            return []

        def _sheet(*frags: str) -> list[dict]:
            for name, data in previous_sheets_data.items():
                nl = _norm_sheet(name)
                if all(f in nl for f in frags):
                    return data or []
            return []

        def _count(rows: list[dict]) -> dict[str, int]:
            counts: dict[str, int] = {}
            for r in rows:
                t = _tid(r)
                if t:
                    counts[t] = counts.get(t, 0) + 1
            return counts

        driver_c = _count(_sheet("driver"))
        vehicle_c = _count(_sheet("vehicle"))
        trailer_c = _count(_sheet("trailer"))
        commodity_c = _count(_sheet("commodit"))
        loss_c = _count(_sheet("loss history") or _sheet("losshistory"))
        lp_c = _count(_sheet("loss payee") or _sheet("losspayee"))

        def _hdr(*frags: str) -> str | None:
            for h in unique_headers:
                hl = h.lower()
                if all(f in hl for f in frags):
                    return h
            return None

        sid_key = _hdr("scenario")
        state_key = _hdr("state")
        entity_key = _hdr("type of entity") or _hdr("entity")
        cols = {
            _hdr("driver", "count"): driver_c,
            _hdr("vehicle", "count"): vehicle_c,
            _hdr("trailer", "count"): trailer_c,
            _hdr("commodit", "count"): commodity_c,
            _hdr("loss history", "count") or _hdr("loss", "count"): loss_c,
            _hdr("loss payee", "count") or _hdr("payee", "count"): lp_c,
        }

        pi_state_key = (_find_col(pi_rows[0], "binding state")
                        or _find_col(pi_rows[0], "rating state")
                        or _find_col(pi_rows[0], "state"))
        pi_entity_key = _find_col(pi_rows[0], "type of entity") or _find_col(pi_rows[0], "org type")

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
            for key, counts in cols.items():
                if key:
                    row[key] = counts.get(tid, 0)
            out.append(row)
        return out


class CargoHandler(_SpgAutoHandler):
    policy_type = "CARGO"

    def _commodities_rules(self) -> str:
        return f"""
SPG CARGO — COMMODITIES RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Information.
- Radius / Radius of Operations: MUST be one of {', '.join(_CARGO_RADIUS)} — never an arbitrary distance. VARY across rows. [CARGO-004]
"""

    def _coerce_dropdowns(self, rows: list[dict[str, Any]]) -> None:
        """CARGO-004: snap any Radius column onto its dropdown, varying invalid
        cells across the allowed set so the field is both valid and exercised."""
        if not rows:
            return
        radius_key = _find_col(rows[0], "radius")
        if not radius_key:
            return
        seed = _col_seed(radius_key)
        for i, row in enumerate(rows):
            snapped = _coerce(row.get(radius_key), _CARGO_RADIUS, None, fill_blank=True)
            if snapped is None:
                snapped = _spread_pick(i, _CARGO_RADIUS, seed=seed)
            row[radius_key] = snapped


class ApdHandler(_SpgAutoHandler):
    policy_type = "APD"
