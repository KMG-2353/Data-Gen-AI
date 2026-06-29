"""SPG commercial-auto Excel Rater handlers — Motor Truck Cargo (CARGO) and
Auto Physical Damage (APD).

Both raters share the same shape: a flat ``Policy Information`` sheet, an
additional-insured / UW sheet, driver + vehicle (power unit) schedules, a
commodities schedule, coverages, loss history, loss payees, additional interests,
and a ``Test Scenario Details`` summary — every sheet joining on ``Test ID``.

Before this module both templates fell through to the GenericHandler, so none of
their dropdown fields were enforced. The logged defects are all dropdown-variety /
format failures:

  * CARGO-001/002, APD-001/002 — Agent phone/fax not in U.S. format.
  * CARGO-003, APD-003 — address state fields collapsed onto the binding/rating
    state instead of spanning all states.
  * CARGO-004 — Radius of Operations populated with arbitrary values instead of
    its dropdown.

Phone/fax + address-state variety are handled uniformly for the whole SPG family
by the universal pass in ``app/rulebook/variety.py`` (which ``main.py`` runs after
this handler now that these templates route here). This handler adds the two
template-specific pieces: stamping the default ``TS-###`` Test IDs and snapping
the Cargo Radius dropdown.
"""
from __future__ import annotations

from typing import Any

from app.rulebook.primitives import (
    coerce_to_allowed as _coerce,
    default_test_id as _default_tid,
    find_col as _find_col,
    spread_pick as _spread_pick,
    column_seed as _col_seed,
)

# Radius of Operations dropdown — sourced verbatim from the SPG Cargo rater's
# 05_Cargo_Commodities data validation. CARGO-004: arbitrary values are snapped
# back onto this set (and varied across rows when invalid). Single source.
_CARGO_RADIUS = ("0-100 Miles", "100-500 Miles", "Over 500 Miles")


class _SpgAutoHandler:
    """Shared CARGO/APD behaviour. Subclasses set ``policy_type`` and may add a
    ``_coerce_dropdowns`` hook for template-specific dropdown snapping."""

    policy_type = "SPG_AUTO"

    # ------------------------------------------------------------------
    # Sheet type detection
    # ------------------------------------------------------------------

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = sheet_name.lower().strip()
        if "policy info" in sn or "policy information" in sn:
            return "policy"
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        if "commodit" in sn:
            return "commodities"
        if "loss history" in sn or "losshistory" in sn:
            return "loss_history"
        if "loss payee" in sn or "losspayee" in sn:
            return "loss_payees"
        if "power unit" in sn or "vehicle" in sn or "trailer" in sn:
            return "units"
        if "driver" in sn:
            return "drivers"
        return "unknown"

    # ------------------------------------------------------------------
    # Sheet context: row count + per-sheet hard-constraint prompt
    # ------------------------------------------------------------------

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        st = self.detect_sheet_type(sheet_name)
        if st == "policy":
            return original_row_count, self._policy_rules()
        if st == "commodities":
            return original_row_count, self._commodities_rules()
        return original_row_count, ""

    def _policy_rules(self) -> str:
        return """
SPG COMMERCIAL AUTO — POLICY INFORMATION RULES (HARD CONSTRAINTS):
- Test ID: sequential TS-001, TS-002, TS-003 … (zero-padded). Reuse the SAME id across every sheet for the same scenario.
- Binding State / Rating State ONLY: restricted to the binding dropdown (VA, MD, DC, PA, NC, WV, DE, CA, TX, GA, NV, SC, OH, AZ).
- Insured Address State / Agency State / Garaging State and every other physical-address state: ANY valid US state, VARIED across records and consistent with its own City/ZIP — do NOT force them to equal the binding state.
- Agency Phone Number / Agency Fax Number / Contact Number: valid U.S. format, e.g. (123) 456-7890.
"""

    def _commodities_rules(self) -> str:
        return ""

    # ------------------------------------------------------------------
    # Deterministic pre-generation (none — scenario sheet is LLM-generated but
    # tagged so the universal variety pass skips its summary State column).
    # ------------------------------------------------------------------

    def pre_generate(self, *args, **kwargs) -> list[dict[str, Any]] | None:
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
            self._stamp_test_ids(rows)
        self._coerce_dropdowns(rows)
        return rows

    @staticmethod
    def _stamp_test_ids(rows: list[dict[str, Any]]) -> None:
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower().strip() == "test id"), None)
            if tid_key is not None:
                row[tid_key] = _default_tid(idx + 1)

    # Overridable per-LOB dropdown snapping (CARGO snaps Radius).
    def _coerce_dropdowns(self, rows: list[dict[str, Any]]) -> None:
        return None


class CargoHandler(_SpgAutoHandler):
    policy_type = "CARGO"

    def _commodities_rules(self) -> str:
        return f"""
SPG CARGO — COMMODITIES RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Information.
- Radius of Operations: MUST be one of the dropdown values {', '.join(_CARGO_RADIUS)} — never an arbitrary distance. VARY across rows. [CARGO-004]
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
            # Snap valid values to their canonical form; blank/arbitrary cells
            # (default sentinel None) are varied across the dropdown.
            snapped = _coerce(row.get(radius_key), _CARGO_RADIUS, None, fill_blank=True)
            if snapped is None:
                snapped = _spread_pick(i, _CARGO_RADIUS, seed=seed)
            row[radius_key] = snapped


class ApdHandler(_SpgAutoHandler):
    policy_type = "APD"
    # APD's logged defects (phone/fax/state) are all covered by the universal
    # variety pass; the handler exists so the template routes here (and so Test
    # IDs are stamped) rather than falling through to GenericHandler.
