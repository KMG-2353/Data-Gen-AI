"""SPG Inland Marine (IM) Excel Rater policy handler.

Enforces the hard validation rules from the SPG IM RuleSet that the LLM alone
gets wrong, deterministically, after generation:

- Date of Quote MUST be <= Effective Date (Rule 9 / DF-IM-001, DF-IM-002)
- Equipment Schedule rows only exist when the insured's "Scheduled Equipment
  Coverage" = Yes on Policy Info (Rule 46 / DF-IM-003, DF-IM-004)
- Loss Date MUST fall within the 3 years before the Effective Date and be
  earlier than it (Rule 64 / DF-IM-005)
- Equipment "Value ($)" is a numeric value $25,000–$50,000 (Rule 50 / DF-IM-006)
- "Total Value of Miscellaneous Articles ($)" is numeric $0–$10,000
  (Rule 60 / DF-IM-007)
- Loss "Amount ($)" is a positive numeric value (Rule 67 / DF-IM-008)

The handler also rebuilds the "Test Scenario Details" sheet deterministically as
a per-insured summary so the scenario-based architecture stays coherent (one
Scenario ID per Policy Info Test ID, with the real child-row counts).
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from typing import Any

from app.rulebook.primitives import (
    parse_date as _parse_date,
    format_date_slash as _fmt_date,
    to_number as _to_number,
    find_col as _find_col,
    tid_value as _tid,
    default_test_id as _default_tid,
    is_yes as _is_yes,
    is_no as _is_no,
)


def _get_policy_rows(previous: dict | None) -> list[dict]:
    """Locate the Policy Info rows in the already-generated sheets map."""
    if not previous:
        return []
    for name, rows in previous.items():
        if "policy info" in name.lower():
            return rows or []
    return []


class ImHandler:
    policy_type = "IM"

    # ------------------------------------------------------------------
    # Sheet type detection
    # ------------------------------------------------------------------

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = sheet_name.lower().strip()
        if "policy info" in sn:
            return "policy"
        if "equipment schedule" in sn:
            return "equipment"
        if "misc articles" in sn or "miscellaneous articles" in sn:
            return "misc_articles"
        if "loss history" in sn:
            return "loss_history"
        if "additional interest" in sn:
            return "additional_interests"
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        return "unknown"

    # ------------------------------------------------------------------
    # Sheet context (row count + rules injected into LLM prompt)
    # ------------------------------------------------------------------

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        sheet_type = self.detect_sheet_type(sheet_name)

        # Test Scenario Details is built deterministically in pre_generate from
        # the upstream sheets — never sent to the LLM.
        if sheet_type == "scenario_details":
            return 0, ""

        if sheet_type == "policy":
            rules = """
SPG INLAND MARINE — POLICY INFO RULES (HARD CONSTRAINTS):
- Test ID: MUST follow the format TS-001, TS-002, TS-003, ... (sequential, zero-padded to 3 digits).
- Binding State / Agency State / Mailing State / Coverage State: MUST be one of VA, MD, DC, PA, NC, WV, DE. All states within a record must be consistent. [Rule 1/19/29/34]
- Effective Date: MM/DD/YYYY; MUST be on or before Expiration Date. [Rule 4]
- Expiration Date: MM/DD/YYYY; later than Effective Date (typically Effective + Policy Term). [Rule 5]
- Date of Quote: MM/DD/YYYY; MUST be LESS THAN OR EQUAL TO the Effective Date — NEVER after it. [Rule 9]
- Type of Entity: one of Individual, Corporation, LLC, Partnership, Joint Venture, Trust, Estate. [Rule 23]
- Trustee fields: populate ONLY when Type of Entity = Trust; otherwise leave ALL Trustee columns blank. [Rule 36/43]
- Scheduled Equipment Coverage: "Yes" or "No". If "No", there must be NO Equipment Schedule rows for this Test ID. [Rule 44/46]
- Miscellaneous Articles Coverage: "Yes" or "No". If "No", the Misc Articles total must be blank. [Rule 45]
"""
            return original_row_count, rules

        if sheet_type == "equipment":
            rules = """
SPG INLAND MARINE — EQUIPMENT SCHEDULE RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Info. Generate equipment ONLY for insureds whose "Scheduled Equipment Coverage" = Yes. [Rule 46]
- Value ($): a PLAIN NUMBER greater than 25000 and not exceeding 50000 (e.g. 38500). NO "$", NO commas, NO quotes. [Rule 50]
- Serial Number: unique alphanumeric per item. [Rule 49]
- Used for Logging?: "Yes" or "No". [Rule 51]
- Loss Payee?: "Yes" or "No". If "No", leave Loss Payee Name and all LP Address fields blank. [Rule 52/53]
- LP Addr State (when Loss Payee? = Yes): one of VA, MD, DC, PA, NC, WV, DE, AL. [Rule 55]
"""
            return original_row_count, rules

        if sheet_type == "misc_articles":
            rules = """
SPG INLAND MARINE — MISC ARTICLES SCHEDULE RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Info.
- Miscellaneous Articles Coverage Selected?: "Yes" or "No". [Rule 59]
- Total Value of Miscellaneous Articles ($): when coverage = Yes, a PLAIN NUMBER greater than 0 and not exceeding 10000 (e.g. 4850). NO "$", NO commas, NO quotes. When coverage = No, leave BLANK. [Rule 60]
"""
            return original_row_count, rules

        if sheet_type == "loss_history":
            eff_hint = ""
            if policy_data:
                # Surface each insured's effective date so loss dates land in the
                # valid 3-year window the first time.
                lines = []
                for r in policy_data:
                    eff_key = _find_col(r, "effective date")
                    lines.append(f"  {_tid(r)}: Effective {r.get(eff_key, '')}")
                eff_hint = (
                    "\nPer-insured Effective Dates (loss dates must fall in the 3 years before these):\n"
                    + "\n".join(lines)
                )
            rules = f"""
SPG INLAND MARINE — LOSS HISTORY RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Info.
- Any Losses in Past 3 Years?: "Yes" or "No". If "No", leave all loss fields (#, Loss Date, Type of Loss, Details, Amount) blank for that Test ID. [Rule 61]
- Loss Date: MM/DD/YYYY; MUST be within the 3 years immediately before the insured's Effective Date AND earlier than the Effective Date. [Rule 64]
- Type of Loss: one of Fire, Water Damage - Weather Related, Water Damage - Non-weather Related, Wind Damage, Hail Damage, Theft, Physical Damage - All Other, Liability. [Rule 65]
- Amount ($): a PLAIN positive NUMBER greater than 0 (e.g. 18450). NO "$", NO commas, NO quotes. [Rule 67]
{eff_hint}
"""
            return original_row_count, rules

        if sheet_type == "additional_interests":
            rules = """
SPG INLAND MARINE — ADDITIONAL INTERESTS RULES (HARD CONSTRAINTS):
- Test ID: reuse the SAME TS-### IDs from Policy Info.
- Any Loss Payees on Scheduled Equipment?: "Yes" or "No". If "No", leave all Loss Payee fields blank. [Rule 69]
- State: one of VA, MD, DC, PA, NC, WV, DE, AL. [Rule 75]
- For Equipment #: must reference an existing Scheduled Equipment item number. [Rule 77]
- Interest Type: one of Loss Payable, Lender's Loss Payable, Contract Sale. [Rule 78]
"""
            return original_row_count, rules

        return original_row_count, ""

    # ------------------------------------------------------------------
    # Deterministic pre-generation
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
            return self._build_scenario_details(
                unique_headers, previous_sheets_data or {}
            )
        return None

    # ------------------------------------------------------------------
    # Post-processing: hard constraint enforcement
    # ------------------------------------------------------------------

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        sheet_type = self.detect_sheet_type(sheet_name)

        if sheet_type == "policy":
            return self._fix_policy_info(rows)
        if sheet_type == "equipment":
            return self._fix_equipment(rows, previous_sheets_data)
        if sheet_type == "misc_articles":
            return self._fix_misc_articles(rows)
        if sheet_type == "loss_history":
            return self._fix_loss_history(rows, previous_sheets_data)
        if sheet_type == "additional_interests":
            return self._fix_additional_interests(rows)
        return rows

    # ------------------------------------------------------------------
    # Per-sheet fixers
    # ------------------------------------------------------------------

    def _fix_policy_info(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Stamp Test IDs, clamp Date of Quote <= Effective Date (Rule 9),
        and blank Trustee fields for non-Trust entities (Rule 36/43)."""
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower().strip() == "test id"), None)
            if tid_key is not None:
                row[tid_key] = _default_tid(idx + 1)

            eff_key = _find_col(row, "effective date")
            quote_key = _find_col(row, "date of quote")
            if eff_key and quote_key:
                eff = _parse_date(row.get(eff_key))
                quote = _parse_date(row.get(quote_key))
                # DF-IM-001 / DF-IM-002: Date of Quote must be <= Effective Date.
                if eff and (quote is None or quote > eff):
                    # Place the quote a few days before the effective date.
                    row[quote_key] = _fmt_date(eff - timedelta(days=random.randint(1, 14)))

            # Rule 36/43: Trustee fields only for Type of Entity = Trust.
            entity_key = _find_col(row, "type of entity")
            if entity_key:
                is_trust = str(row.get(entity_key, "")).strip().lower() == "trust"
                if not is_trust:
                    for k in list(row.keys()):
                        if "trustee" in k.lower():
                            row[k] = ""

        return rows

    def _fix_equipment(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None,
    ) -> list[dict[str, Any]]:
        """Drop equipment for No-coverage insureds (Rule 46 / DF-IM-003/004),
        coerce Value ($) to a numeric $25k–$50k (Rule 50 / DF-IM-006), and blank
        Loss Payee fields when Loss Payee? = No (Rule 53)."""
        pi_rows = _get_policy_rows(previous_sheets_data)

        # Map Test ID -> Scheduled Equipment Coverage (Yes/No) from Policy Info.
        coverage_by_tid: dict[str, bool] = {}
        for pi in pi_rows:
            cov_key = _find_col(pi, "scheduled equipment coverage")
            if cov_key is not None:
                coverage_by_tid[_tid(pi)] = _is_yes(pi.get(cov_key))

        kept: list[dict[str, Any]] = []
        for row in rows:
            tid = _tid(row)
            # DF-IM-003 / DF-IM-004: skip equipment rows for insureds whose
            # Scheduled Equipment Coverage = No.
            if tid in coverage_by_tid and not coverage_by_tid[tid]:
                continue

            # DF-IM-006: Value ($) numeric, clamped to (25000, 50000].
            val_key = _find_col(row, "value")
            if val_key:
                num = _to_number(row.get(val_key))
                if num is None:
                    num = random.randint(25001, 50000)
                num = int(round(num))
                if num <= 25000:
                    num = random.randint(25001, 50000)
                elif num > 50000:
                    num = 50000
                row[val_key] = num

            # Rule 53-56: blank Loss Payee detail fields when Loss Payee? = No.
            lp_flag_key = _find_col(row, "loss payee?")
            if lp_flag_key and _is_no(row.get(lp_flag_key)):
                for frag in ("loss payee name", "lp addr", "interest type"):
                    k = _find_col(row, frag)
                    if k:
                        row[k] = ""

            kept.append(row)

        return kept

    def _fix_misc_articles(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Coerce Total Value to numeric $0–$10k; blank it when coverage = No
        (Rule 59/60 / DF-IM-007)."""
        for row in rows:
            sel_key = _find_col(row, "miscellaneous articles coverage")
            total_key = _find_col(row, "total value")
            if not total_key:
                continue

            if sel_key and _is_no(row.get(sel_key)):
                row[total_key] = ""
                continue

            # DF-IM-007: numeric value > 0 and <= 10000.
            num = _to_number(row.get(total_key))
            if num is None or num <= 0:
                num = random.randint(500, 10000)
            num = int(round(num))
            if num > 10000:
                num = 10000
            row[total_key] = num

        return rows

    def _fix_loss_history(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None,
    ) -> list[dict[str, Any]]:
        """Clamp Loss Date into the 3-year-before-effective window (Rule 64 /
        DF-IM-005), coerce Amount ($) to a positive number (Rule 67 /
        DF-IM-008), and blank loss fields when there were no losses (Rule 61)."""
        pi_rows = _get_policy_rows(previous_sheets_data)
        eff_by_tid: dict[str, date] = {}
        for pi in pi_rows:
            eff_key = _find_col(pi, "effective date")
            eff = _parse_date(pi.get(eff_key)) if eff_key else None
            if eff:
                eff_by_tid[_tid(pi)] = eff

        for row in rows:
            tid = _tid(row)
            eff = eff_by_tid.get(tid)

            # Rule 61: if no losses in past 3 years, blank the loss detail fields.
            losses_key = _find_col(row, "losses in past 3 years")
            if losses_key and _is_no(row.get(losses_key)):
                for frag in ("loss date", "type of loss", "details", "amount"):
                    k = _find_col(row, frag)
                    if k:
                        row[k] = ""
                num_key = next((k for k in row if k.strip() == "#"), None)
                if num_key:
                    row[num_key] = ""
                continue

            # DF-IM-005: Loss Date within [effective - 3y, effective - 1d].
            ld_key = _find_col(row, "loss date")
            if ld_key:
                ld = _parse_date(row.get(ld_key))
                if eff is not None:
                    window_start = eff - timedelta(days=3 * 365)
                    latest = eff - timedelta(days=1)
                    if ld is None or ld < window_start or ld >= eff:
                        span = (latest - window_start).days
                        row[ld_key] = _fmt_date(
                            window_start + timedelta(days=random.randint(0, max(span, 0)))
                        )

            # DF-IM-008: Amount ($) numeric > 0.
            amt_key = _find_col(row, "amount")
            if amt_key:
                num = _to_number(row.get(amt_key))
                if num is None or num <= 0:
                    num = random.randint(1000, 50000)
                row[amt_key] = int(round(num))

        return rows

    def _fix_additional_interests(
        self, rows: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """Blank Loss Payee fields when "Any Loss Payees..." = No (Rule 69)."""
        for row in rows:
            flag_key = _find_col(row, "any loss payees")
            if flag_key and _is_no(row.get(flag_key)):
                for frag in (
                    "full name", "street address", "city", "state", "zip",
                    "for equipment", "interest type", "phone", "email",
                ):
                    k = _find_col(row, frag)
                    if k:
                        row[k] = ""
        return rows

    # ------------------------------------------------------------------
    # Test Scenario Details (deterministic summary, one row per insured)
    # ------------------------------------------------------------------

    def _build_scenario_details(
        self,
        unique_headers: list[str],
        previous_sheets_data: dict[str, list[dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        """One summary row per Policy Info insured, with the real child-row
        counts — keeps the scenario sheet consistent with generated data.

        Columns: Scenario ID, State, Type of Entity, Equipment Count,
        Additional Interest Count, Loss Count.
        """
        pi_rows = _get_policy_rows(previous_sheets_data)
        if not pi_rows:
            return []

        def _sheet(*frags: str) -> list[dict]:
            for name, data in previous_sheets_data.items():
                nl = name.lower()
                if all(f in nl for f in frags):
                    return data or []
            return []

        equip_rows = _sheet("equipment schedule")
        ai_rows = _sheet("additional interest")
        loss_rows = _sheet("loss history")

        def _count_by_tid(rows: list[dict], require_loss_amount: bool = False) -> dict[str, int]:
            counts: dict[str, int] = {}
            for r in rows:
                tid = _tid(r)
                if not tid:
                    continue
                if require_loss_amount:
                    # Only count populated loss rows (blank rows = "no losses").
                    ld_key = _find_col(r, "loss date")
                    if not str(r.get(ld_key, "")).strip():
                        continue
                counts[tid] = counts.get(tid, 0) + 1
            return counts

        equip_count = _count_by_tid(equip_rows)
        ai_count = _count_by_tid(ai_rows)
        loss_count = _count_by_tid(loss_rows, require_loss_amount=True)

        def _hdr(*frags: str) -> str | None:
            for h in unique_headers:
                hl = h.lower()
                if all(f in hl for f in frags):
                    return h
            return None

        sid_key = _hdr("scenario")
        state_key = _hdr("state")
        entity_key = _hdr("type of entity") or _hdr("entity")
        equip_key = _hdr("equipment", "count")
        ai_key = _hdr("additional interest", "count") or _hdr("additional", "count")
        loss_key = _hdr("loss", "count")

        pi_state_key = _find_col(pi_rows[0], "binding state") or _find_col(pi_rows[0], "state")
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
            if equip_key:
                row[equip_key] = equip_count.get(tid, 0)
            if ai_key:
                row[ai_key] = ai_count.get(tid, 0)
            if loss_key:
                row[loss_key] = loss_count.get(tid, 0)
            out.append(row)

        return out
