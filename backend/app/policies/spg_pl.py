"""SPG Personal Lines Excel Rater handlers — Dwelling Fire (DW) and Homeowners (HO).

Both raters share the same shape: a common ``Policy Info`` sheet, an
underwriting-history sheet, a property sheet, a coverages sheet, child schedules
(loss payees, loss history, and — DW only — locations), and a ``Test Scenario
Details`` summary. Every sheet joins on the ``Test ID`` column.

These handlers follow the SPG DW / SPG Homeowners rulesets: each sheet injects
its hard constraints into the LLM prompt, and post-processing deterministically
enforces the dependency rules the LLM gets wrong (blank-when-No fields, numeric
coercions, date windows). ``Test ID`` is stamped with the default ``TS-01``
convention from the L0 base rule (``format_test_case_id``) on Policy Info and
reused across every sheet, and ``Test Scenario Details`` is rebuilt
deterministically as a per-scenario summary of the real child-row counts.
"""
from __future__ import annotations

import random
from datetime import date, datetime, timedelta
from typing import Any

from app.rulebook.l0_base import format_test_case_id, format_zip5


# ---------------------------------------------------------------------------
# Shared helpers (mirror app/policies/im.py)
# ---------------------------------------------------------------------------

def _parse_date(val: Any) -> date | None:
    s = str(val or "").strip()
    for fmt in ("%m/%d/%Y", "%m%d%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _fmt_date(d: date) -> str:
    return d.strftime("%m/%d/%Y")


def _to_number(val: Any) -> float | int | None:
    s = str(val or "").strip().replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else round(f, 2)


def _find_col(row: dict, *keywords: str) -> str | None:
    """First key whose lowercase form contains ALL keyword substrings."""
    for key in row:
        kl = key.lower()
        if all(k.lower() in kl for k in keywords):
            return key
    return None


def _tid_key(row: dict) -> str | None:
    return next((k for k in row if k.lower().strip() == "test id"), None)


def _tid(row: dict) -> str:
    k = _tid_key(row)
    return str(row.get(k, "")).strip() if k else ""


def _is_yes(val: Any) -> bool:
    return str(val or "").strip().lower() in ("yes", "y", "true", "1")


def _is_no(val: Any) -> bool:
    return str(val or "").strip().lower() in ("no", "n", "false", "0")


def _blank_fields(row: dict, *fragments: str) -> None:
    """Blank every column matching any of the keyword fragments."""
    for frag in fragments:
        k = _find_col(row, frag)
        if k:
            row[k] = ""


def _normalize_common(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Snap every sheet's date columns to MM/DD/YYYY and ZIP columns to 5 digits.

    The rulesets require MM/DD/YYYY dates (Rule 4/5) and 5-digit ZIPs (Rule 20);
    the LLM frequently emits MMDDYYYY dates and ZIP+4 codes. This pass is generic
    and idempotent, so it is safe to run on already-normalized values.
    """
    for row in rows:
        for key in list(row.keys()):
            kl = key.lower()
            val = row.get(key)
            if val in (None, ""):
                continue
            if "zip" in kl:
                row[key] = format_zip5(val)
            elif "date" in kl or "dob" in kl:
                d = _parse_date(val)
                if d is not None:
                    row[key] = _fmt_date(d)
    return rows


def _get_sheet(previous: dict | None, *frags: str) -> list[dict]:
    if not previous:
        return []
    for name, rows in previous.items():
        nl = name.lower()
        if all(f in nl for f in frags):
            return rows or []
    return []


def _policy_info_rows(previous: dict | None) -> list[dict]:
    return _get_sheet(previous, "policy info")


# The approved binding states + the two products, shared by both rulesets
# (DW Rule 1/2, HO Rule 1/2).
_BINDING_STATES = ("VA", "MD", "PA", "NC")
_PRODUCTS = ("HO-3 Homeowners", "Dwelling Fire DP-3")


# ---------------------------------------------------------------------------
# Shared base handler
# ---------------------------------------------------------------------------

class _SpgPersonalLinesHandler:
    """Common DW/HO behavior. Subclasses set ``policy_type`` and the per-LOB
    property/coverages prompt rules + sheet-name fragments."""

    policy_type = "SPG_PL"
    # Subclasses override these sheet-name fragments.
    _property_frags: tuple[str, ...] = ()       # DF Locations / HO Dwelling
    _product_default = "Dwelling Fire DP-3"

    # ------------------------------------------------------------------
    # Sheet type detection
    # ------------------------------------------------------------------

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = sheet_name.lower().strip()
        if "policy info" in sn:
            return "policy"
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        if "losspayee" in sn or "loss payee" in sn:
            return "loss_payees"
        if "losshistory" in sn or "loss history" in sn:
            return "loss_history"
        if "coverage" in sn:
            return "coverages"
        if any(f in sn for f in self._property_frags):
            return "property"
        # Remaining "<LOB> Policy" sheet = underwriting / applicant history.
        if sn.endswith("policy"):
            return "uw_history"
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

        # Test Scenario Details is built deterministically in pre_generate.
        if st == "scenario_details":
            return 0, ""

        tid_rule = (
            "- Test ID: use sequential TS-01, TS-02, TS-03 … (zero-padded to 2 "
            "digits). Reuse the SAME Test ID across every sheet for the same "
            "scenario."
        )
        reuse_rule = (
            "- Test ID: reuse the EXACT TS-## ids already present on Policy Info "
            "(shown in the cross-sheet context). One scenario per Test ID."
        )

        if st == "policy":
            rules = f"""
SPG PERSONAL LINES — POLICY INFO RULES (HARD CONSTRAINTS):
{tid_rule}
- Binding State: one of {', '.join(_BINDING_STATES)}. [Rule 1]
- Product Selected: one of {', '.join(_PRODUCTS)}. [Rule 2]
- Effective Date / Expiration Date / Quote Date: MM/DD/YYYY. Expiration = Effective + 1 year. Quote Date <= Effective Date. [Rule 4/5]
- Agent Commission: one of 10%,10.5%,…,20%. If Binding State = PA, Agent Commission MUST be 10%. [Rule 21]
- Type of Entity: one of Individual, Corporation, LLC, Partnership, Joint Venture, Estate, Trust. [Rule 23]
- Date of Birth: an adult (18-100), never in the future. [Rule 28]
- Additional Resident/Spouse?: Yes/No. If No, leave ALL "Add. Resident" fields blank. [Rule 29]
- Mailing Address Different?: Yes/No. If No, leave ALL Mailing fields blank. [Rule 39]
- All state fields must be 2-letter codes consistent with their city/ZIP.
"""
            return original_row_count, rules

        if st == "uw_history":
            return original_row_count, self._uw_history_rules(reuse_rule)

        if st == "property":
            return original_row_count, self._property_rules(reuse_rule)

        if st == "coverages":
            return original_row_count, self._coverages_rules(reuse_rule)

        if st == "loss_payees":
            rules = f"""
SPG PERSONAL LINES — LOSS PAYEES RULES (HARD CONSTRAINTS):
{reuse_rule}
- Generate realistic lender / bank / mortgage company names. State is a 2-letter code consistent with City/ZIP.
- Is Mortgagee?: Yes/No. If No, leave Loan Number and Mortgage Current? BLANK. If Yes, both MUST be populated (Loan Number = unique alphanumeric). [Rule 119/120/121]
- Generate 0-2 loss-payee rows per scenario; leave the sheet empty for scenarios with no mortgagee.
"""
            return original_row_count, rules

        if st == "loss_history":
            eff_hint = self._effective_date_hint(policy_data)
            rules = f"""
SPG PERSONAL LINES — LOSS HISTORY RULES (HARD CONSTRAINTS):
{reuse_rule}
- Any Open Claims?: Yes/No. [Rule 144/122]
- Any Losses in Past 5 Years?: Yes/No. If No, leave #, Loss Date, Type of Loss, Details, Amount BLANK and Unrepaired Damage blank. [Rule 145/123]
- Unrepaired Damage from Prior Losses?: Yes/No, required only when there were losses.
- Loss Date: MM/DD/YYYY, within the 5 years before the Effective Date AND earlier than it. [Rule 148/126]
- Type of Loss: one of Fire, Water Damage – Weather Related, Water Damage – Non-Weather Related, Wind Damage, Hail Damage, Theft, Physical Damage – All Other, Liability. [Rule 149/127]
- Amount ($): a PLAIN positive number > 0 (no $, no commas). [Rule 151/129]
- For scenarios WITH losses generate 1-3 loss rows; otherwise a single blank-loss row.
{eff_hint}
"""
            return original_row_count, rules

        return original_row_count, ""

    def _effective_date_hint(self, policy_data: list[dict] | None) -> str:
        if not policy_data:
            return ""
        lines = []
        for r in policy_data:
            eff_key = _find_col(r, "effective date")
            lines.append(f"  {_tid(r)}: Effective {r.get(eff_key, '')}")
        return (
            "\nPer-scenario Effective Dates (loss dates must fall in the 5 years "
            "before these):\n" + "\n".join(lines)
        )

    # --- per-LOB rule blocks (overridden by subclasses) ------------------
    def _uw_history_rules(self, reuse_rule: str) -> str:
        return ""

    def _property_rules(self, reuse_rule: str) -> str:
        return ""

    def _coverages_rules(self, reuse_rule: str) -> str:
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
            return self._build_scenario_details(
                unique_headers, previous_sheets_data or {}
            )
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
            return _normalize_common(self._fix_policy_info(rows))
        # Every other generated sheet joins back to Policy Info by Test ID.
        self._normalize_test_ids(rows, previous_sheets_data)
        if st == "uw_history":
            rows = self._fix_uw_history(rows)
        elif st == "property":
            rows = self._fix_property(rows)
        elif st == "coverages":
            rows = self._fix_coverages(rows)
        elif st == "loss_payees":
            rows = self._fix_loss_payees(rows)
        elif st == "loss_history":
            rows = self._fix_loss_history(rows, previous_sheets_data)
        return _normalize_common(rows)

    # ------------------------------------------------------------------
    # Test ID handling
    # ------------------------------------------------------------------

    def _fix_policy_info(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Stamp the default TS-## Test IDs and enforce Policy Info dependencies."""
        for idx, row in enumerate(rows):
            tk = _tid_key(row)
            if tk is not None:
                row[tk] = format_test_case_id(idx + 1)

            # Expiration = Effective + 1 year; Quote Date <= Effective.
            eff_key = _find_col(row, "effective date")
            exp_key = _find_col(row, "expiration date")
            quote_key = _find_col(row, "quote date")
            eff = _parse_date(row.get(eff_key)) if eff_key else None
            if eff:
                if exp_key:
                    try:
                        row[exp_key] = _fmt_date(eff.replace(year=eff.year + 1))
                    except ValueError:  # Feb 29
                        row[exp_key] = _fmt_date(eff + timedelta(days=365))
                if quote_key:
                    q = _parse_date(row.get(quote_key))
                    if q is None or q > eff:
                        row[quote_key] = _fmt_date(eff - timedelta(days=random.randint(1, 14)))

            # Rule 21: PA binding state forces 10% commission.
            bind_key = _find_col(row, "binding state")
            comm_key = _find_col(row, "agent commission")
            if bind_key and comm_key and str(row.get(bind_key, "")).strip().upper() == "PA":
                row[comm_key] = "10%"

            # Rule 29: blank additional-resident fields when No.
            ar_key = _find_col(row, "additional resident") or _find_col(row, "resident/spouse")
            if ar_key and _is_no(row.get(ar_key)):
                _blank_fields(row, "add. resident", "additional resident full",
                              "additional resident dob", "additional resident occupation",
                              "additional resident employer")

            # Rule 39: blank mailing fields when "Mailing Address Different?" = No.
            md_key = _find_col(row, "mailing address different")
            if md_key and _is_no(row.get(md_key)):
                _blank_fields(row, "mailing street", "mailing city",
                              "mailing state", "mailing zip")
        return rows

    def _normalize_test_ids(self, rows: list[dict], previous: dict | None) -> None:
        """Snap each child row's Test ID to the canonical TS-## set from Policy
        Info: zero-pad TS-N → TS-0N, and map any unrecognised id by row order."""
        canonical = [_tid(r) for r in _policy_info_rows(previous) if _tid(r)]
        canon_set = set(canonical)
        for i, row in enumerate(rows):
            tk = _tid_key(row)
            if tk is None:
                continue
            raw = str(row.get(tk, "")).strip().upper().replace(" ", "")
            digits = "".join(c for c in raw if c.isdigit())
            if raw.startswith("TS-") and digits:
                norm = format_test_case_id(int(digits))
                row[tk] = norm if (not canon_set or norm in canon_set) else (
                    canonical[i % len(canonical)] if canonical else norm
                )
            elif canonical:
                row[tk] = canonical[i % len(canonical)]

    # ------------------------------------------------------------------
    # Per-sheet fixers (shared dependency rules)
    # ------------------------------------------------------------------

    def _fix_uw_history(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Prior-insurance + termination + purchase dependencies."""
        for row in rows:
            prior_key = _find_col(row, "prior insurance on this account")
            if prior_key and _is_no(row.get(prior_key)):
                _blank_fields(
                    row, "previously with commonwealth", "prior carrier name",
                    "expiration date of prior", "prior policy number",
                    "prior insurance premium", "risk new to agency", "lapse",
                    "terminated at company", "reason for termination",
                    "previous wind/hail",
                )

            # Reason for Termination only when terminated at company request.
            term_key = _find_col(row, "terminated at company")
            reason_key = _find_col(row, "reason for termination")
            if term_key and reason_key and not _is_yes(row.get(term_key)):
                row[reason_key] = ""

            # Purchase dependencies (HO Rule 51-54 / DW New Purchase).
            np_key = (_find_col(row, "new purchase")
                      or _find_col(row, "is dwelling a new purchase"))
            if np_key and _is_no(row.get(np_key)):
                _blank_fields(row, "year purchased", "purchase price",
                              "was dwelling foreclosed", "was foreclosed")
        return rows

    def _fix_loss_payees(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Mortgagee dependency (Rule 119-121/139-141)."""
        for row in rows:
            mort_key = _find_col(row, "is mortgagee") or _find_col(row, "mortgagee")
            if mort_key and _is_no(row.get(mort_key)):
                _blank_fields(row, "loan number", "mortgage current")
        return rows

    def _fix_loss_history(
        self, rows: list[dict[str, Any]], previous: dict | None
    ) -> list[dict[str, Any]]:
        """Blank loss fields when no losses; clamp loss date into the 5-year
        pre-effective window; coerce Amount to a positive number."""
        eff_by_tid: dict[str, date] = {}
        for pi in _policy_info_rows(previous):
            ek = _find_col(pi, "effective date")
            eff = _parse_date(pi.get(ek)) if ek else None
            if eff:
                eff_by_tid[_tid(pi)] = eff

        for row in rows:
            losses_key = _find_col(row, "losses in past 5 years")
            if losses_key and _is_no(row.get(losses_key)):
                _blank_fields(row, "loss date", "type of loss", "details", "amount",
                              "unrepaired damage")
                num_key = next((k for k in row if k.strip() == "#"), None)
                if num_key:
                    row[num_key] = ""
                continue

            eff = eff_by_tid.get(_tid(row))
            ld_key = _find_col(row, "loss date")
            if ld_key and eff is not None:
                ld = _parse_date(row.get(ld_key))
                window_start = eff - timedelta(days=5 * 365)
                if ld is None or ld < window_start or ld >= eff:
                    span = (eff - timedelta(days=1) - window_start).days
                    row[ld_key] = _fmt_date(
                        window_start + timedelta(days=random.randint(0, max(span, 0)))
                    )

            amt_key = _find_col(row, "amount")
            if amt_key:
                num = _to_number(row.get(amt_key))
                if num is None or num <= 0:
                    num = random.randint(1500, 60000)
                row[amt_key] = int(round(num))
        return rows

    # Overridable per-LOB property/coverage fixers.
    def _fix_property(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return rows

    def _fix_coverages(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return rows

    # ------------------------------------------------------------------
    # Test Scenario Details (deterministic per-scenario summary)
    # ------------------------------------------------------------------

    def _build_scenario_details(
        self,
        unique_headers: list[str],
        previous: dict[str, list[dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        pi_rows = _policy_info_rows(previous)
        if not pi_rows:
            return []

        loss_payee_rows = _get_sheet(previous, "losspayee") or _get_sheet(previous, "loss payee")
        loss_rows = _get_sheet(previous, "losshistory") or _get_sheet(previous, "loss history")
        location_rows = _get_sheet(previous, "location")

        def _count(rows: list[dict], require_frag: str | None = None) -> dict[str, int]:
            counts: dict[str, int] = {}
            for r in rows:
                tid = _tid(r)
                if not tid:
                    continue
                if require_frag:
                    k = _find_col(r, require_frag)
                    if not str(r.get(k, "")).strip():
                        continue
                counts[tid] = counts.get(tid, 0) + 1
            return counts

        loc_count = _count(location_rows)
        lp_count = _count(loss_payee_rows)
        loss_count = _count(loss_rows, require_frag="loss date")

        def _hdr(*frags: str) -> str | None:
            for h in unique_headers:
                hl = h.lower()
                if all(f in hl for f in frags):
                    return h
            return None

        sid_key = _hdr("scenario")
        state_key = _hdr("state")
        product_key = _hdr("product")
        locc_key = _hdr("location count")
        lpc_key = _hdr("loss payee count")
        lossc_key = _hdr("loss count")

        pi_state_key = _find_col(pi_rows[0], "binding state") or _find_col(pi_rows[0], "state")
        pi_product_key = _find_col(pi_rows[0], "product selected")

        out: list[dict[str, Any]] = []
        for pi in pi_rows:
            tid = _tid(pi)
            row: dict[str, Any] = {}
            if sid_key:
                row[sid_key] = tid
            if state_key:
                row[state_key] = pi.get(pi_state_key, "") if pi_state_key else ""
            if product_key:
                row[product_key] = (pi.get(pi_product_key, "") if pi_product_key
                                    else self._product_default)
            if locc_key:
                row[locc_key] = loc_count.get(tid, 0)
            if lpc_key:
                row[lpc_key] = lp_count.get(tid, 0)
            if lossc_key:
                row[lossc_key] = loss_count.get(tid, 0)
            out.append(row)
        return out


# ---------------------------------------------------------------------------
# Dwelling Fire (DW) handler
# ---------------------------------------------------------------------------

class DwHandler(_SpgPersonalLinesHandler):
    policy_type = "DW"
    _property_frags = ("location",)         # DF Locations
    _product_default = "Dwelling Fire DP-3"

    def _uw_history_rules(self, reuse_rule: str) -> str:
        return f"""
SPG DW — APPLICANT / INSURANCE HISTORY RULES (HARD CONSTRAINTS):
{reuse_rule}
- Insured Credit History: one of Good, Fair, Poor. [Rule 45]
- Arson/Fraud, Bankruptcy, Foreclosure, Child Support, Repossessions: Yes/No. [Rule 46-50]
- Prior Insurance on This Account?: Yes/No. If No, leave Prior Carrier, Prior Expiration, Prior Policy Number, Prior Premium, Risk New to Agency, Lapse>30, Terminated, Reason, Previous Wind/Hail ALL blank. [Rule 51-60]
- Terminated at Company Request? = Yes requires a Reason for Termination; otherwise blank. [Rule 59]
- Is There a Management Company?: Yes/No (DF only). If No, Management Company Name/Phone blank. [Rule 61-63]
"""

    def _property_rules(self, reuse_rule: str) -> str:
        return f"""
SPG DW — DF LOCATIONS RULES (HARD CONSTRAINTS):
{reuse_rule}
- Loc #: sequential per scenario starting at 1. Generate 1+ locations per scenario.
- Coverage A ($): positive number. Exclude Cov B?: Yes/No. If Yes, Coverage B Value blank; if No, Coverage B Value populated. [Rule 68-70/115-116]
- Protection Class: 1-10. Good Condition / Existing Damage / Renovation / Single Family?: Yes/No. [Rule 71-75]
- Single Family? = Yes → # Families blank; = No → # Families is 2, 3, or 4. [Rule 75-76/117-118]
- Is Rented? = Yes → Rental Type is Tenant or Seasonal; = No → Rental Type blank. [Rule 77-78/119-120]
- Year Built ≤ current year; Sq Footage positive int; # Stories 1-4. [Rule 81-83]
- Foundation Type: Permanent Masonry (Slab/Crawlspace/Basement), Pilings, Pier, Other. Construction Type: Frame, Log Home, Masonry Veneer, Mixed - Frame and Masonry, Manufactured Home, Town Home. [Rule 84-85]
- Wood Burning Stove? = Yes → WBS Primary Heat populated; = No → blank. [Rule 88-89/121-122]
- Water Heater Yr / Roof Replacement Yr ≥ Year Built and ≤ current year. [Rule 90/93]
- Roofing Material: Asphalt Shingle, Metal, Slate, Wood Shake, Rubberized Membrane. Siding Material per Rule 94.
- >2 Acres? = No → >10 Acres? blank or No. [Rule 95-96/123-124]
- New Purchase? = Yes → Year Purchased + Purchase Price populated; = No → blank. [Rule 111-114/125-128]
- Yes/No columns must be exactly "Yes" or "No".
"""

    def _coverages_rules(self, reuse_rule: str) -> str:
        return f"""
SPG DW — DF COVERAGES RULES (HARD CONSTRAINTS):
{reuse_rule}
- Coverage E / L – Limit of Liability: one of $0, $25,000, $50,000, $100,000, $300,000, $500,000. [Rule 152]
- Loss of Rents (Coverage D): N/A or $1,000…$20,000 (thousands). [Rule 153]
- Owners Contents Coverage (Coverage C): N/A or $1,000…$10,000. Owners Contents Burglary Coverage: N/A, $1,000, $2,000. If Owners Contents = N/A, Burglary MUST be N/A. [Rule 154-155/163-164]
- Home Systems Protection / Service Line Coverage / Identity Theft Coverage / Central Station Alarms: Yes/No. [Rule 156-160]
- Higher Deductible (AOP): $1,000 (default) or $2,500. [Rule 159]
- Roof Valuation Endorsement: TYS 573-1 (default), TYS 573-2, TYS 573-3, TYS 573-4. [Rule 161/170]
- Any Additional Comments?: required (non-blank) when Foundation = Other. [Rule 162/171]
"""

    def _fix_property(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        for row in rows:
            # Exclude Cov B? = Yes → Coverage B Value blank. [Rule 115]
            excl_key = _find_col(row, "exclude cov b")
            covb_key = _find_col(row, "coverage b value")
            if excl_key and covb_key and _is_yes(row.get(excl_key)):
                row[covb_key] = ""

            # Single Family? = Yes → # Families blank. [Rule 117]
            sf_key = _find_col(row, "single family")
            fam_key = _find_col(row, "# families") or _find_col(row, "families")
            if sf_key and fam_key and _is_yes(row.get(sf_key)):
                row[fam_key] = ""

            # Is Rented? = No → Rental Type blank. [Rule 120]
            rent_key = _find_col(row, "is rented")
            if rent_key and _is_no(row.get(rent_key)):
                _blank_fields(row, "rental type", "rented to students",
                              "renters ins required")

            # Wood Burning Stove? = No → WBS Primary blank. [Rule 122]
            ws_key = _find_col(row, "wood burning stove")
            if ws_key and _is_no(row.get(ws_key)):
                _blank_fields(row, "wbs primary")

            # >2 Acres? = No → >10 Acres? blank. [Rule 123]
            acre_key = _find_col(row, ">2 acres")
            if acre_key and _is_no(row.get(acre_key)):
                _blank_fields(row, ">10 acres")

            # New Purchase? = No → Year Purchased / Purchase Price blank. [Rule 127-128]
            np_key = _find_col(row, "new purchase")
            if np_key and _is_no(row.get(np_key)):
                _blank_fields(row, "year purchased", "was foreclosed", "purchase price")
        return rows

    def _fix_coverages(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        for row in rows:
            # Owners Contents = N/A → Burglary = N/A. [Rule 163]
            oc_key = _find_col(row, "owners contents coverage")
            burg_key = _find_col(row, "burglary")
            if oc_key and burg_key and str(row.get(oc_key, "")).strip().upper() in ("N/A", "NA", ""):
                row[burg_key] = "N/A"

            # Roof Valuation default. [Rule 170]
            rv_key = _find_col(row, "roof valuation")
            if rv_key and not str(row.get(rv_key, "")).strip():
                row[rv_key] = "TYS 573-1"
        return rows


# ---------------------------------------------------------------------------
# Homeowners (HO) handler
# ---------------------------------------------------------------------------

class HoHandler(_SpgPersonalLinesHandler):
    policy_type = "HO"
    _property_frags = ("dwelling",)         # HO Dwelling
    _product_default = "HO-3 Homeowners"

    def _uw_history_rules(self, reuse_rule: str) -> str:
        return f"""
SPG HO — APPLICANT / INSURANCE HISTORY RULES (HARD CONSTRAINTS):
{reuse_rule}
- Insured Credit History: one of Good, Fair, Poor. [Rule 45]
- Arson/Fraud, Bankruptcy, Foreclosure, Child Support, Repossessions: Yes/No. [Rule 46-50]
- Is Dwelling a New Purchase?: Yes/No. If No, leave Year Purchased, Was Foreclosed, Purchase Price blank. [Rule 51-54]
- Year Purchased ≤ current year; Purchase Price a positive number. [Rule 52/54/67/69]
- Prior Insurance on This Account?: Yes/No. If No, leave Previously with Commonwealth, Prior Carrier, Prior Expiration, Prior Premium, Risk New to Agency, Lapse>30, Terminated, Reason, Previous Wind/Hail ALL blank. [Rule 55-65]
- Terminated at Company Request? = Yes requires a Reason for Termination; otherwise blank. [Rule 63/66]
- Previous Wind/Hail Deductible: one of 0,1,2,3,4,5 (only when prior insurance = Yes). [Rule 64]
"""

    def _property_rules(self, reuse_rule: str) -> str:
        return f"""
SPG HO — HO DWELLING RULES (HARD CONSTRAINTS):
{reuse_rule}
- Dwelling ZIP Code: 5 digits. Protection Class: 1-10. Year Built ≤ current year. [Rule 70-72]
- Square Footage: positive int (≥ 1,344 if Manufactured Home). Number of Stories: realistic int. [Rule 73-74]
- Is Single Family Residence? / Is Dwelling Owner Occupied?: should be Yes (else ineligible). [Rule 75-76]
- Type of Dwelling (Primary/Secondary). Is Dwelling Rented to Others? only for Secondary. [Rule 77-78]
- Is Dwelling a Manufactured Home?: Yes/No. [Rule 79]
- Good Condition? = Yes; Existing Damage? = No; Undergoing Renovation? = No (eligibility). [Rule 80-82]
- Type of Foundation: permanent masonry types preferred; Foundation Explain populated only when foundation = Other. [Rule 86-87]
- Wood Burning Stove? = Yes → Is Wood Stove Primary Heating populated; = No → blank. [Rule 89-90]
- Polybutylene/Qwest = No; Fuse Boxes = No; Aluminum Wiring = No; Knob/Tube = No; Lead Plumbing = No (eligibility). [Rule 91/93-96]
- Year of Last Water Heater/Roof Replacement ≤ current year. Roofing Material excludes Slate, Wood Shake, Rubberized Membrane. Roof Flat? = No. [Rule 92/101-103]
- Siding excludes Hardboard Composite, EIFS, Asbestos. [Rule 104]
- More Than 2 Acres? then More Than 10 Acres? applies; >10 acres ineligible. [Rule 105-106]
- Unfenced Pool / Animals With Bite History / Business Pursuits = No (eligibility). [Rule 107/110/112]
- Yes/No columns must be exactly "Yes" or "No".
"""

    def _coverages_rules(self, reuse_rule: str) -> str:
        return f"""
SPG HO — HO COVERAGES RULES (HARD CONSTRAINTS):
{reuse_rule}
- Coverage A — Replacement Cost ($): ≥ $100,000 (regular) or $75,000-$150,000 (Manufactured), TIV cap $750,000. [Rule 130]
- Coverage B — Other Structures (% of Cov A): Excluded, 10%, 20%, 30%, 40%. [Rule 131]
- Coverage C — Personal Property (% of Cov A): Excluded or 10%-70%. [Rule 132]
- Coverage D — Loss of Use (% of Cov A): Excluded, 10%, 20%. [Rule 133]
- Coverage E — Limit of Liability: Excluded, $0, $100,000, $200,000, $300,000. [Rule 134]
- Coverage F — Increased Medical Payments: $1,000, $5,000, $10,000 (only if Coverage E > $0). [Rule 135]
- Home Systems Protection / Service Line / Identity Theft / Replacement Cost / Extended Replacement Cost / Special Computer / Identity Fraud: Yes/No. Several not allowed for Manufactured homes. [Rule 136-142]
- Water Backup Coverage Limit: No Coverage or $1,000…$10,000, $25,000 ($10k+ blocked if Cov A < $350k). [Rule 143]
- Deadbolts/Smoke Alarms: Yes for non-Manufactured. Central Station Fire & Burglar Alarms: Yes/No. [Rule 144-146]
- Higher Deductible: $1,000, $2,500, $5,000. [Rule 147]
- Roof Valuation Endorsement: TYS 572 / 572-1 / 572-2 / 572-3 / 572-4. [Rule 148]
- Additional Comments: required when Foundation = Other or Reason for Termination = Other. [Rule 149]
"""

    def _fix_property(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        for row in rows:
            # Wood stove not present → primary-heating answer blank. [Rule 90]
            ws_key = _find_col(row, "wood burning stove")
            if ws_key and _is_no(row.get(ws_key)):
                _blank_fields(row, "wood stove primary", "primary heating")

            # More Than 2 Acres? = No → More Than 10 Acres? blank. [Rule 106]
            acre_key = _find_col(row, "more than 2 acres")
            if acre_key and _is_no(row.get(acre_key)):
                _blank_fields(row, "more than 10 acres")

            # Foundation Explain only when foundation = Other. [Rule 87]
            found_key = _find_col(row, "type of foundation")
            explain_key = _find_col(row, "foundation explain")
            if found_key and explain_key and str(row.get(found_key, "")).strip().lower() != "other":
                row[explain_key] = ""

            # Rented-to-others applies only to Secondary dwellings. [Rule 78]
            type_key = _find_col(row, "type of dwelling")
            rented_key = _find_col(row, "rented to others")
            if type_key and rented_key and str(row.get(type_key, "")).strip().lower() != "secondary":
                row[rented_key] = ""
        return rows

    def _fix_coverages(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        for row in rows:
            # Coverage F only available when Coverage E > $0. [Rule 135]
            e_key = _find_col(row, "coverage e")
            f_key = _find_col(row, "coverage f")
            if e_key and f_key:
                e_num = _to_number(row.get(e_key))
                e_excluded = str(row.get(e_key, "")).strip().lower() in ("excluded", "") or (e_num == 0)
                if e_excluded:
                    row[f_key] = ""

            # Roof Valuation default for regular homes. [Rule 148]
            rv_key = _find_col(row, "roof valuation")
            if rv_key and not str(row.get(rv_key, "")).strip():
                row[rv_key] = "TYS 572"
        return rows
