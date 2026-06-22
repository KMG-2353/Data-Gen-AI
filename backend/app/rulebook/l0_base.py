"""L0 generic base rule pool.

Universal rules that apply to almost every template — phone, ZIP, state code,
SSN, dollar-as-number, date format — authored from the scattered guidance in
``app/llm_service.py`` ("IMPORTANT RULES" block) so they become an owned,
reusable layer instead of inline prompt text.

Per-template profiles inherit these by id, and may drop or override them (e.g.
RRG overrides ``l0.state`` with its 8-state enum, and drops ``l0.ssn`` because it
has no SSN field). See ``profiles.py``.

The effective/expiration **date** logic is intentionally NOT re-expressed here as
a value formatter — it stays in the existing
``_enforce_effective_expiration_date_range`` code hook to avoid date regressions
(plan U2). ``l0.date`` carries only prompt guidance.

Note: ``apply_to_row`` snaps the first column matching a rule's keywords. The L0
pool is a generic safety net; templates needing per-column precision (multiple
state columns, etc.) override with specific L1 rules in their profile.
"""
from __future__ import annotations

from .rules import FormatRule, Rule


# --- generic formatters (kept local; the rulebook never imports handlers) ----

def format_phone10(value) -> str:
    """Strip to a bare 10-digit phone (generic intent: no dashes/parens).

    Leaves non-10-digit input unchanged so a template with its own phone format
    (e.g. RRG's spaced "XXX XXX XXXX") can override without this mangling it.
    """
    digits = "".join(c for c in str(value or "") if c.isdigit())
    return digits if len(digits) == 10 else (value if value is not None else "")


def format_zip5(value):
    """Normalize to a strict 5-digit ZIP, preserving leading zeros."""
    digits = "".join(c for c in str(value or "") if c.isdigit())
    if not digits:
        return value
    if len(digits) < 5:
        return digits.zfill(5)
    return digits[:5]


def format_state2(value):
    """Uppercase/trim to a 2-letter-style state code (generic)."""
    if value is None:
        return value
    return str(value).strip().upper()


def format_ssn(value):
    """Format 9 digits as XXX-XX-XXXX; leave other input unchanged."""
    digits = "".join(c for c in str(value or "") if c.isdigit())
    if len(digits) == 9:
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
    return value


# Default test-case-id convention. When a rater template carries a test-case /
# scenario id column but defines no convention of its own, ids are stamped
# sequentially as ``TS-01, TS-02, …`` — the same convention IMS uses. Raters
# inherit this as their default (handlers import this helper so the L0 base rule
# stays the single source of the format).
def format_test_case_id(index: int) -> str:
    """1-based row index -> default ``TS-01`` test-case id (zero-padded to 2)."""
    return f"TS-{int(index):02d}"


# --- L0 rule definitions -----------------------------------------------------

def l0_rules() -> list[Rule]:
    """The L0 generic base rule pool (fresh instances per call)."""
    return [
        FormatRule(
            id="l0.phone",
            field_keywords=("phone",),
            formatter=format_phone10,
            prompt_text="- Phone: 10 digits, no dashes or parentheses.",
        ),
        FormatRule(
            id="l0.zip",
            field_keywords=("zip",),
            formatter=format_zip5,
            prompt_text="- ZIP: 5 digits (numeric only).",
        ),
        FormatRule(
            id="l0.state",
            field_keywords=("state",),
            formatter=format_state2,
            prompt_text="- State: 2-letter US abbreviation (uppercase).",
        ),
        FormatRule(
            id="l0.ssn",
            field_keywords=("ssn",),
            formatter=format_ssn,
            prompt_text="- SSN: XXX-XX-XXXX format (fake but valid).",
        ),
        Rule(
            id="l0.dollar",
            field_keywords=(),
            prompt_text=(
                "- Dollar values: numeric (not strings), prefixed with a $ sign."
            ),
        ),
        # Date enforcement stays in _enforce_effective_expiration_date_range (code
        # hook); this L0 rule contributes prompt guidance only.
        Rule(
            id="l0.date",
            field_keywords=(),
            prompt_text=(
                "- Effective/Expiration dates: consistent format across sheets; "
                "expiration exactly 1 year after effective."
            ),
        ),
        # Default test-case-id convention: when a template defines no convention
        # of its own, the test-case / scenario id column uses sequential TS-01,
        # TS-02, … (the IMS convention). Deterministic per-row stamping happens in
        # the handlers (which import ``format_test_case_id``) because it needs the
        # row index; this base rule carries the prompt guidance.
        Rule(
            id="l0.test_case_id",
            field_keywords=(),
            prompt_text=(
                "- Test ID / Test Case No / Scenario ID: when no other convention "
                "is specified, use sequential TS-01, TS-02, TS-03 … (zero-padded). "
                "Keep the SAME id across every sheet describing the same "
                "scenario/insured."
            ),
        ),
    ]


def l0_by_id() -> dict[str, Rule]:
    """Map of L0 rule id -> rule instance, for profile inherit/drop/override."""
    return {r.id: r for r in l0_rules()}
