"""RRG L1 profile + extracted rules.

This module owns the canonical RRG value lists and formatters (moved here from
the handler so there is a single source of truth — the handler imports them
back, preserving exact behavior). The four extracted rules — Contact Number,
State, ZIP, Org Type — dual-emit their prompt fragment and their deterministic
validator; ``rrg.py`` delegates the post-process value enforcement for exactly
these fields to them (parity-gated by ``config.RULEBOOK_ENABLED``).

Import direction: handlers import from here; this module never imports handlers.
"""
from __future__ import annotations

import random

from .conditions import sheet_is
from .profiles import Profile
from .rules import EnumRule, FormatRule


# --- canonical RRG value lists (single source of truth) ----------------------

VALID_STATES = ["NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ"]

VALID_ORG_TYPES = [
    "Corporation", "Individual", "Joint Venture", "Limited Partnership",
    "LLC", "Other", "Partnership", "Trust",
]


# --- canonical RRG formatters (single source of truth) -----------------------

def rrg_format_zip5(val) -> str | None:
    """Normalize a ZIP to a strict 5-digit string (Rules 3 / 10 / 28).

    Strips non-digits, left-pads to 5 (preserving leading zeros for CT/NJ ZIPs
    like 06103) or trims to 5. Returns None when the input had no digits at all,
    so callers can fall back on the insured's ZIP.
    """
    digits = "".join(ch for ch in str(val or "") if ch.isdigit())
    if not digits:
        return None
    if len(digits) < 5:
        return digits.zfill(5)
    return digits[:5]


def rrg_format_us_phone(val) -> str:
    """Format a contact/phone value as a USA 10-digit number "XXX XXX XXXX".

    Groups 3-3-4 separated by single spaces (212 555 0147) per Rule 1 / Rule 9.
    A non-10-digit value is replaced by a fabricated valid 10-digit number
    (area code not starting with 0 or 1).
    """
    digits = "".join(ch for ch in str(val or "") if ch.isdigit())
    if len(digits) != 10:
        digits = str(random.randint(2, 9)) + "".join(
            str(random.randint(0, 9)) for _ in range(9)
        )
    return f"{digits[:3]} {digits[3:6]} {digits[6:]}"


# --- extracted rules (Contact Number, State, ZIP, Org Type) ------------------
# Each rule mirrors the exact field-matching and snapping the handler used, so
# the engine path reproduces handler output byte-for-byte (parity).

_ON_POLICY = sheet_is("policy information")


RRG_CONTACT_RULE = FormatRule(
    id="rrg.contact.phone",
    field_match=lambda name: name.lower() == "contact number",
    formatter=rrg_format_us_phone,
    condition=_ON_POLICY,
    prompt_text=(
        '- Contact Number: USA 10-digit phone in "XXX XXX XXXX" format — three '
        "groups of 3-3-4 separated by single spaces (e.g. 212 555 0147). No "
        "dashes, parentheses, or other special chars [Rule 1 / DS_044]"
    ),
)

RRG_STATE_RULE = EnumRule(
    id="rrg.state.enum",
    field_match=lambda name: name.lower() in ("state", "state of operation", "rating state"),
    all_columns=True,
    allowed=tuple(VALID_STATES),
    snap=lambda v, allowed: random.choice(allowed),
    condition=_ON_POLICY,
    prompt_text=f"- State: MUST be one of {VALID_STATES} [Rule 2]",
)

RRG_ZIP_RULE = FormatRule(
    id="rrg.zip.format",
    field_match=lambda name: name.lower() in ("zip code", "zip"),
    # Returns None when the value has no digits, matching the handler's
    # ``cleaned = _format_zip5(val); if cleaned: assign`` guarded write.
    formatter=rrg_format_zip5,
    condition=_ON_POLICY,
    prompt_text="- ZIP Code: 5-digit numeric only [Rule 3]",
)

RRG_ORG_RULE = EnumRule(
    id="rrg.org.enum",
    field_match=lambda name: "org type" in name.lower()
    or ("entity" in name.lower() and "org" in name.lower()),
    all_columns=True,
    allowed=tuple(VALID_ORG_TYPES),
    snap=lambda v, allowed: random.choice(allowed),
    condition=_ON_POLICY,
    prompt_text=f"- Org Type / Entity: MUST be one of {VALID_ORG_TYPES} [Rule 4]",
)

# Order matches the prompt lines they replace in build_sheet_context.
RRG_EXTRACTED_RULES = [RRG_CONTACT_RULE, RRG_STATE_RULE, RRG_ZIP_RULE, RRG_ORG_RULE]


def rrg_profile() -> Profile:
    return Profile(
        policy_type="RRG",
        drops=frozenset({"l0.ssn"}),  # RRG workbooks have no SSN field
        overrides={},
        added=list(RRG_EXTRACTED_RULES),
    )
