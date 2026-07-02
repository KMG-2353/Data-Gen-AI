"""Single-sourced pure primitives for dates, numbers, and row/column access.

This is the base layer: it imports nothing from ``app.policies`` or
``app.llm_service`` so the import direction stays handlers -> rulebook. Every
handler and both date-enforcement paths route their date/number/string helpers
here, replacing the copies that had drifted (notably ``_format_mmddyyyy``, which
existed in two files with *different* output formats).

Every formatter is idempotent on already-formatted input.
"""
from __future__ import annotations

import re
import random
from datetime import date, datetime
from typing import Any, Sequence


def normalize_sheet_name(name: str) -> str:
    """Canonicalise a sheet name for type detection.

    Newer SPG blank templates prefix sheets with an ordinal + LOB token and use
    underscores (``01_Policy_Info``, ``03_IM_Equipment``, ``09_APD_LossPayees``)
    where the older templates used plain spaced names (``Policy Info``, ``IM
    LossPayees``). This strips a leading ``NN_`` ordinal and turns underscores
    into spaces so a single set of substring checks matches both generations of
    template. Idempotent on already-spaced names.
    """
    s = str(name or "").strip()
    s = re.sub(r"^\s*\d+\s*[_\-\s]+", "", s)   # drop a leading "01_" / "10 - " ordinal
    s = s.replace("_", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def parse_date(val: Any) -> date | None:
    """Parse MM/DD/YYYY, MMDDYYYY, or YYYY-MM-DD into a date, else None."""
    s = str(val or "").strip()
    for fmt in ("%m/%d/%Y", "%m%d%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def format_date_slash(d: date) -> str:
    """MM/DD/YYYY, e.g. ``06/22/2026`` (the IMS/handler date format)."""
    return d.strftime("%m/%d/%Y")


def format_date_compact(d: date) -> str:
    """MMDDYYYY, e.g. ``06222026`` (the generic llm_service date format)."""
    return d.strftime("%m%d%Y")


def add_one_year(d: date) -> date:
    """Same calendar date one year later; Feb 29 -> Feb 28 on non-leap years."""
    try:
        return d.replace(year=d.year + 1)
    except ValueError:
        return d.replace(month=2, day=28, year=d.year + 1)


def pin_quote_effective_expiration(row: dict, today: date | None = None) -> None:
    """SPG-wide policy-date rule (single source of truth).

    The QA-approved DW/HO behaviour (``spg_pl``) is now the standard for **every**
    SPG LOB (IM, Cargo, APD, Wind/Hail): the Date of Quote is the data-creation
    date — always *today*, never a past or future date — the Effective Date is
    clamped to be on or after today (coverage cannot begin before the quote is
    created), and the Expiration Date is derived as Effective + 1 year.

    Closes the SPG quote-date defect class in one place: DF-IM-020/022, WH-005,
    APD-016, CARGO-008(quote), alongside the DW/HO originals DEF-002/023,
    HO-002/022. Idempotent and safe to call on any row (no-ops on missing cols).
    """
    today = today or date.today()
    eff_key = find_col(row, "effective date")
    exp_key = find_col(row, "expiration date")
    # "Quote Date" (numbered templates) / "Date of Quote" (legacy IM template).
    quote_key = find_col(row, "quote date") or find_col(row, "date of quote")

    eff = parse_date(row.get(eff_key)) if eff_key else None
    if eff and eff < today:
        eff = today
        if eff_key:
            row[eff_key] = format_date_slash(eff)
    if eff and exp_key:
        row[exp_key] = format_date_slash(add_one_year(eff))
    if quote_key:
        row[quote_key] = format_date_slash(today)


def to_number(val: Any) -> float | int | None:
    """Strip ``$``, commas, whitespace; return int when whole, else 2dp float."""
    s = str(val or "").strip().replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else round(f, 2)


def find_col(row: dict, *keywords: str) -> str | None:
    """First key whose lowercase form contains ALL keyword substrings."""
    for key in row:
        kl = key.lower()
        if all(k.lower() in kl for k in keywords):
            return key
    return None


def find_header_key(row: dict, candidates: Sequence[str]) -> str | None:
    """First key whose lowercase form contains ANY candidate substring."""
    for key in row.keys():
        key_lower = key.lower()
        if any(c in key_lower for c in candidates):
            return key
    return None


def default_test_id(index: int, width: int = 3) -> str:
    """Global default Test ID: ``TS-001`` (1-based, zero-padded to ``width``).

    GLOBAL RULE: Test IDs use the ``TS-`` prefix unless a specific rulebook
    *explicitly* mandates another (e.g. RRG -> ``DS-``). Handlers without an
    explicit prefix in their source rulebook MUST route here rather than
    hardcoding a prefix, so the default stays single-sourced.
    """
    return f"TS-{index:0{width}d}"


def tid_value(row: dict) -> str:
    """The row's ``Test ID`` value (cross-sheet join key), trimmed."""
    k = next((key for key in row if key.lower().strip() == "test id"), None)
    return str(row.get(k, "")).strip() if k else ""


def format_us_phone(val: Any) -> str:
    """Format any phone/fax value as a standard U.S. number ``(XXX) XXX-XXXX``.

    Strips to digits, drops a leading country-code ``1`` on 11-digit input, then
    renders the canonical U.S. format. When the input cannot yield 10 digits
    (the field is unusable as-is) a plausible 10-digit number is generated so the
    cell is still a valid U.S. phone number. Idempotent on already-formatted
    input. Used for continuous-digit phone/fax cells the LLM emits unformatted.
    """
    digits = "".join(c for c in str(val or "") if c.isdigit())
    if len(digits) == 11 and digits[0] == "1":
        digits = digits[1:]
    if len(digits) != 10:
        digits = f"{random.randint(200, 999)}{random.randint(200, 999)}{random.randint(0, 9999):04d}"
    return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"


def coerce_to_allowed(
    val: Any, allowed: Sequence[Any], default: Any, *, fill_blank: bool = False
) -> Any:
    """Snap a value to a member of ``allowed`` (case/space-insensitive match).

    Returns the canonical allowed option when ``val`` matches one; otherwise
    ``default``. Blank values are left untouched unless ``fill_blank`` is set
    (for mandatory dropdown fields that must always carry a value). The allowed
    set is the ruleset's dropdown list, passed in by the caller — this primitive
    holds no domain values itself.
    """
    s = str(val or "").strip()
    if not s:
        return default if fill_blank else val
    low = s.lower()
    for opt in allowed:
        if low == str(opt).strip().lower():
            return opt
    return default


# ---------------------------------------------------------------------------
# US state codes — single source of the address-state domain.
#
# Sourced verbatim from the raters' own full-state dropdown (the
# ``15_LKP_Dropdowns`` column the Address/Garaging/Mailing/Agency/Insured/
# Location/Loss-Payee state fields validate against, e.g. SPG IM/Cargo/APD list
# ``AL,AK,AZ,...`` and the SPG PL list ``$A$5:$A$54``). These are the *physical
# address* states — distinct from a template's restricted **Binding/Rating
# State** dropdown (a small per-template subset). The recurring QA defect is that
# the agent collapses address-state fields onto the binding subset instead of
# exercising this full set; this list is the shared, single-sourced pool the
# variety pass spreads across (it is US domain data, not a per-template magic
# list — handlers/profiles never re-declare it).
# ---------------------------------------------------------------------------
US_STATES: tuple[str, ...] = (
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
)
# DC is offered by some raters (IM/Cargo/APD) but not the SPG PL list; callers
# that want it pass this wider pool explicitly.
US_STATES_DC: tuple[str, ...] = US_STATES + ("DC",)

_US_STATE_SET = frozenset(US_STATES_DC)


def is_us_state(val: Any) -> bool:
    """True when ``val`` is a 2-letter US state (or DC) code, case-insensitive."""
    return str(val or "").strip().upper() in _US_STATE_SET


def is_binding_state_field(name: str) -> bool:
    """True when a column is the restricted **Binding/Rating** state dropdown.

    Every other ``*state*`` column on these raters is a physical-address state
    that validates against the full :data:`US_STATES` list. This single predicate
    lets the variety pass tell the two apart by header name alone (no per-template
    column list), so it stays adaptive to new raters that follow the convention.
    """
    nl = str(name or "").lower()
    return "state" in nl and ("binding" in nl or "rating" in nl)


def spread_pick(
    index: int,
    allowed: Sequence[Any],
    *,
    seed: int = 0,
    avoid: Any = None,
) -> Any:
    """Deterministically pick one of ``allowed`` for row ``index``.

    Cycles ``allowed`` by ``(seed + index)`` so a column's values fan out across
    rows (guaranteeing the full set is exercised once enough rows exist) while
    staying stable/idempotent for a given (seed, index). ``avoid`` is dropped from
    the pool first (used to keep an address state off the row's binding state).
    Returns ``None`` only when the pool is empty after exclusion.
    """
    pool = [a for a in allowed if a != avoid] or list(allowed)
    if not pool:
        return None
    return pool[(int(seed) + int(index)) % len(pool)]


def column_seed(name: str) -> int:
    """Stable small non-negative seed derived from a column name.

    Lets sibling columns (Agency State vs Mailing State vs Address State) spread
    on different phases so a single row shows *combinations* of states rather than
    the same value repeated across its address fields.
    """
    return sum(ord(c) for c in str(name or "")) % 97


def is_yes(val: Any) -> bool:
    return str(val or "").strip().lower() in ("yes", "y", "true", "1")


def is_no(val: Any) -> bool:
    return str(val or "").strip().lower() in ("no", "n", "false", "0")


# NOTE (S1 finding, RESOLVED in S2): the two date-enforcement paths historically
# emitted DIFFERENT formats under the same name `_format_mmddyyyy`:
#   - app/policies/ims.py        -> format_date_slash   "06/22/2026"
#   - app/llm_service.py         -> format_date_compact "06222026"  (the bug)
# The compact path shipped Effective/Expiration as "07272026" (DF-IM-001 sibling).
# S2 decision (QA sign-off): UNIFIED to MM/DD/YYYY everywhere — llm_service now
# imports format_date_slash. format_date_compact is retained only for parsing
# legacy/compact input via parse_date; no app code emits it.
