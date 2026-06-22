"""Single-sourced pure primitives for dates, numbers, and row/column access.

This is the base layer: it imports nothing from ``app.policies`` or
``app.llm_service`` so the import direction stays handlers -> rulebook. Every
handler and both date-enforcement paths route their date/number/string helpers
here, replacing the copies that had drifted (notably ``_format_mmddyyyy``, which
existed in two files with *different* output formats).

Every formatter is idempotent on already-formatted input.
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any, Sequence


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


def tid_value(row: dict) -> str:
    """The row's ``Test ID`` value (cross-sheet join key), trimmed."""
    k = next((key for key in row if key.lower().strip() == "test id"), None)
    return str(row.get(k, "")).strip() if k else ""


def is_yes(val: Any) -> bool:
    return str(val or "").strip().lower() in ("yes", "y", "true", "1")


def is_no(val: Any) -> bool:
    return str(val or "").strip().lower() in ("no", "n", "false", "0")


# NOTE (S1 finding): the two date-enforcement paths historically emitted
# DIFFERENT formats under the same name `_format_mmddyyyy`:
#   - app/policies/ims.py        -> format_date_slash   "06/22/2026"
#   - app/llm_service.py         -> format_date_compact "06222026"
# This slice preserves both (parity). Which is correct is a QA decision tracked
# for slice S2 (validation harness) — do not unify without sign-off.
