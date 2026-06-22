"""Rule abstractions and the dual-emit contract.

Every rule exposes two faces from one definition:

- ``prompt_fragment(ctx)`` — text injected into the LLM prompt (soft guidance).
- ``validate(value, ctx)`` / ``apply_to_row(row, ctx)`` — deterministic
  snap/format/cap applied in post-processing (hard correctness).

This keeps the prompt and the enforcement from drifting apart. The value-snapping
idioms mirror the handlers (``_normalize_vehicle_type``, ``_format_zip5``,
``_format_us_phone``) and the ``_find_col`` header matcher.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Sequence

from .conditions import Condition, always


@dataclass
class RuleContext:
    """Runtime context a rule's condition and validator read from.

    ``sheet`` is always set. ``row`` is the current row during post-process
    value enforcement, or ``None`` during sheet-level prompt emit. ``policy_type``
    is the detected template key. ``extra`` carries anything else a code hook
    wants to thread through (e.g. insured maps).
    """

    sheet: str = ""
    row: dict[str, Any] | None = None
    policy_type: str | None = None
    extra: dict[str, Any] = field(default_factory=dict)


def find_field(row: dict[str, Any], keywords: Sequence[str]) -> str | None:
    """Return the first column whose lowercase name contains ALL keywords.

    Mirrors ``_find_col`` in the handlers so extracted rules match the exact
    same columns the handler would have matched.
    """
    for key in row:
        kl = key.lower()
        if all(k.lower() in kl for k in keywords):
            return key
    return None


@dataclass
class Rule:
    """Base rule: a stable id, the columns it applies to, a firing condition,
    and a static prompt fragment. Subclasses add deterministic enforcement.

    ``field_keywords`` is AND-matched against column names. An empty tuple means
    the rule is not bound to a specific column (e.g. a sheet-level CountRule).
    ``field_match`` is an optional column-name predicate that takes precedence
    over ``field_keywords`` when the handler's original matching was exact-name
    or otherwise not a simple substring AND. ``all_columns`` snaps every matching
    column rather than just the first (mirrors handlers that loop all columns).
    """

    id: str
    field_keywords: tuple[str, ...] = ()
    condition: Condition = field(default_factory=always)
    prompt_text: str = ""
    field_match: Callable[[str], bool] | None = None
    all_columns: bool = False
    # Which side of the determinism boundary this rule's field sits on.
    # True  -> the engine produces/snaps the value; the LLM is not trusted.
    # False -> the field is open; the LLM generates within bounds.
    deterministic: bool = True

    def fires(self, ctx: RuleContext) -> bool:
        return self.condition(ctx)

    def _column_matches(self, name: str) -> bool:
        if self.field_match is not None:
            return self.field_match(name)
        if self.field_keywords:
            nl = name.lower()
            return all(k.lower() in nl for k in self.field_keywords)
        return False

    def matching_fields(self, row: dict[str, Any]) -> list[str]:
        keys = [k for k in row if self._column_matches(k)]
        if not self.all_columns:
            return keys[:1]
        return keys

    def find_field(self, row: dict[str, Any]) -> str | None:
        keys = self.matching_fields(row)
        return keys[0] if keys else None

    def prompt_fragment(self, ctx: RuleContext) -> str:
        """Static fragment by default; subclasses may compute dynamically."""
        return self.prompt_text

    def validate(self, value: Any, ctx: RuleContext) -> Any:
        """Deterministic value enforcement. Default: identity (no-op)."""
        return value

    def apply_to_row(self, row: dict[str, Any], ctx: RuleContext) -> dict[str, Any]:
        """Locate this rule's column(s) in ``row`` and snap the value(s) in place.

        No-op when the rule is not column-bound or no column matches. Snaps every
        matching column when ``all_columns`` is set, else only the first.
        """
        for key in self.matching_fields(row):
            row[key] = self.validate(row.get(key), ctx)
        return row


@dataclass
class FormatRule(Rule):
    """Snap a value to a canonical string form (phone, ZIP, date).

    ``formatter`` is a pure function ``value -> formatted_value``. It must be
    idempotent: ``formatter(formatter(x)) == formatter(x)``.
    """

    formatter: Callable[[Any], Any] | None = None

    def validate(self, value: Any, ctx: RuleContext) -> Any:
        if self.formatter is None:
            return value
        return self.formatter(value)


@dataclass
class EnumRule(Rule):
    """Snap a value to one of an allowed set.

    A value already in ``allowed`` is left untouched (so the rule is idempotent).
    An out-of-list value is replaced via ``snap`` — by default the first allowed
    value, but handlers that randomize can pass ``snap=lambda v, allowed: ...``.
    """

    allowed: tuple[str, ...] = ()
    snap: Callable[[Any, tuple[str, ...]], Any] | None = None
    case_insensitive: bool = False

    def _in_allowed(self, value: Any) -> bool:
        if self.case_insensitive:
            vl = str(value).strip().lower()
            return any(vl == a.lower() for a in self.allowed)
        return value in self.allowed

    def validate(self, value: Any, ctx: RuleContext) -> Any:
        if not self.allowed or self._in_allowed(value):
            return value
        if self.snap is not None:
            return self.snap(value, self.allowed)
        return self.allowed[0]


@dataclass
class CountRule(Rule):
    """Child-row multiplicity for a sheet (scenario-driven counts).

    Not a value snapper — it caps a requested per-insured count to ``max_count``
    and reports whether the cap fired, so a scenario asking for 25 vehicles
    against a max of 20 yields 20 plus a surfaced adjustment (R11).
    """

    max_count: int = 0
    min_count: int = 1

    def cap(self, requested: int) -> tuple[int, bool]:
        """Return (effective_count, was_capped)."""
        capped = requested
        was_capped = False
        if self.max_count and requested > self.max_count:
            capped = self.max_count
            was_capped = True
        if capped < self.min_count:
            capped = self.min_count
        return capped, was_capped
