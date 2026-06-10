"""Firing-condition predicates.

Profiles decide *cross-template* applicability (which rules a template carries);
conditions decide *row/sheet-level* firing within a template at runtime. A
condition is any callable ``(RuleContext) -> bool``.

Conditions are intentionally lightweight — sheet-exists / field-equals style
checks — mirroring the substring header-matching already used by ``_find_col``
in the handlers.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable

if TYPE_CHECKING:  # avoid a runtime import cycle; RuleContext lives in rules.py
    from .rules import RuleContext

Condition = Callable[["RuleContext"], bool]


def always() -> Condition:
    """A condition that always fires."""

    def _cond(ctx: "RuleContext") -> bool:
        return True

    return _cond


def sheet_is(*fragments: str) -> Condition:
    """Fire when the context sheet name contains ANY of the given fragments.

    Matching is case-insensitive substring, matching how handlers detect sheet
    types (e.g. ``"policy information"``, ``"sched of vehicle"``).
    """
    lowered = [f.lower() for f in fragments]

    def _cond(ctx: "RuleContext") -> bool:
        sheet = (ctx.sheet or "").lower()
        return any(frag in sheet for frag in lowered)

    return _cond


def field_equals(field_keywords: tuple[str, ...] | str, value: str) -> Condition:
    """Fire when the context row's matching column equals ``value``.

    ``field_keywords`` is AND-matched (all substrings must appear) against the
    row's column names, the same idiom as ``_find_col``. Comparison is
    case-insensitive and whitespace-trimmed. When there is no row in context
    (sheet-level emit), the condition does not fire.
    """
    keywords = (field_keywords,) if isinstance(field_keywords, str) else tuple(field_keywords)
    target = value.strip().lower()

    def _cond(ctx: "RuleContext") -> bool:
        if ctx.row is None:
            return False
        for key in ctx.row:
            kl = key.lower()
            if all(k.lower() in kl for k in keywords):
                return str(ctx.row[key]).strip().lower() == target
        return False

    return _cond
