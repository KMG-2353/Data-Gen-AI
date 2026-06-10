"""Engine: emit prompt constraints and apply validators for a rule set.

Both entry points evaluate each rule's firing condition before it acts, so a
rule that doesn't apply to the current sheet/row neither contributes prompt text
nor mutates values. The engine has no dependency on any handler module.
"""
from __future__ import annotations

from typing import Any, Iterable

from .rules import CountRule, Rule, RuleContext


def emit_prompt_constraints(
    rules: Iterable[Rule], sheet: str, ctx: RuleContext | None = None
) -> str:
    """Join prompt fragments for every rule that fires for ``sheet``.

    Evaluated at the sheet level (no row), so conditions are sheet-scoped here.
    Returns a newline-joined block, or an empty string when nothing fires.
    """
    base = ctx or RuleContext()
    sheet_ctx = RuleContext(
        sheet=sheet, row=None, policy_type=base.policy_type, extra=base.extra
    )
    fragments: list[str] = []
    for rule in rules:
        if not rule.fires(sheet_ctx):
            continue
        fragment = rule.prompt_fragment(sheet_ctx)
        if fragment:
            fragments.append(fragment)
    return "\n".join(fragments)


def apply_validators(
    rules: Iterable[Rule],
    rows: list[dict[str, Any]],
    sheet: str,
    ctx: RuleContext | None = None,
) -> list[dict[str, Any]]:
    """Apply each firing rule's deterministic validator to every row in place.

    CountRules carry no per-value enforcement and are skipped here. The pass is
    idempotent: a value snapped once is already canonical the second time.
    """
    base = ctx or RuleContext()
    value_rules = [r for r in rules if not isinstance(r, CountRule)]
    for row in rows:
        row_ctx = RuleContext(
            sheet=sheet, row=row, policy_type=base.policy_type, extra=base.extra
        )
        for rule in value_rules:
            if rule.fires(row_ctx):
                rule.apply_to_row(row, row_ctx)
    return rows
