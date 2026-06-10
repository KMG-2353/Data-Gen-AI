"""Layered, declarative rulebook engine.

A rule is authored once and dual-emits: a prompt-constraint fragment injected
into the LLM prompt AND a deterministic post-process validator/snapper. Rules
live in two layers — an L0 generic base pool plus L1 per-template profiles that
inherit / drop / override / add — and the engine composes the effective set per
template (see ``profiles.py``).

Import direction is strictly ``handlers -> rulebook``; the rulebook never
imports from ``app.policies`` so it stays reusable and testable in isolation.
"""
from __future__ import annotations

from .conditions import always, field_equals, sheet_is
from .rules import CountRule, EnumRule, FormatRule, Rule, RuleContext
from .engine import apply_validators, emit_prompt_constraints

__all__ = [
    "always",
    "field_equals",
    "sheet_is",
    "Rule",
    "RuleContext",
    "FormatRule",
    "EnumRule",
    "CountRule",
    "apply_validators",
    "emit_prompt_constraints",
]
