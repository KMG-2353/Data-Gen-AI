"""Profile model, registry, and runtime selection.

A ``Profile`` is the L1 overlay for one template: which L0 rules it inherits,
which it drops, which it overrides, and which L1 rules it adds. ``compose()``
resolves the effective rule set = L0 inherited + L1 added − dropped, with
overrides applied.

Selection reuses the existing ``detect_policy_type_from_headers()`` output as the
key (no new detection path). An unknown/GENERIC type composes to L0-only,
preserving generic behavior.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .l0_base import l0_rules
from .rules import Rule


@dataclass
class Profile:
    """L1 overlay for a single template."""

    policy_type: str
    drops: frozenset[str] = frozenset()
    overrides: dict[str, Rule] = field(default_factory=dict)
    added: list[Rule] = field(default_factory=list)

    def compose(self) -> list[Rule]:
        """Resolve the effective rule set for this template."""
        effective: list[Rule] = []
        for rule in l0_rules():
            if rule.id in self.drops:
                continue
            effective.append(self.overrides.get(rule.id, rule))
        effective.extend(self.added)
        return effective


def generic_profile() -> Profile:
    """No-op profile: inherit all L0, no L1. Used for unknown/GENERIC types."""
    return Profile(policy_type="GENERIC")


def select_profile(policy_type: str | None) -> Profile:
    """Map a detected policy type to its profile.

    Imports the per-template profile factories lazily to avoid an import cycle
    (profile modules import this module).
    """
    key = (policy_type or "").upper()
    if key == "RRG":
        from .profile_rrg import rrg_profile

        return rrg_profile()
    if key == "IMS":
        from .profile_ims import ims_profile

        return ims_profile()
    return generic_profile()


def compose_rules(policy_type: str | None) -> list[Rule]:
    """Convenience: select the profile and return its composed rule set."""
    return select_profile(policy_type).compose()
