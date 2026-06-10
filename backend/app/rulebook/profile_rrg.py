"""RRG L1 profile.

Skeleton in U3: drops the L0 SSN rule (RRG has no SSN field). U4 fills in the
extracted enum/format rules (State, Org Type, ZIP, Contact Number) as overrides
and additions, with parity against the current handler.
"""
from __future__ import annotations

from .profiles import Profile


def rrg_profile() -> Profile:
    return Profile(
        policy_type="RRG",
        drops=frozenset({"l0.ssn"}),  # RRG workbooks have no SSN field
        overrides={},
        added=[],
    )
