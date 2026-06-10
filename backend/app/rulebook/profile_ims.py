"""IMS L1 profile.

Skeleton in U3. U7 fills in a couple of extracted IMS enum rules (e.g. number of
stories, coinsurance, cause of loss) as additions, proving the engine
generalizes beyond RRG. IMS cross-sheet/LOB logic stays in the handler.
"""
from __future__ import annotations

from .profiles import Profile


def ims_profile() -> Profile:
    return Profile(
        policy_type="IMS",
        drops=frozenset(),
        overrides={},
        added=[],
    )
