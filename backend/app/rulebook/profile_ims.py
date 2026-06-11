"""IMS L1 profile + extracted rules.

Generalization proof (U7): two low-risk IMS property enums — Coinsurance and
Cause of Loss — are extracted into dual-emit rules. ``ims.py`` delegates the
matching post-process snap to them under ``config.RULEBOOK_ENABLED``, keeping all
IMS cross-sheet/LOB logic untouched. The canonical value lists live here (single
source) and the handler imports them back.

Import direction: handlers import from here; this module never imports handlers.
"""
from __future__ import annotations

import random

from .conditions import sheet_is
from .profiles import Profile
from .rules import EnumRule


# --- canonical IMS value lists (single source of truth) ----------------------

VALID_COINSURANCE = ["80", "90", "100"]
VALID_CAUSE_OF_LOSS = ["Basic", "Special", "Broad"]


_ON_PROPERTY = sheet_is("property")

IMS_COINSURANCE_RULE = EnumRule(
    id="ims.coinsurance.enum",
    field_keywords=("coinsurance",),
    allowed=tuple(VALID_COINSURANCE),
    snap=lambda v, allowed: random.choice(allowed),
    condition=_ON_PROPERTY,
    prompt_text="- Coinsurance: MUST be one of 80 / 90 / 100 [IMS Property]",
)

IMS_CAUSE_OF_LOSS_RULE = EnumRule(
    id="ims.cause_of_loss.enum",
    field_keywords=("cause of loss",),
    allowed=tuple(VALID_CAUSE_OF_LOSS),
    snap=lambda v, allowed: random.choice(allowed),
    condition=_ON_PROPERTY,
    prompt_text="- Cause of Loss: MUST be one of Basic / Special / Broad [IMS Property]",
)

IMS_EXTRACTED_RULES = [IMS_COINSURANCE_RULE, IMS_CAUSE_OF_LOSS_RULE]


def ims_profile() -> Profile:
    return Profile(
        policy_type="IMS",
        drops=frozenset(),
        overrides={},
        added=list(IMS_EXTRACTED_RULES),
    )
