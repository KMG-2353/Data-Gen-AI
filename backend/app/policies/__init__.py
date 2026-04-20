"""Policy-type handler registry.

Each handler encapsulates client-specific hardcoded rules, sheet-context
building, deterministic pre-generation, and post-processing for a given
insurance policy type (PAP Quincy, IMS, MCA, …).

Handlers are selected by `detect_policy_type(filename)` in llm_service.
Unknown filenames fall back to the GenericHandler, which preserves the
baseline "main" behavior (no client-specific logic).
"""
from __future__ import annotations

from .base import PolicyHandler
from .generic import GenericHandler
from .pap_quincy import PapQuincyHandler
from .ims import ImsHandler


_HANDLERS: dict[str, PolicyHandler] = {
    "GENERIC": GenericHandler(),
    "PAP": PapQuincyHandler(),
    "IMS": ImsHandler(),
}


def get_handler(policy_type: str) -> PolicyHandler:
    """Return the handler registered for `policy_type`, or GenericHandler."""
    return _HANDLERS.get((policy_type or "").upper(), _HANDLERS["GENERIC"])


__all__ = ["PolicyHandler", "get_handler"]
