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
from .rrg import RrgHandler
from .im import ImHandler
from .spg_pl import DwHandler, HoHandler
from .spg_auto import CargoHandler, ApdHandler
from .spg_wh import WhHandler


_HANDLERS: dict[str, PolicyHandler] = {
    "GENERIC": GenericHandler(),
    "PAP": PapQuincyHandler(),
    "IMS": ImsHandler(),
    "RRG": RrgHandler(),
    "IM": ImHandler(),
    "DW": DwHandler(),
    "HO": HoHandler(),
    "CARGO": CargoHandler(),
    "APD": ApdHandler(),
    "WH": WhHandler(),
}


def get_handler(policy_type: str) -> PolicyHandler:
    """Return the handler registered for `policy_type`, or GenericHandler."""
    return _HANDLERS.get((policy_type or "").upper(), _HANDLERS["GENERIC"])


__all__ = ["PolicyHandler", "get_handler"]
