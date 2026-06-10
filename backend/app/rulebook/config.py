"""Rulebook runtime configuration.

``RULEBOOK_ENABLED`` is the fallback flag (default on). When off, handlers take
their pure inline path — a fast escape if engine parity ever breaks. Read it as
``config.RULEBOOK_ENABLED`` (attribute access) so it can be toggled at runtime
and in tests.
"""
from __future__ import annotations

import os

RULEBOOK_ENABLED: bool = os.getenv("RULEBOOK_ENABLED", "1").strip().lower() not in (
    "0",
    "false",
    "no",
    "off",
)
