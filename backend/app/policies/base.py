"""Abstract contract every policy handler implements.

The dispatcher (main.py) talks to handlers through this interface only.
Concrete handlers (pap_quincy, ims, generic, future mca) own their own
sheet-type detection, hardcoded rule blocks, deterministic pre-generation,
and post-processing.

Design principle: handlers must remain useful WITHOUT a rule book. The
rule book (special_instructions) narrows behavior further, but every
handler must produce valid output from hardcoded rules alone.
"""
from __future__ import annotations

from typing import Any, Protocol


class PolicyHandler(Protocol):
    """Protocol every policy-type handler must satisfy."""

    policy_type: str

    def detect_sheet_type(self, sheet_name: str) -> str:
        """Map a sheet name to a canonical sheet type for this policy.

        Returns "unknown" when the sheet doesn't match any known type.
        """
        ...

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        """Build per-sheet prompt augmentation + adjusted row count.

        Returns (adjusted_row_count, extra_prompt_text). Extra prompt is
        empty string when the handler has no sheet-specific rules.
        """
        ...

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
    ) -> list[dict[str, Any]] | None:
        """Optionally generate rows deterministically (skip LLM).

        Returns None when the handler has no deterministic path for this
        sheet, in which case the LLM path runs.
        """
        ...

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        """Enforce hard constraints deterministically after LLM generation."""
        ...
