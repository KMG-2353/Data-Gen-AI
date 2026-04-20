"""IMS (commercial insurance) handler.

Phase-1 stub. The real port of IMS rules from origin/IMS-new-business
(ParsedInstructions, LOB filtering, sheet enforcers, cross-sheet
consistency) lands in Phase 3. For now the handler behaves like the
generic fallback so IMS file uploads do not crash on this branch.
"""
from __future__ import annotations

from typing import Any


class ImsHandler:
    policy_type = "IMS"

    def detect_sheet_type(self, sheet_name: str) -> str:
        # Phase 3 will add IMS-specific sheet-name detection (IMS screen,
        # NetRate Policy, NetRate Property, etc.).
        return "unknown"

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        return original_row_count, ""

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
    ) -> list[dict[str, Any]] | None:
        return None

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
    ) -> list[dict[str, Any]]:
        from app.llm_service import _enforce_effective_expiration_date_range
        return _enforce_effective_expiration_date_range(rows, special_instruction)
