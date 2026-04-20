"""Generic fallback handler — preserves baseline main-branch behavior.

Used when the uploaded filename does not match any known policy type
(not IMS*, not PAP*). Applies only the universal effective/expiration
date policy that already exists in production main.
"""
from __future__ import annotations

from typing import Any


class GenericHandler:
    policy_type = "GENERIC"

    def detect_sheet_type(self, sheet_name: str) -> str:
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
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        # Date enforcement already runs inside generate_test_data for the
        # generic path; nothing further to do here.
        return rows
