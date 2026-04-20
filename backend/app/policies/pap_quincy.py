"""Personal Auto Policy (Quincy) handler.

Phase-1 scaffold: delegates to the existing llm_service + assignment_logic
functions so behavior on claude/tender-vaughan is preserved byte-for-byte.
Phase 2 will migrate the rule blocks, VIN/infraction tables, and sheet
enforcers out of llm_service into this module so the file becomes the
single source of truth for Quincy PAP.
"""
from __future__ import annotations

from typing import Any


class PapQuincyHandler:
    policy_type = "PAP"

    def detect_sheet_type(self, sheet_name: str) -> str:
        from app.llm_service import detect_sheet_type as _detect
        return _detect(sheet_name)

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        from app.llm_service import build_insurance_context
        return build_insurance_context(
            sheet_name=sheet_name,
            policy_data=policy_data,
            driver_data=driver_data,
            original_row_count=original_row_count,
            policy_type="PAP",
        )

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
    ) -> list[dict[str, Any]] | None:
        sheet_type = self.detect_sheet_type(sheet_name)
        if sheet_type != "assignment":
            return None
        if not (driver_data and vehicle_data and policy_data):
            return None
        from app.assignment_logic import build_assignment_rows
        return build_assignment_rows(
            driver_data=driver_data,
            vehicle_data=vehicle_data,
            policy_data=policy_data,
            assignment_headers=unique_headers,
        )

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
    ) -> list[dict[str, Any]]:
        from app.llm_service import _enforce_effective_expiration_date_range
        return _enforce_effective_expiration_date_range(rows, special_instruction)
