"""
Personal Auto Driver Assignment Logic
======================================
Deterministic rotation-based algorithm for generating driver-to-vehicle
assignment matrices per the Personal Auto Driver Assignment Implementation Plan.

Eligible vehicle types: Private Passenger, Classic
Eligible driver types:  Principal, Occasional

Pattern (per Quincy rule book):
  N = 1 eligible vehicle : ordinal = driver_rank (1, 2, 3 …)
  N > 1 eligible vehicles: Forward rotation → ordinal = (i + j) % N + 1
"""

from __future__ import annotations

import re
from typing import Any, Optional


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ELIGIBLE_VEH_KEYWORDS = ["Private Passenger", "Classic"]
ELIGIBLE_DRIVER_TYPES = {"Principal", "Occasional"}
_ORDINAL_SUFFIXES = {1: "st", 2: "nd", 3: "rd"}


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def is_eligible_vehicle(veh_type: str) -> bool:
    """Return True only for Private Passenger or Classic vehicle types."""
    if not veh_type or str(veh_type).strip().lower() in ("nan", "none", ""):
        return False
    veh_lower = str(veh_type).lower()
    return any(k.lower() in veh_lower for k in ELIGIBLE_VEH_KEYWORDS)


def is_eligible_driver(driver_type: str) -> bool:
    """Return True for Principal or Occasional driver types."""
    if not driver_type:
        return False
    return str(driver_type).strip() in ELIGIBLE_DRIVER_TYPES


def to_ordinal(n: int) -> str:
    """Convert an integer to its ordinal string (1 → '1st', 2 → '2nd', …)."""
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = _ORDINAL_SUFFIXES.get(n % 10, "th")
    return f"{n}{suffix}"


def get_transaction_group(test_case_no: str) -> Optional[str]:
    """Extract the transaction group (TS-XX-YY) from a full test case no (TS-XX-YY-ZZ)."""
    if not test_case_no:
        return None
    parts = str(test_case_no).strip().split("-")
    # TS-XX-YY → 3 parts; TS-XX-YY-ZZ → 4 parts
    return "-".join(parts[:3]) if len(parts) >= 3 else str(test_case_no).strip()


def _get_field(row: dict[str, Any], *candidates: str) -> str:
    """Get a field value from a row by case-insensitive substring match on key name."""
    for key, val in row.items():
        key_lower = key.lower()
        for candidate in candidates:
            if candidate.lower() in key_lower:
                return str(val).strip() if val is not None else ""
    return ""


def _get_tc_field(row: dict[str, Any]) -> str:
    """Get the Test Case No value, handling column header variants.

    Matches 'Test Case No', 'Test Case #', 'Test Case Number', 'Test Case'
    (bare form) but deliberately skips 'Test Case Details', 'Test Case Description',
    etc. to avoid returning the wrong column.
    """
    # Pass 1: exact match on known variants (fastest, most precise)
    _EXACT = {"test case no", "test case #", "test case number", "test case"}
    for key, val in row.items():
        if key.lower().strip() in _EXACT:
            return str(val).strip() if val is not None else ""
    # Pass 2: starts-with 'test case' but excludes columns like 'test case details'
    _EXCLUDE = ("detail", "description", "type", "status", "note")
    for key, val in row.items():
        kl = key.lower().strip()
        if kl.startswith("test case") and not any(exc in kl for exc in _EXCLUDE):
            return str(val).strip() if val is not None else ""
    return ""


# ---------------------------------------------------------------------------
# Core assignment matrix computation
# ---------------------------------------------------------------------------

def compute_assignment_matrix(
    drivers: list[dict[str, Any]],
    vehicles: list[dict[str, Any]],
    pattern: str = "backward",
) -> list[dict[str, Any]]:
    """
    Compute the assignment matrix for one transaction group.

    Parameters
    ----------
    drivers : list of driver dicts, each with at minimum:
        - 'test_case'      : full test case no (TS-XX-YY-ZZ)
        - 'transaction_type': e.g. 'New Business'
        - 'name'           : driver name
        - 'driver_type'    : 'Principal' | 'Occasional' | other
    vehicles : list of vehicle dicts, each with at minimum:
        - 'veh_type'       : vehicle type string used for eligibility check
    pattern : 'backward' (default) or 'forward'

    Returns
    -------
    List of row dicts, one per driver, each containing all original driver
    fields plus:
        - 'assignments' : list[str] — ordinal per eligible vehicle (length N)
        - 'n_vehicles'  : int       — number of eligible vehicles (N)
    """
    eligible_vehicles = [v for v in vehicles if is_eligible_vehicle(v.get("veh_type", ""))]
    N = len(eligible_vehicles)

    rows: list[dict[str, Any]] = []
    for i, driver in enumerate(drivers):
        if not is_eligible_driver(driver.get("driver_type", "")):
            assignments = [""] * N
        elif N == 0:
            assignments = []
        elif N == 1:
            # Single eligible vehicle: ordinal = driver rank (1st, 2nd, 3rd …)
            # (modulo formula always gives 1 when N=1, so use index directly)
            assignments = [to_ordinal(i + 1)]
        else:
            # Forward rotation (matches Quincy rule book table):
            # ordinal = (i + j) % N + 1
            assignments = [to_ordinal((i + j) % N + 1) for j in range(N)]

        rows.append({**driver, "assignments": assignments, "n_vehicles": N})

    return rows


# ---------------------------------------------------------------------------
# Sheet-level builder: driver_data + vehicle_data → assignment rows
# ---------------------------------------------------------------------------

def build_assignment_rows(
    driver_data: list[dict[str, Any]],
    vehicle_data: list[dict[str, Any]],
    policy_data: list[dict[str, Any]],
    assignment_headers: list[str],
    pattern: str = "backward",
) -> list[dict[str, Any]]:
    """
    Build the complete assignment sheet row list from generated driver,
    vehicle, and policy data.

    Parameters
    ----------
    driver_data         : rows from the generated Driver sheet
    vehicle_data        : rows from the generated Vehicle sheet
    policy_data         : rows from the generated Policy sheet
    assignment_headers  : column headers of the assignment sheet (from the Excel)
    pattern             : rotation pattern ('backward' or 'forward')

    Returns
    -------
    List of row dicts keyed exactly by assignment_headers values.
    """
    # ------------------------------------------------------------------
    # 1. Build transaction-type lookup keyed by TS-XX-YY
    # ------------------------------------------------------------------
    policy_lookup: dict[str, str] = {}
    for row in policy_data:
        tc = _get_tc_field(row)
        txn_type = _get_field(row, "transaction type")
        if tc:
            policy_lookup[tc] = txn_type

    # ------------------------------------------------------------------
    # 2. Group driver rows by transaction group (TS-XX-YY)
    # ------------------------------------------------------------------
    driver_groups: dict[str, list[dict[str, Any]]] = {}
    for dr in driver_data:
        tc_no = _get_tc_field(dr)
        grp = get_transaction_group(tc_no) or tc_no
        driver_groups.setdefault(grp, []).append(dr)

    # ------------------------------------------------------------------
    # 3. Group vehicle rows by transaction group (TS-XX-YY)
    #    Track each vehicle's 1-based position (from its TC suffix) so we
    #    can map ordinals to the correct Veh# column with no gaps.
    # ------------------------------------------------------------------
    vehicle_groups: dict[str, list[dict[str, Any]]] = {}
    for veh in vehicle_data:
        tc_no = _get_tc_field(veh)
        veh_type = _get_field(veh, "veh type", "vehicle type")
        grp = get_transaction_group(tc_no) or tc_no
        parts = str(tc_no).strip().split("-")
        try:
            veh_pos = int(parts[-1])  # 1-indexed position within test case
        except (ValueError, IndexError):
            veh_pos = len(vehicle_groups.get(grp, [])) + 1
        vehicle_groups.setdefault(grp, []).append({"veh_type": veh_type, "veh_pos": veh_pos})

    # Sort each group by vehicle position so ordinals are assigned in order
    for grp in vehicle_groups:
        vehicle_groups[grp].sort(key=lambda x: x.get("veh_pos", 0))

    # ------------------------------------------------------------------
    # 4. Determine maximum vehicle column index from headers
    # ------------------------------------------------------------------
    max_veh_col = 0
    for h in assignment_headers:
        m = re.match(r"veh\s*#\s*(\d+)", h.strip(), re.IGNORECASE)
        if m:
            max_veh_col = max(max_veh_col, int(m.group(1)))

    # ------------------------------------------------------------------
    # 5. For each transaction group, compute matrix and build rows
    # ------------------------------------------------------------------
    all_rows: list[dict[str, Any]] = []

    for grp, drivers_in_group in driver_groups.items():
        vehicles_in_group = vehicle_groups.get(grp, [])
        txn_type = policy_lookup.get(grp, "")

        # Build eligible-vehicle position map:
        #   veh_col_number (1-based) → ordinal_index (0-based into assignments list)
        # Only Private Passenger and Classic vehicles get ordinals; others get blank.
        eligible_col_map: dict[int, int] = {}
        ordinal_idx = 0
        for v in vehicles_in_group:
            pos = v.get("veh_pos", 0)
            if is_eligible_vehicle(v.get("veh_type", "")):
                eligible_col_map[pos] = ordinal_idx
                ordinal_idx += 1

        # Build typed input for compute_assignment_matrix
        drivers_input: list[dict[str, Any]] = []
        for dr in drivers_in_group:
            drivers_input.append(
                {
                    "test_case": _get_tc_field(dr),
                    "transaction_type": txn_type,
                    "name": _get_field(dr, "name"),
                    "driver_type": _get_field(dr, "driver type"),
                }
            )

        matrix_rows = compute_assignment_matrix(drivers_input, vehicles_in_group, pattern)

        for matrix_row in matrix_rows:
            assignments: list[str] = matrix_row.get("assignments", [])
            row_dict: dict[str, Any] = {}

            for header in assignment_headers:
                h_strip = header.strip()
                h_lower = h_strip.lower()

                # Vehicle column  (Veh #1, Veh #2, …)
                veh_match = re.match(r"veh\s*#\s*(\d+)", h_strip, re.IGNORECASE)
                if veh_match:
                    veh_col = int(veh_match.group(1))  # 1-based column number
                    if veh_col in eligible_col_map:
                        ord_idx = eligible_col_map[veh_col]
                        row_dict[header] = assignments[ord_idx] if ord_idx < len(assignments) else ""
                    else:
                        row_dict[header] = ""  # ineligible vehicle column → blank
                    continue

                # Test Case number column
                if "test case" in h_lower and "detail" not in h_lower:
                    row_dict[header] = matrix_row.get("test_case", "")
                    continue

                # Transaction Type column
                if "transaction type" in h_lower:
                    row_dict[header] = matrix_row.get("transaction_type", "")
                    continue

                # Name column (but not "Driver Type")
                if "name" in h_lower and "driver" not in h_lower:
                    row_dict[header] = matrix_row.get("name", "")
                    continue

                # Driver Type column
                if "driver type" in h_lower:
                    row_dict[header] = matrix_row.get("driver_type", "")
                    continue

                # Any other column → blank
                row_dict[header] = ""

            all_rows.append(row_dict)

    return all_rows


# ---------------------------------------------------------------------------
# Prompt helper: build rotation instructions for LLM fallback
# ---------------------------------------------------------------------------

def build_assignment_prompt_instructions(
    policy_structure: list[dict[str, Any]],
    driver_data: list[dict[str, Any]] | None,
    vehicle_data: list[dict[str, Any]] | None,
    pattern: str = "backward",
) -> str:
    """
    Build detailed row-by-row assignment instructions with the full rotation
    matrix pre-computed. Used as the LLM fallback when deterministic generation
    is not triggered.
    """
    if not policy_structure:
        return ""

    lines: list[str] = []
    row_num = 0

    for ps in policy_structure:
        base_no = ps["test_case_no"]
        txn_type = ps["transaction_type"]
        driver_count = ps["driver_count"]
        vehicle_count = ps["vehicle_count"]

        # Try to get eligible vehicle count from actual vehicle data
        grp = base_no  # TS-XX-YY
        N = vehicle_count  # fallback: use total count from policy
        if vehicle_data:
            eligible = [
                v for v in vehicle_data
                if get_transaction_group(_get_tc_field(v)) == grp
                and is_eligible_vehicle(_get_field(v, "veh type", "vehicle type"))
            ]
            if eligible:
                N = len(eligible)

        for i in range(driver_count):
            row_num += 1
            expanded_no = f"{base_no}-{i + 1:02d}"

            if N == 0:
                veh_parts: list[str] = []
            elif N == 1:
                veh_parts = [f'Veh #1 = "{to_ordinal(i + 1)}"']
            else:
                veh_parts = [f'Veh #{j + 1} = "{to_ordinal((i + j) % N + 1)}"' for j in range(N)]

            lines.append(
                f'  Row {row_num}: Test Case No = "{expanded_no}", '
                f'Transaction Type = "{txn_type}", '
                + (", ".join(veh_parts) if veh_parts else "no vehicle columns")
            )

    return (
        "EXACT ROW MAPPING FOR ASSIGNMENT SHEET (full rotation matrix pre-computed):\n"
        f"You MUST generate exactly {row_num} rows with these exact values:\n"
        + "\n".join(lines)
        + "\n\nDo NOT deviate from this mapping."
        + "\nThe Name and Driver Type columns MUST match the corresponding Driver sheet rows exactly."
        + "\nLeave any vehicle columns beyond the ones listed here blank."
    )
