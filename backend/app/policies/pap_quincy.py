"""Personal Auto Policy (Quincy) handler.

Owns all Quincy-PAP-specific logic: sheet-name detection, the six valid
driver/vehicle combinations, policy-structure parsing, per-sheet rule
blocks injected into the prompt, the VIN and infraction reference tables,
deterministic assignment-sheet generation, and row-count expansion for
driver/vehicle/assignment/infraction/summary sheets.

The core rules here are HARDCODED so PAP generation works without any
rule book present. Special instructions from the user narrow behavior
further but are never required.
"""
from __future__ import annotations

import re
from typing import Any


# ---------------------------------------------------------------------------
# Valid driver/vehicle combinations — hardcoded constraint.
# The user's rule book allows ONLY these six combinations. Any other
# combination in Test Case Details is clamped to "1 Driver & 1 Vehicle".
# ---------------------------------------------------------------------------

VALID_DRIVER_VEHICLE_COMBINATIONS: set[tuple[int, int]] = {
    (1, 1),
    (2, 1),
    (1, 2),
    (2, 2),
    (3, 1),
    (4, 6),
}

# Number of Principal drivers for each valid combination.
# Rule: n_principal = number of eligible (PP / Classic) vehicles, up to driver_count.
# For (2,2) both vehicles are PP/Classic → both drivers are Principal.
# For (4,6) all six vehicles are eligible → all 4 drivers are Principal.
_N_PRINCIPAL: dict[tuple[int, int], int] = {
    (1, 1): 1,
    (2, 1): 1,  # only 1 eligible vehicle → 1 Principal, 1 Occasional
    (1, 2): 1,  # single driver must be Principal
    (2, 2): 2,  # both vehicles PP/Classic → both Principal
    (3, 1): 1,  # only 1 eligible vehicle → 1 Principal, 2 Occasional
    (4, 6): 4,  # 6 eligible vehicles → all 4 drivers Principal
}


def _get_n_principal(driver_count: int, vehicle_count: int) -> int:
    """Return how many of the drivers (from driver 1 onward) should be Principal."""
    return _N_PRINCIPAL.get((driver_count, vehicle_count), 1)


# ---------------------------------------------------------------------------
# Sheet-type detection
# ---------------------------------------------------------------------------

def detect_sheet_type(sheet_name: str) -> str:
    """Map a PAP workbook sheet name to a canonical sheet type."""
    name = sheet_name.lower().strip()
    if "infraction" in name:
        return "infraction"
    if "assignment" in name:
        return "assignment"
    if "policy info" in name or "policy_info" in name:
        return "policy_info"
    if "summary" in name:
        return "summary"
    if "questions" in name or "remarks" in name:
        return "questions_remarks"
    if "driver" in name:
        return "driver"
    # "Vehical" is a known typo in the Quincy template — treat as vehicle.
    if "vehicle" in name or "vehical" in name:
        return "vehicle"
    if "policy" in name:
        return "policy"
    return "unknown"


# ---------------------------------------------------------------------------
# Policy-structure parsing (drives row expansion for downstream sheets)
# ---------------------------------------------------------------------------

def parse_policy_structure(policy_data: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Parse Policy sheet rows to extract test-case structure.

    Each returned dict carries driver_count, vehicle_count, transaction_type,
    insured_name, and Test Case No for one policy row. Invalid combinations
    are clamped to (1, 1) with a warning.
    """
    structure = []
    for row in policy_data:
        info: dict[str, Any] = {
            "test_case_no": "",
            "test_case_details": "",
            "transaction_type": "",
            "driver_change": "",
            "vehicle_change": "",
            "assignment_change": "",
            "insured_name": "",
            "driver_count": 1,
            "vehicle_count": 1,
        }

        _TC_EXACT = {"test case no", "test case #", "test case number", "test case"}
        _TC_EXCL = ("detail", "description", "type", "status", "note")
        for key, value in row.items():
            kl = key.lower().strip()
            val = str(value).strip() if value else ""
            # Test Case No — handle variants like bare 'Test Case'
            if (
                kl in _TC_EXACT
                or (kl.startswith("test case") and not any(e in kl for e in _TC_EXCL))
            ):
                if not info["test_case_no"]:
                    info["test_case_no"] = val
            elif "test case detail" in kl:
                info["test_case_details"] = val
            elif "transaction type" in kl:
                info["transaction_type"] = val
            elif "driver change" in kl:
                info["driver_change"] = val
            elif "vehicle change" in kl:
                info["vehicle_change"] = val
            elif "assignment change" in kl:
                info["assignment_change"] = val
            elif "insured name" in kl:
                info["insured_name"] = val

        m = re.search(r"(\d+)\s*Driver.*?(\d+)\s*Vehicle", info["test_case_details"], re.IGNORECASE)
        if m:
            d = int(m.group(1))
            v = int(m.group(2))
            if (d, v) in VALID_DRIVER_VEHICLE_COMBINATIONS:
                info["driver_count"] = d
                info["vehicle_count"] = v
            else:
                print(
                    f"WARNING: Invalid driver/vehicle combination ({d} Driver & {v} Vehicle) "
                    f"in test case '{info['test_case_no']}'. Defaulting to 1 Driver & 1 Vehicle."
                )
                info["driver_count"] = 1
                info["vehicle_count"] = 1

        structure.append(info)

    return structure


def calculate_expanded_row_count(
    sheet_type: str,
    policy_structure: list[dict[str, Any]],
    driver_data: list[dict[str, Any]] | None = None,
) -> int | None:
    """Compute row count for sheets driven by Policy-sheet structure."""
    if not policy_structure:
        return None

    if sheet_type == "driver":
        return sum(ps["driver_count"] for ps in policy_structure)
    elif sheet_type == "vehicle":
        return sum(ps["vehicle_count"] for ps in policy_structure)
    elif sheet_type == "assignment":
        return sum(ps["driver_count"] for ps in policy_structure)
    elif sheet_type == "infraction":
        # Every driver gets a row in the infraction sheet (infraction fields blank
        # for those without an infraction). Row count = total driver count.
        if driver_data:
            return len(driver_data)
        return sum(ps["driver_count"] for ps in policy_structure)
    elif sheet_type in ("summary", "policy_info", "questions_remarks"):
        return len(policy_structure)

    return None


def build_row_mapping_instructions(
    sheet_type: str,
    policy_structure: list[dict[str, Any]],
    driver_data: list[dict[str, Any]] | None = None,
) -> str:
    """Pre-compute exact row-by-row Test Case No + transaction instructions."""
    if not policy_structure:
        return ""

    lines: list[str] = []

    if sheet_type == "driver":
        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            insured = ps["insured_name"]
            is_new_business = "new business" in txn_type.lower()
            n_principal = _get_n_principal(ps["driver_count"], ps["vehicle_count"])

            for d in range(1, ps["driver_count"] + 1):
                row_num += 1
                expanded_no = f"{base_no}-{d:02d}"
                txn = "Add" if is_new_business else "Edit"
                # First n_principal drivers are Principal; rest are Occasional
                driver_type = "Principal" if d <= n_principal else "Occasional"
                name_hint = (
                    f', Name = "{insured}" (same as Insured Name), Driver Type = "{driver_type}"'
                    if d == 1 and insured
                    else f", Name = random unique name, Driver Type = \"{driver_type}\""
                )
                lines.append(
                    f"  Row {row_num}: Test Case No = \"{expanded_no}\", "
                    f"Transaction Type = \"{txn_type}\", Transaction = \"{txn}\"{name_hint}"
                )

        return (
            "EXACT ROW MAPPING FOR DRIVER SHEET:\n"
            f"You MUST generate exactly {row_num} rows with these exact Test Case No values and Driver Types:\n"
            + "\n".join(lines)
            + "\n\nDo NOT deviate from this mapping. Each row above is one driver record."
            + "\nIMPORTANT: Driver Type values above are MANDATORY — do not change them."
        )

    elif sheet_type == "vehicle":
        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            is_new_business = "new business" in txn_type.lower()

            for v in range(1, ps["vehicle_count"] + 1):
                row_num += 1
                expanded_no = f"{base_no}-{v:02d}"
                txn = "Add" if is_new_business else "Edit"
                lines.append(
                    f"  Row {row_num}: Test Case No = \"{expanded_no}\", "
                    f"Transaction Type = \"{txn_type}\", Transaction = \"{txn}\""
                )

        return (
            "EXACT ROW MAPPING FOR VEHICLE SHEET:\n"
            f"You MUST generate exactly {row_num} rows with these exact Test Case No values:\n"
            + "\n".join(lines)
            + "\n\nDo NOT deviate from this mapping. Each row above is one vehicle record."
            + "\nEach vehicle in the same test case MUST have a UNIQUE VIN - no duplicate VINs allowed."
        )

    elif sheet_type == "assignment":
        from app.assignment_logic import to_ordinal as _to_ordinal

        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            driver_count = ps["driver_count"]
            N = ps["vehicle_count"]

            for i in range(driver_count):
                row_num += 1
                expanded_no = f"{base_no}-{i + 1:02d}"
                if N == 0:
                    veh_parts: list[str] = []
                else:
                    veh_parts = [
                        f'Veh #{j + 1} = "{_to_ordinal((j - i) % N + 1)}"'
                        for j in range(N)
                    ]

                lines.append(
                    f'  Row {row_num}: Test Case No = "{expanded_no}", '
                    f'Transaction Type = "{txn_type}", '
                    + (", ".join(veh_parts) if veh_parts else "no vehicle columns")
                )

        return (
            "EXACT ROW MAPPING FOR ASSIGNMENT SHEET (backward rotation formula applied):\n"
            f"You MUST generate exactly {row_num} rows with these exact values:\n"
            + "\n".join(lines)
            + "\n\nDo NOT deviate from this mapping."
            + "\nThe Name and Driver Type columns MUST match the corresponding Driver sheet rows exactly."
            + "\nLeave any vehicle columns beyond the ones listed here blank."
        )

    elif sheet_type == "summary":
        row_num = 0
        for ps in policy_structure:
            row_num += 1
            lines.append(
                f"  Row {row_num}: Test Case No = \"{ps['test_case_no']}\", "
                f"Transactions = \"{ps['transaction_type']}\""
            )
        return (
            "EXACT ROW MAPPING FOR SUMMARY SHEET:\n"
            f"You MUST generate exactly {row_num} rows:\n"
            + "\n".join(lines)
            + "\n\nExecute Transaction = \"No\", Hold Transaction = \"No\"."
            + "\nAll other fields (Wins Reference Quote, Python Reference Quote, etc.) should be BLANK."
        )

    elif sheet_type == "questions_remarks":
        row_num = 0
        for ps in policy_structure:
            row_num += 1
            lines.append(
                f"  Row {row_num}: Test Case No = \"{ps['test_case_no']}\""
            )
        return (
            "EXACT ROW MAPPING FOR QUESTIONS/REMARKS SHEET:\n"
            f"You MUST generate exactly {row_num} rows:\n"
            + "\n".join(lines)
            + "\n\nProducer Name: blank for ME and RI states. CT state only: NANCY MENDIZABAL or MICHAEL PRENDERGAST."
            + "\nAll other fields: blank."
        )

    elif sheet_type == "infraction":
        if not driver_data:
            return ""
        row_num = 0
        _TC_EXACT = {"test case no", "test case #", "test case number", "test case"}
        _TC_EXCLUDE = ("detail", "description", "type", "status", "note")
        for drv in driver_data:
            has_infraction = False
            test_case_no = ""
            txn_type = ""
            for key, val in drv.items():
                kl = key.lower().strip()
                if "add infraction" in kl and str(val).strip().lower() == "yes":
                    has_infraction = True
                if not test_case_no:
                    if kl in _TC_EXACT or (
                        kl.startswith("test case")
                        and not any(exc in kl for exc in _TC_EXCLUDE)
                    ):
                        test_case_no = str(val).strip() if val else ""
                if "transaction type" in kl:
                    txn_type = str(val).strip() if val else ""

            row_num += 1
            is_new_business = "new business" in txn_type.lower()
            txn = "Add" if is_new_business else "Edit"
            infraction_note = (
                "FILL infraction fields (Infraction Type, Violation Date, SDIP Points, etc.)"
                if has_infraction
                else "Leave ALL infraction fields BLANK (no infraction for this driver)"
            )
            lines.append(
                f"  Row {row_num}: Test Case No = \"{test_case_no}\", "
                f"Transaction Type = \"{txn_type}\", Transaction = \"{txn}\" → {infraction_note}"
            )

        return (
            "EXACT ROW MAPPING FOR INFRACTION SHEET:\n"
            f"You MUST generate exactly {row_num} rows — one for EVERY driver (even those without infractions):\n"
            + "\n".join(lines)
            + "\n\nFor rows marked 'Leave ALL infraction fields BLANK': only Test Case No and Transaction Type are filled."
            + "\nFor rows marked 'FILL infraction fields': populate Infraction Type, Violation Date, SDIP Points, etc."
        )

    return ""


# ---------------------------------------------------------------------------
# Hardcoded PAP rule blocks (injected into the LLM prompt).
# These apply regardless of whether a rule book is provided — they encode
# the non-negotiable Quincy-PAP constraints.
# ---------------------------------------------------------------------------

PAP_CORE_RULES = """
PERSONAL AUTO POLICY (PAP) – CORE RULES (always enforced):
- Policy structure: every test case starts with exactly one "New Business" transaction.
- Test Case Details MUST be one of exactly 6 valid combinations:
    "1 Driver & 1 Vehicle", "2 Driver & 1 Vehicle", "1 Driver & 2 Vehicle",
    "2 Driver & 2 Vehicle", "3 Driver & 1 Vehicle", "4 Driver & 6 Vehicle"
- Driver count and vehicle count defined by Test Case Details must be reproduced EXACTLY across all sheets.
- At least one vehicle per test case must be "Private Passenger" or "Classic-Refer to CO (Stated Amount)".
- Date format for ALL date fields: MMDDYYYY (8 digits, NO slashes, NO dashes, NO separators of any kind).
- Effective Date: within ±60 days of today for non-Renewal transactions.
  For Renewal: 300+ days in the past.
- Expiration Date: exactly one year after Effective Date (same month/day, next year).
- Transaction Type value in Driver, Vehicle, and Assignment sheets must match the Policy sheet exactly.
- No blank Test Case No fields — every row in every sheet must have a valid Test Case No.
- Assignment sheet: each driver gets one row; ordinals fill only eligible-vehicle columns (no gaps).
"""

POLICY_SHEET_RULES = """
POLICY SHEET RULES:
- Test Case No format: TS-XX-XX (first XX = test case number, second XX = transaction number within test case)
- Transaction Types (must follow this sequence logic):
    New Business = first transaction for every test case
    Endorsement = after New Business
    Flat Cancel = cancellation date equals effective date of New Business
    Mid Term Cancel = cancellation done after the effective date
    Add From Cancel = only after Flat Cancel, same date as cancellation
    Reinstatement = only after Mid Term Cancel, same date as cancellation
    Retroactive = at least two endorsements must have been completed after New Business
    Renewal = at least 300 days must have elapsed from the New Business effective date
- Policy Change: Yes (default for New Business). For Endorsement/Renewal/Retroactive/Add From Cancel: Yes if any policy field changed, No otherwise. Always No for Cancellation/Reinstatement.
- Driver Change: Yes (default for New Business). For Endorsement/Renewal/Retroactive/Add From Cancel: Yes if any driver field changed, No otherwise. Always No for Cancellation/Reinstatement.
- Vehicle Change: Yes (default for New Business). For Endorsement/Renewal/Retroactive/Add From Cancel: Yes if any vehicle field changed, No otherwise. Always No for Cancellation/Reinstatement.
- Assignment Change: Yes (default for New Business). For Endorsement/Renewal/Retroactive/Add From Cancel: Yes if any assignment field changed, No otherwise. Always No for Cancellation/Reinstatement.
- Test Case Details: ONLY these 6 combinations are valid:
    "1 Driver & 1 Vehicle", "2 Driver & 1 Vehicle", "1 Driver & 2 Vehicle",
    "2 Driver & 2 Vehicle", "3 Driver & 1 Vehicle", "4 Driver & 6 Vehicle"
  This field is the single source of truth for driver and vehicle counts across ALL sheets.
- Effective Date: MMDDYYYY format (8 digits, NO slashes, NO dashes). Current date ±60 days.
  CRITICAL: All transactions within the same test case (TS-XX) must have the SAME effective date.
  Exception: Renewal transaction gets its own effective date (300+ days in the past).
- Endorsement Date: MMDDYYYY format (NO slashes).
  For New Business: Endorsement Date = Effective Date (must be identical).
  For other transactions: >= New Business date and < 50 days in future.
  For Cancellation: > New Business date. For Reinstatement: same as Cancellation date.
  For Retroactive: same as the previous Endorsement date.
- Rating state = ME (default unless overridden):
    Company Code = 0020, Agent = 05899 or 05297, Loss Free = 1-9 (any number)
    Type, UMPD, CT PIP-BRB = Blank (these are CT-only fields)
- Rating state = RI: Company Code = 0010, Agent = 00130
- Rating state = CT: Company Code = 0010, Agent = 00130
- Phone: 10 digits, format "1111111111" (no dashes, no spaces, no parentheses)
- Email: ends with @test.com or @gmail.com
- Liability CSL and Liability BI are MUTUALLY EXCLUSIVE — select one or the other, never both.
  CSL options: 125,000 / 200,000 / 300,000 / 500,000 / 1,000,000
  BI options: 50/100 / 100/300 / 250/500 / 500/1,000
- UM/UIM CSL: value must be <= Liability CSL. Select from: 100,000 / 125,000 / 200,000 / 300,000 / 500,000 / 1,000,000
- UM/UIM BI: value must be <= Liability BI. Only if Liability BI selected.
- PD (25,000 / 50,000 / 100,000 / 250,000): only fill if Liability BI is selected.
- Med Pay: 2,000 / 5,000 / 10,000 / 25,000 / 50,000
- Payment Plan: one value for the entire test case.
  Options: "Direct Bill - 2 Pay", "Direct Bill - 4 Pay", "Direct Bill - 9 Pay",
  "EFT - 10 Pay Pick a Day", "Direct Bill - One Pay 5% Discount", "EFT - 12 Month Installment Plan"
- Monthly Due Day: fill ONLY if Payment Plan = "EFT - 10 Pay Pick a Day". Value: 1-30.
- WINS Quote Number, Python Quote Number, Wins Policy Number, Python Policy Number,
  Insurance Score, Client ID, EFT: always Blank.
- Less than 3 years at current address = Yes: fill Previous Street/City/State/ZIP. No = leave blank.
- Retroactive options: Endorse=option 1, Final Cancel=option 3, Reinstate=option 4
- Type (Standard/Conversion): CT state only. Blank for ME and RI.
- UMPD (25,000 / 50,000): CT and RI states only. Blank for ME.
- CT PIP-BRB (Yes/blank): CT state only. No (blank) if Med Pay is selected.
- Group (Yes/blank), Corporate Car (Yes/blank): any state.
- Credit Company: RI state only. Options: Quincy Mutual GRP / Narragansett / Andover.
  If Quincy Mutual GRP: fill Policy# and Policy No. If Narragansett/Andover: leave blank.
- Account (Yes/blank): if blank, Policy# and Policy No must also be blank.
- Policy# options for ME: AUT / DWL / HOM / SON / PIM
- Policy No: 806993 or 806995 (fill only if Policy# is selected)
- Loss Free: ME state only, value 1-9
- Producer Name: CT state only — NANCY MENDIZABAL or MICHAEL PRENDERGAST
"""

DRIVER_SHEET_RULES = """
DRIVER SHEET RULES:
- Test Case No format: TS-XX-XX-XX (last XX = driver number: 01, 02, 03 ...)
- Transaction: "Add" for New Business. "Edit" for Endorsement/Renewal/Retroactive/Add From Cancel. "Delete" only if driver count > 1.
- First driver's Name MUST equal Insured Name from Policy sheet. Second driver onward = random two-word name.
- Date of Birth: MMDDYYYY format (NO slashes). Year range: 1930 to (current year - 16), i.e. 1930–2008.
- Age: calculated from Date of Birth to today.
- Gender: Male or Female.
- Driver Type:
    The number of Principal drivers equals the number of eligible vehicles (Private Passenger or Classic).
    Combination rules (STRICT — use the pre-computed Driver Type from the row mapping above):
      "1 Driver & 1 Vehicle" → 1 Principal
      "2 Driver & 1 Vehicle" → 1 Principal, 1 Occasional
      "1 Driver & 2 Vehicle" → 1 Principal (only 1 driver)
      "2 Driver & 2 Vehicle" → 2 Principal (both eligible vehicles → both Principal)
      "3 Driver & 1 Vehicle" → 1 Principal, 2 Occasional
      "4 Driver & 6 Vehicle" → 4 Principal (all 4 drivers are Principal)
    Occasional driver type: valid only when there are multiple drivers and vehicles.
    "Occasional" is NOT valid when all vehicles are eligible and drivers ≤ eligible vehicles.
- Marital Status: Single / Married / Divorced / Widowed.
- Relationship To Insured: first driver = "Insured". Others = Father/Mother/Spouse/Sibling/Son/Daughter/Other.
- License Date: Date of Birth + 18 years, MMDDYYYY format (NO slashes).
- Lic State: CT, RI, or ME. Preferably same as Rating State.
- License #: CT state = exactly 9 digits. ME state = exactly 7 digits. RI state = 7 to 9 digits.
- Occupation: CONSTRUCTION AND EXTRACTION / HEALTHCARE SUPPORT / STUDENT / MANAGEMENT /
  RETIRED / LEGAL / UNEMPLOYED (pick any one).
- Driver Training: Yes or Blank. Yes ONLY if driver age <= 20.
  If Yes: Driver Training Completion Date = License Date.
- Mature Credit: Yes or Blank. If Yes: Mature Credit Completion Date = License Date.
- Good Student: Yes or Blank. Yes ONLY if driver age <= 25.
- Operator student 100+ miles from home: Yes or Blank. Yes ONLY if driver age <= 20.
- MVR Re-Order: always NO.
- Claims Report Re-Order: always NO.
- Violations in last 3 years: always Blank.
- Accidents/Claims in last 3 years: always Blank.
- Add Infraction: Yes or Blank. If Yes, Infraction sheet must have a record for this driver.
- Remove Infraction: Yes or Blank. If Yes, Remove Infraction SDIP is mandatory.
  Only possible if an infraction was previously added for this driver.
"""

VEHICLE_SHEET_RULES = """
VEHICLE SHEET RULES:
- Test Case No format: TS-XX-XX-XX (last XX = vehicle number: 01, 02, 03 ...)
- Transaction: "Add" for New Business. "Edit" for Endorsement/Renewal/Retroactive/Add From Cancel. "Delete" only if vehicle count > 1.
- Territory: always Blank.
- Mature Driver#: always Blank.
- Mature Driver: always Blank.
- Garaging Street: realistic street address (street only, no city/state/ZIP).
- City, State, Zipcode: consistent real location; State must be the same as Rating State (ME, RI, or CT).
- STS: always Blank.
- Vehicle Types: Antique-Refer to CO (Stated Amount), Classic-Refer to CO (Stated Amount),
  Motorhome (Cost New), Private Passenger, Recreational Trailer (Cost New), Utility Trailer (Stated Amt).
  If only 1 vehicle: MUST be Private Passenger or Classic-Refer to CO (Stated Amount).
  If multiple vehicles: at least one MUST be Private Passenger or Classic.
- Veh Use: Pleasure / Business / Commute / Farming.
  Motorhome, Utility Trailer, Antique → Pleasure ONLY.
  If Commute: Miles One-way field is mandatory (01 to 02 / 03 to 14 / 15+).
- VIN, Model Year, Make/Model, Style: use EXACT rows from VIN table provided. NO duplicates within same test case.
- Vehicle Make column: always Blank.
- Make/Model: copy exactly as shown in the VIN table. Do NOT split the data.
- Pass/Rest: "00% No Passive Restraint" / "20% Seat Belt - Driver only" / "20% Air Bag - Driver only" /
  "30% Seat Belts- Both Sides" / "30% Air Bags - Both Sides" / "30% Not classified Both Sides"
- Anti-Theft: "00% No Anti-Theft Credit" / "05% Alarm Only" / "05% Active Disabling Device" /
  "15% Passive Disabling Device" / "All Other"
- Anti-Lock: always Yes.
- Comp Ded: blank / 50 / 100 / 200 / 250 / 500 / 1,000 (use 50 very rarely).
- Full Glass: Yes ONLY if vehicle type is Private Passenger AND Comp Ded >= 200. Otherwise blank.
- Coll Ded: blank / 50 / 100 / 200 / 250 / 500 / 1,000 (use 50 very rarely).
- Sub Trans (30/900, 40/1,000, 50/1,500 or blank): only if Comp Ded >= 200.
- Towing (25, 50, 75, 100 or blank): only if Comp Ded >= 200.
- Excess Electronic (1,500–5,000): only if Comp Ded >= 200.
- Corporate Car Discount: Yes only if Comp Ded >= 100, else blank.
- Loan/Lease Coverage: always blank.
- RI UMPD: blank for CT and ME states.
- Joint Ownership: Yes only if Comp Ded >= 100, else blank.
- Enhancement Endorsement: Yes only if Coll Ded >= 200, else blank.
  Enhancement Endorsement and Trip Interrupt 600 CANNOT both be Yes — only one at a time.
- Delete Liability: Yes or blank.
- Trip Interrupt 600: Yes only if Comp Ded >= 200, else blank.
  Enhancement Endorsement and Trip Interrupt 600 CANNOT both be Yes — only one at a time.
- Cost New: random 1,000–100,000 (no $ symbol). ONLY for: Private Passenger, Motorhome, Recreational Trailer.
- Stated Amt: random 1,000–100,000 (no $ symbol). ONLY for: Classic, Antique, Utility Trailer.
- Customized Amt: thousands format only (2,000 / 5,000 / 10,000 / 20,000 / 35,000 / 45,000 / 60,000). No $ symbol.
- Suspend Liability: must equal Delete Liability (both Yes or both blank).
"""

ASSIGNMENT_SHEET_RULES = """
ASSIGNMENT SHEET RULES:
- Test Case No format: TS-XX-XX-XX (last XX = driver number, matches Driver sheet)
- Name: must match exactly the driver's Name from Driver sheet.
- Driver Type: must match exactly the driver's Driver Type from Driver sheet.
- Only vehicles of type "Private Passenger" or "Classic-Refer to CO (Stated Amount)" get assignment columns.
  All other types (Motorhome, Recreational Trailer, Utility Trailer, Antique, etc.) are SKIPPED — their
  Veh# column must be blank; no ordinal value should appear in that column.
- N = count of ELIGIBLE (Private Passenger or Classic) vehicles in the transaction group.
- Each driver row covers exactly the columns for eligible vehicles only; ineligible vehicle columns are blank.
- NO gaps: ordinals fill Veh# columns for eligible vehicles consecutively (Veh#1, Veh#2, etc. for the
  columns that correspond to eligible vehicles); columns for ineligible vehicles stay blank.
- Assignment uses backward circular rotation:
    ordinal for eligible-driver i, eligible-vehicle j = (j - i) % N + 1 → "1st", "2nd", "3rd", …
  Example (2 drivers, 2 eligible vehicles):
    Driver 1: Veh#1=1st, Veh#2=2nd
    Driver 2: Veh#1=2nd, Veh#2=1st
  Example (4 drivers, 6 vehicles):
    Driver 1: 1st  2nd  3rd  4th  5th  6th
    Driver 2: 6th  1st  2nd  3rd  4th  5th
    Driver 3: 5th  6th  1st  2nd  3rd  4th
    Driver 4: 4th  5th  6th  1st  2nd  3rd
- Eligible drivers (Principal or Occasional) get ordinals. Ineligible driver types get blank cells.
- The EXACT row mapping below pre-computes all ordinals — follow it exactly.
"""

INFRACTION_SHEET_RULES = """
INFRACTION SHEET RULES:
- Generate one row for EVERY driver (all test case numbers must appear), not just those with infractions.
- Test Case No: same as the driver's Test Case No from Driver sheet (TS-XX-XX-XX).
- For drivers where Add Infraction = Yes in Driver sheet: fill all infraction fields.
- For drivers where Add Infraction = blank/No: row exists but ALL infraction fields are BLANK
  (Infraction Type, Violation Date, SDIP Points, Infraction Description, Violation Code, At Fault Accident,
   Accident City, Accident State, Claim Number, Claim Type, all Amount fields = blank).
- Infraction Type (only for Add Infraction = Yes): "Accident/Loss" or "Moving Violation"
- Violation/Accident Date: MMDDYY format (6 digits, no slashes), within last 3 years
- SDIP Points: 1-99. If multiple infractions, sum must equal 99.
- Infraction Description and Violation Code: use exact pairs from the provided table
- At Fault Accident: Yes or blank. If Comprehensive claim type: must be blank.
- Accident City/State/Claim Number/Claim Type/Amount fields: only if Infraction Type = "Accident/Loss"
- If Claim Type = "No Payment": all Amount fields must be blank.
"""

SUMMARY_SHEET_RULES = """
SUMMARY SHEET RULES:
- Test Case No: same as Policy sheet Test Case No
- Transactions: same as Policy sheet Transaction Type
- Execute Transaction: default "No"
- Hold Transaction: default "No"
- Test Case Details: same as Policy sheet Test Case Details
- All other fields (Wins Reference Quote, Python Reference Quote, Wins Issued Policy Number, Python Issued Policy Number, Wins Premium, Python Premium, Status, Wins Screenshot Link, Python Screenshot Link): ALWAYS BLANK
"""

POLICY_INFO_SHEET_RULES = """
POLICY INFO SHEET RULES:
- All fields are optional and can be blank
- Premium Finance: random name/address
- Bill to address: random name/address
- 3rd Party: random name/address/phone/comments
- Manually Rated: always blank or NO
- Early Issue Start Date: blank or same as effective date in MMDDYY format
- Full Pay Credit: Yes or No
- Declaration Print Option: one of the three options
- Non-Renewal Flag: Yes or No. If Yes, Non-Renew Reason required.
- Insured 1, Score, Status, Reasons: always blank
"""

QUESTIONS_REMARKS_SHEET_RULES = """
PERSONAL AUTO QUESTIONS / REMARKS SHEET RULES:
- Test Case No: same as the Policy sheet Test Case No (TS-XX-XX format, one row per test case).
- Producer Name: CT state only — NANCY MENDIZABAL or MICHAEL PRENDERGAST.
  For ME and RI states: Producer Name must be BLANK.
- All other fields: blank unless the column label makes the value obvious.
"""


def _get_sheet_rules(sheet_type: str) -> str:
    """Return PAP_CORE_RULES + the sheet-specific rule block."""
    rules_map = {
        "policy": POLICY_SHEET_RULES,
        "driver": DRIVER_SHEET_RULES,
        "vehicle": VEHICLE_SHEET_RULES,
        "assignment": ASSIGNMENT_SHEET_RULES,
        "infraction": INFRACTION_SHEET_RULES,
        "summary": SUMMARY_SHEET_RULES,
        "policy_info": POLICY_INFO_SHEET_RULES,
        "questions_remarks": QUESTIONS_REMARKS_SHEET_RULES,
    }
    sheet_rules = rules_map.get(sheet_type, "")
    if not sheet_rules:
        return ""
    return f"{PAP_CORE_RULES}\n{sheet_rules}"


# ---------------------------------------------------------------------------
# Reference tables injected into the prompt for specific sheets
# ---------------------------------------------------------------------------

VIN_TABLE = """
Use ONLY the following VIN data rows. Pick one row per vehicle. Do NOT reuse the same VIN within a test case.

Model Year | Vehicle (VIN)         | Make/Model       | Style
2009       | 2T1BU40E09C096800     | TOYT COROLLA     | SEDAN
1990       | JF2BJ63C7LG943627     | SUBA LEGACY      | WAGON
2010       | 1FAHP3FN2AW144804     | FORD FOCUS       | SEDAN
2013       | 1GTN2TE03DZ342263     | GMC  SIERRA      | PICKUP
2022       | 4S3BWAM61N3003166     | SUBA LEGACY      | SEDAN
2015       | KL4CJHSB2FB263363     | BUIC ENCORE      | UTILITY
2016       | JA4AR4AW2GZ058082     | MITS OUTLANDER   | UTILITY
2020       | 1C4GJXAG2LW143520     | JEEP WRANGLER    | UTILITY
2021       | JF2SKAAC6MH498991     | SUBA FORESTER    | UTILITY

Vehicle Make column should always be blank.
Mature Driver column should always be blank.
Make/Model must be copied exactly as shown (do not split or modify).
"""

INFRACTION_TABLE = """
Use ONLY these Infraction Description and Violation Code pairs (use exact values):

Infraction Description          | Violation Code
LIC SUSPENSION                  | 11105
SUSPEPENSION                    | 30110
DWI                             | 11310
DUI DRUGS                       | 11321
FELONY DUI                      | 11330
REFUSAL TO SUBMIT TO CHEM TEST  | 11335
HOMICIDE                        | 11510
RECKLESS DRIVING                | 11110
NEGLIGENT DRIVING               | 11210
RACING                          | 11830
PREARRANGED RACING              | 11835
SCHOOL BUS VIOLATION            | 11840
SPEED                           | 21010
SPEED 26-30                     | 21625
25-34                           | 21634
SPEED 25 PLUS OVER LIMIT        | 21653
SPEED 31-35 OVER LIMIT          | 21751
SPEED 31 PLUS OVER LIMIT        | 21752
SPEED 36-40 PLUS OVER LIMIT     | 21760
SPEED 40PLUS OVER LIMIT         | 21763
46 PLUS OVER LIMIT              | 21880
SPEED 75 PLUS OVER LIMIT        | 21885
PASSED STOPPED SCHOOL BUS       | 27600
REVOCATION                      | 30120
FAILURE TO REPORT INJURY        | 41240
FAILURE TO REPORT PROP DAMAGE   | 41250
FAILURE TO REPORT-UNSPECIFIED   | 41350
UNATHERIZED USE OF MOTOR VEH    | 11710
MOV                             | 72100
"""


# ---------------------------------------------------------------------------
# Post-processing helpers
# ---------------------------------------------------------------------------

def _get_tc_value(row: dict[str, Any]) -> str:
    """Return the Test Case No value from a row, handling header variants."""
    _EXACT = {"test case no", "test case #", "test case number", "test case"}
    _EXCL = ("detail", "description", "type", "status", "note")
    for key, val in row.items():
        kl = key.lower().strip()
        if kl in _EXACT:
            return str(val).strip() if val else ""
    for key, val in row.items():
        kl = key.lower().strip()
        if kl.startswith("test case") and not any(e in kl for e in _EXCL):
            return str(val).strip() if val else ""
    return ""


def _fill_missing_vehicle_rows(
    rows: list[dict[str, Any]],
    previous_sheets_data: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    """Detect and fill vehicle rows dropped by the LLM.

    The vehicle sheet must have one row per expanded Test Case No
    (TS-XX-YY-01, TS-XX-YY-02, …).  If the LLM only emitted one row per
    policy row (missing the second/third vehicle), we detect the gap by
    comparing generated rows against the expected set computed from the
    Policy sheet, then synthesise the missing rows by cloning the first
    vehicle of that test case and bumping the TC suffix.
    """
    # Find the Policy sheet (first key whose name matches 'policy')
    policy_data: list[dict[str, Any]] = []
    for sname, sdata in previous_sheets_data.items():
        if detect_sheet_type(sname) == "policy":
            policy_data = sdata
            break
    if not policy_data:
        return rows

    policy_structure = parse_policy_structure(policy_data)

    # Build the complete expected TC set
    expected_tcs: list[str] = []
    for ps in policy_structure:
        base = ps["test_case_no"]
        for v in range(1, ps["vehicle_count"] + 1):
            expected_tcs.append(f"{base}-{v:02d}")

    generated_tcs = {_get_tc_value(r) for r in rows}
    missing = [tc for tc in expected_tcs if tc not in generated_tcs]

    if not missing:
        return rows

    print(
        f"[post_process] Vehicle sheet missing {len(missing)} row(s): {missing}. "
        "Synthesising placeholders."
    )

    # For each missing TC, clone the first existing row for that test case
    # and update the TC field(s).
    # Identify the TC column key from existing rows.
    tc_key: str | None = None
    if rows:
        _EXACT_SET = {"test case no", "test case #", "test case number", "test case"}
        _EXCL = ("detail", "description", "type", "status", "note")
        for key in rows[0].keys():
            kl = key.lower().strip()
            if kl in _EXACT_SET or (
                kl.startswith("test case") and not any(e in kl for e in _EXCL)
            ):
                tc_key = key
                break

    for missing_tc in missing:
        # parent group: TS-XX-YY (first three segments)
        parts = missing_tc.split("-")
        parent_grp = "-".join(parts[:3]) if len(parts) >= 3 else missing_tc

        # Find a donor row from the same group
        donor: dict[str, Any] | None = None
        for r in rows:
            tc_val = _get_tc_value(r)
            if tc_val.startswith(parent_grp):
                donor = r
                break

        new_row: dict[str, Any]
        if donor and tc_key:
            new_row = {k: v for k, v in donor.items()}
            new_row[tc_key] = missing_tc
        else:
            # No donor found — minimal placeholder
            new_row = {tc_key or "Test Case No": missing_tc} if tc_key else {}

        rows.append(new_row)

    # Re-sort rows to match expected_tcs order
    tc_order = {tc: idx for idx, tc in enumerate(expected_tcs)}
    rows.sort(key=lambda r: tc_order.get(_get_tc_value(r), 9999))

    return rows


def _fix_driver_ages(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Recalculate Age from Date of Birth (MMDDYYYY) for every driver row.

    The LLM sometimes produces an off-by-one age. This replaces the LLM
    value with the correct integer age as of today.
    """
    from datetime import date as _date

    today = _date.today()
    dob_keys_seen: set[str] = set()
    age_key: str | None = None

    if not rows:
        return rows

    # Locate Age and DOB column keys from the first row
    dob_key: str | None = None
    for key in rows[0].keys():
        kl = key.lower()
        if "date of birth" in kl or "dob" in kl:
            dob_key = key
        if kl.strip() == "age" or kl.strip() == "driver age":
            age_key = key

    if not dob_key or not age_key:
        return rows  # can't fix without both columns

    for row in rows:
        raw_dob = str(row.get(dob_key, "")).strip()
        if len(raw_dob) != 8 or not raw_dob.isdigit():
            continue
        try:
            month = int(raw_dob[:2])
            day = int(raw_dob[2:4])
            year = int(raw_dob[4:])
            dob = _date(year, month, day)
            age = (
                today.year - dob.year
                - ((today.month, today.day) < (dob.month, dob.day))
            )
            row[age_key] = age
        except (ValueError, OverflowError):
            pass  # leave LLM value intact if DOB is unparseable

    return rows


def _blank_vehicle_mature_driver(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Ensure Mature Driver# and Mature Driver columns are always blank."""
    for row in rows:
        for k in list(row.keys()):
            kl = k.lower().strip()
            if "mature driver" in kl:
                row[k] = ""
    return rows


def _normalize_pap_dates(
    rows: list[dict[str, Any]],
    sheet_type: str,
) -> list[dict[str, Any]]:
    """Strip slashes/dashes from date fields so format is always MMDDYYYY.

    The Quincy template sometimes has legacy data with MM/DD/YYYY; the LLM
    may copy that style. This enforcer ensures generated output is clean.
    Infraction dates use MMDDYY (6 digits) — those are untouched here because
    they rarely have slashes and have a different length.
    """
    _DATE_KEYWORDS = (
        "effective date", "endorsement date", "cancel date",
        "reinstate date", "retro endorsement date",
        "date of birth", "dob", "license date",
        "early issue start date",
    )
    # Infraction sheet dates are 6-digit MMDDYY — skip to avoid stripping wrong fields
    if sheet_type == "infraction":
        return rows

    for row in rows:
        for k in list(row.keys()):
            kl = k.lower()
            if any(dk in kl for dk in _DATE_KEYWORDS):
                val = str(row.get(k, "") or "").strip()
                # Remove slashes and dashes: "03/30/2026" → "03302026"
                cleaned = re.sub(r"[/\-]", "", val)
                if cleaned != val and cleaned.isdigit():
                    row[k] = cleaned
    return rows


# ---------------------------------------------------------------------------
# Handler class (dispatcher-facing interface)
# ---------------------------------------------------------------------------

class PapQuincyHandler:
    policy_type = "PAP"

    def detect_sheet_type(self, sheet_name: str) -> str:
        return detect_sheet_type(sheet_name)

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        sheet_type = detect_sheet_type(sheet_name)
        rules = _get_sheet_rules(sheet_type)

        if not rules:
            return original_row_count, ""

        extra_parts = [rules]
        if sheet_type == "vehicle":
            extra_parts.append(VIN_TABLE)
        if sheet_type == "infraction":
            extra_parts.append(INFRACTION_TABLE)

        adjusted_row_count = original_row_count
        if policy_data:
            policy_structure = parse_policy_structure(policy_data)

            expanded_count = calculate_expanded_row_count(sheet_type, policy_structure, driver_data)
            if expanded_count is not None and expanded_count > 0:
                adjusted_row_count = expanded_count

            mapping = build_row_mapping_instructions(sheet_type, policy_structure, driver_data)
            if mapping:
                extra_parts.append(mapping)

        return adjusted_row_count, "\n\n".join(extra_parts)

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
    ) -> list[dict[str, Any]] | None:
        if detect_sheet_type(sheet_name) != "assignment":
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
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        sheet_type = detect_sheet_type(sheet_name)

        # ── Vehicle sheet: fill any rows the LLM silently dropped ────────────
        if sheet_type == "vehicle" and previous_sheets_data:
            rows = _fill_missing_vehicle_rows(rows, previous_sheets_data)

        # ── Vehicle sheet: strip "Mature Driver" columns ─────────────────────
        if sheet_type == "vehicle":
            rows = _blank_vehicle_mature_driver(rows)

        # ── Driver sheet: correct Age based on Date of Birth ─────────────────
        if sheet_type == "driver":
            rows = _fix_driver_ages(rows)

        # ── All sheets: normalize date fields to MMDDYYYY (strip slashes) ────
        rows = _normalize_pap_dates(rows, sheet_type)

        return rows
