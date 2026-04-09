from openai import OpenAI
import json
import os
import time
import random
import re
from datetime import date
from typing import Any


def _clean_response_text(text: str) -> str:
    """Normalize common fenced-code formatting around model output."""
    cleaned = text.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    if cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    return cleaned.strip()


def _parse_json_array(text: str) -> list[dict[str, Any]]:
    """Parse model output, extracting array from wrapped objects when needed."""
    cleaned = _clean_response_text(text)

    parsed = json.loads(cleaned)

    if isinstance(parsed, list):
        return parsed

    if isinstance(parsed, dict):
        # Look for any value that is a list of dicts (the generated rows)
        for value in parsed.values():
            if isinstance(value, list) and value and isinstance(value[0], dict):
                return value
        # Fallback: single dict might be one row
        return [parsed]

    raise ValueError(f"Model output is not a JSON array or wrapped array. Got: {type(parsed)}")


def _extract_effective_expiration_year_range(special_instruction: str) -> tuple[int, int] | None:
    """Extract year range like 2020-2026 when instructions mention effective/expiration dates."""
    if not special_instruction:
        return None

    text = special_instruction.lower()
    has_date_intent = any(
        phrase in text
        for phrase in [
            "effective",
            "expiration",
            "expiry",
            "effective date",
            "expiration date",
        ]
    )
    if not has_date_intent:
        return None

    match = re.search(r"(20\d{2})\s*(?:-|to|–|—)\s*(20\d{2})", text)
    if not match:
        return None

    start_year = int(match.group(1))
    end_year = int(match.group(2))
    if start_year > end_year:
        start_year, end_year = end_year, start_year
    return start_year, end_year


def _has_date_intent(special_instruction: str) -> bool:
    text = (special_instruction or "").lower()
    return any(
        phrase in text
        for phrase in [
            "effective",
            "expiration",
            "expiry",
            "date range",
            "from 20",
            "current date",
            "current year",
        ]
    )


def _include_every_year_requested(special_instruction: str) -> bool:
    text = (special_instruction or "").lower()
    return any(
        phrase in text
        for phrase in [
            "every year",
            "each year",
            "all years",
            "year from",
            "years from",
            "to be included",
            "include all years",
        ]
    )


def _find_header_key(row: dict[str, Any], candidates: list[str]) -> str | None:
    """Find a header key by case-insensitive substring match."""
    for key in row.keys():
        key_lower = key.lower()
        if any(candidate in key_lower for candidate in candidates):
            return key
    return None


def _format_mmddyyyy(value: date) -> str:
    return value.strftime("%m/%d/%Y")


def _add_one_year(value: date) -> date:
    try:
        return value.replace(year=value.year + 1)
    except ValueError:
        # Handle Feb 29 -> Feb 28 on non-leap years.
        return value.replace(month=2, day=28, year=value.year + 1)


def _build_date_policy_summary(special_instruction: str) -> str:
    """Return compact guidance injected into the prompt."""
    year_range = _extract_effective_expiration_year_range(special_instruction)
    include_every_year = _include_every_year_requested(special_instruction)

    if year_range:
        start_year, end_year = year_range
        if include_every_year:
            return (
                "Apply date range policy from user instruction: include all years in "
                f"{start_year}-{end_year} across generated rows, and keep expiration one year after effective when possible."
            )
        return (
            "Apply date range policy from user instruction: effective/expiration dates must follow "
            f"the range {start_year}-{end_year}."
        )

    return "No explicit date range instruction. Use current year for effective date; expiration is one year later."


def _enforce_effective_expiration_date_range(
    rows: list[dict[str, Any]],
    special_instruction: str,
) -> list[dict[str, Any]]:
    """Apply instruction-driven policy for effective/expiration dates."""
    year_range = _extract_effective_expiration_year_range(special_instruction)
    effective_candidates = ["effective date", "effective"]
    expiration_candidates = ["expiration date", "expiry date", "expiry", "expiration", "exp date"]

    include_every_year = _include_every_year_requested(special_instruction)
    today = date.today()

    effective_cycle: list[int] = []
    if year_range:
        start_year, end_year = year_range
        if end_year - start_year >= 1:
            effective_cycle = list(range(start_year, end_year + 1))
        else:
            effective_cycle = [start_year]

        if include_every_year and not effective_cycle:
            effective_cycle = [start_year]

    for idx, row in enumerate(rows):
        effective_key = _find_header_key(row, effective_candidates)
        expiration_key = _find_header_key(row, expiration_candidates)
        if not effective_key and not expiration_key:
            continue

        if not year_range and not _has_date_intent(special_instruction):
            effective_dt = today
            expiration_dt = _add_one_year(effective_dt)
        elif year_range:
            start_year, end_year = year_range
            if include_every_year and effective_cycle:
                effective_year = effective_cycle[idx % len(effective_cycle)]
            else:
                if end_year - start_year >= 1:
                    effective_year = random.randint(start_year, end_year)
                else:
                    effective_year = start_year

            effective_month = random.randint(1, 12)
            effective_day = random.randint(1, 28)
            effective_dt = date(effective_year, effective_month, effective_day)
            expiration_dt = _add_one_year(effective_dt)

            if expiration_dt.year > end_year:
                expiration_dt = date(end_year, effective_month, effective_day)
        else:
            effective_dt = today
            expiration_dt = _add_one_year(effective_dt)

        if effective_key:
            row[effective_key] = _format_mmddyyyy(effective_dt)
        if expiration_key:
            row[expiration_key] = _format_mmddyyyy(expiration_dt)

    return rows


def configure_openai() -> OpenAI:
    """Configure OpenAI client with API key from environment."""
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise ValueError("OPENAI_API_KEY not found")
    return OpenAI(api_key=key)


# ---------------------------------------------------------------------------
# Insurance domain: sheet type detection and multi-row expansion
# ---------------------------------------------------------------------------

def detect_sheet_type(sheet_name: str) -> str:
    """Detect insurance sheet type from sheet name."""
    name = sheet_name.lower().strip()
    if "infraction" in name:
        return "infraction"
    if "assignment" in name:
        return "assignment"
    if "policy info" in name or "policy_info" in name:
        return "policy_info"
    if "summary" in name:
        return "summary"
    if "driver" in name:
        return "driver"
    if "vehicle" in name:
        return "vehicle"
    if "policy" in name:
        return "policy"
    return "unknown"


def parse_policy_structure(policy_data: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Parse Policy sheet data to extract test case structure (driver/vehicle counts, etc.)."""
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

        for key, value in row.items():
            kl = key.lower()
            val = str(value).strip() if value else ""
            if "test case no" in kl:
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

        # Parse "X Driver & Y Vehicle" from test_case_details
        m = re.search(r"(\d+)\s*Driver.*?(\d+)\s*Vehicle", info["test_case_details"], re.IGNORECASE)
        if m:
            info["driver_count"] = int(m.group(1))
            info["vehicle_count"] = int(m.group(2))

        structure.append(info)

    return structure


def calculate_expanded_row_count(
    sheet_type: str,
    policy_structure: list[dict[str, Any]],
    driver_data: list[dict[str, Any]] | None = None,
) -> int | None:
    """Calculate the correct row count for a sheet based on Policy structure.

    Returns None if no expansion is needed (use original row_count).
    """
    if not policy_structure:
        return None

    if sheet_type == "driver":
        return sum(ps["driver_count"] for ps in policy_structure)
    elif sheet_type == "vehicle":
        return sum(ps["vehicle_count"] for ps in policy_structure)
    elif sheet_type == "assignment":
        return sum(ps["driver_count"] for ps in policy_structure)
    elif sheet_type == "infraction":
        # Infraction rows depend on which drivers have Add Infraction = Yes
        if driver_data:
            count = 0
            for drv in driver_data:
                for key, val in drv.items():
                    if "add infraction" in key.lower() and str(val).strip().lower() == "yes":
                        count += 1
                        break
            return max(count, 0)
        return None  # Will be determined by LLM
    elif sheet_type in ("summary", "policy_info"):
        return len(policy_structure)

    return None


def build_row_mapping_instructions(
    sheet_type: str,
    policy_structure: list[dict[str, Any]],
    driver_data: list[dict[str, Any]] | None = None,
) -> str:
    """Build explicit row-by-row instructions for expanded sheets."""
    if not policy_structure:
        return ""

    lines = []

    if sheet_type == "driver":
        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            insured = ps["insured_name"]
            is_new_business = "new business" in txn_type.lower()

            for d in range(1, ps["driver_count"] + 1):
                row_num += 1
                expanded_no = f"{base_no}-{d:02d}"
                txn = "Add" if is_new_business else "Edit"
                name_hint = f', Name = "{insured}" (same as Insured Name)' if d == 1 and insured else f", Name = random unique name"
                lines.append(
                    f"  Row {row_num}: Test Case No = \"{expanded_no}\", "
                    f"Transaction Type = \"{txn_type}\", Transaction = \"{txn}\"{name_hint}"
                )

        return (
            "EXACT ROW MAPPING FOR DRIVER SHEET:\n"
            f"You MUST generate exactly {row_num} rows with these exact Test Case No values:\n"
            + "\n".join(lines)
            + "\n\nDo NOT deviate from this mapping. Each row above is one driver record."
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
        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            driver_count = ps["driver_count"]
            vehicle_count = ps["vehicle_count"]

            for d in range(1, driver_count + 1):
                row_num += 1
                expanded_no = f"{base_no}-{d:02d}"

                # Build vehicle assignment hint
                veh_hints = []
                if driver_count == 1 and vehicle_count == 1:
                    veh_hints.append("Veh #1 = \"1st\"")
                else:
                    for v in range(1, vehicle_count + 1):
                        if d <= vehicle_count and v == d:
                            veh_hints.append(f"Veh #{v} = \"1st\"")
                        elif d > vehicle_count and v == 1:
                            veh_hints.append(f"Veh #{v} = \"1st\"")
                        else:
                            veh_hints.append(f"Veh #{v} = \"\" (blank)")

                lines.append(
                    f"  Row {row_num}: Test Case No = \"{expanded_no}\", "
                    f"Transaction Type = \"{txn_type}\", "
                    + ", ".join(veh_hints)
                )

        return (
            "EXACT ROW MAPPING FOR ASSIGNMENT SHEET:\n"
            f"You MUST generate exactly {row_num} rows with these exact Test Case No values:\n"
            + "\n".join(lines)
            + "\n\nDo NOT deviate from this mapping."
            + "\nEach driver MUST have at least one vehicle assigned as \"1st\"."
            + "\nThe Name and Driver Type columns must match the corresponding Driver sheet rows."
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

    elif sheet_type == "infraction":
        if not driver_data:
            return ""
        row_num = 0
        for drv in driver_data:
            has_infraction = False
            test_case_no = ""
            txn_type = ""
            for key, val in drv.items():
                kl = key.lower()
                if "add infraction" in kl and str(val).strip().lower() == "yes":
                    has_infraction = True
                if "test case no" in kl:
                    test_case_no = str(val).strip()
                if "transaction type" in kl:
                    txn_type = str(val).strip()

            if has_infraction:
                row_num += 1
                is_new_business = "new business" in txn_type.lower()
                txn = "Add" if is_new_business else "Edit"
                lines.append(
                    f"  Row {row_num}: Test Case No = \"{test_case_no}\", "
                    f"Transaction Type = \"{txn_type}\", Transaction = \"{txn}\""
                )

        if not lines:
            return "No drivers have Add Infraction = Yes. Generate 0 rows (empty sheet with headers only)."

        return (
            "EXACT ROW MAPPING FOR INFRACTION SHEET:\n"
            f"You MUST generate exactly {row_num} rows for drivers with Add Infraction = Yes:\n"
            + "\n".join(lines)
        )

    return ""


# ---------------------------------------------------------------------------
# Insurance-specific sheet rules (condensed for prompt injection)
# ---------------------------------------------------------------------------

POLICY_SHEET_RULES = """
POLICY SHEET RULES:
- Test Case No format: TS-XX-XX where first XX = test case number, second XX = transaction number
- Transaction Types: New Business (always first), Endorsement, Cancellation (Flat Cancel or Mid Term Cancel), Add From Cancel (only after Flat Cancel, same date), Reinstatement (only after Mid Term Cancel, same date), Retroactive (needs 2+ prior endorsements), Renewal (300+ days after effective date)
- Policy Change: Yes for New Business (default). For Endorsement/Renewal/Retroactive/Add From Cancel: Yes if any policy fields changed, No otherwise. Always No for Cancellation/Reinstatement.
- Driver Change: Yes for New Business. Yes for Endorsement/Renewal/Retroactive/Add From Cancel only if driver fields changed. No for Cancellation/Reinstatement.
- Vehicle Change: Yes for New Business. Yes for Endorsement/Renewal/Retroactive/Add From Cancel only if vehicle fields changed. No for Cancellation/Reinstatement.
- Assignment Change: Yes for New Business. Yes for Endorsement/Renewal/Retroactive/Add From Cancel only if assignment fields changed. No for Cancellation/Reinstatement.
- Test Case Details: Format is "X Driver & Y Vehicle" - this defines EXACT driver and vehicle counts.
- Effective Date: mmddyyyy format (NO slashes). Current date ±60 days. For Renewal: 300 days in past.
- Endorsement Date: mmddyyyy format (NO slashes). Same as Effective Date for New Business. For other transactions: >= New Business date and < 50 days future. For Cancellation: > New Business date. For Reinstatement: same as Cancellation date. For Retroactive: previous Endorsement date.
- Rating state = ME:
  - Company Code = 0020
  - Agent = 05899, 05297 or 05899
  - Loss Free = 1 to 9
  - Type, UMPD, CT PIP-BRB = Blank (CT-only fields)
- Phone format: 1111111111 (10 digits, no formatting)
- Email: ends with test.com or gmail.com
- Liability CSL and Liability BI are mutually exclusive (only one can be selected)
- UM/UIM CSL <= Liability CSL; UM/UIM BI <= Liability BI
- PD only if Liability BI is selected
- Payment Plan: one value for entire test case
- Monthly Due Day: only if Payment Plan = "EFT - 10 Pay Pick a Day", value 1-30
- WINS Quote Number, Python Quote Number, Wins Policy Number, Python Policy Number, Insurance Score, Client ID, EFT = always Blank
- Less than 3 years at current address = Yes: fill Previous Street/City/State/ZIP. No: leave blank.
- Retroactive options: Endorse(option=1), Final Cancel(option=3), Reinstate(option=4)
"""

DRIVER_SHEET_RULES = """
DRIVER SHEET RULES:
- Test Case No format: TS-XX-XX-XX where last XX = driver number (01, 02, etc.)
- Transaction: Add for New Business, Edit for Endorsement/Renewal/Retroactive/Add From Cancel, Delete only if driver count > 1
- First driver's Name = Insured Name from Policy sheet. Additional drivers = random names.
- Date of Birth: mmddyyyy format (NO slashes). Year between 1930 and current year minus 16.
- Age: calculated from Date of Birth
- Driver Type: Principal (mandatory if only 1 driver), Occasional for additional drivers
- Relationship To Insured: first driver = "Insured"
- License Date: Date of Birth + 18 years, mmddyyyy format (NO slashes)
- License #: ME state = 7 digits, CT state = 9 digits, RI state = 7-9 digits
- Lic State: preferably same as Rating State
- Driver Training: Yes only if driver age <= 20. If Yes, completion date = License Date.
- Good Student: Yes only if driver age <= 25
- Operator student 100+ miles: Yes only if driver age <= 20
- Mature Credit: if Yes, completion date = License Date
- MVR Re-Order: always NO
- Claims Report Re-Order: always NO
- Violations/Accidents in last 3 years: always Blank
- Add Infraction: Yes or Blank. If Yes, Infraction sheet must have data for this driver.
- Remove Infraction: Yes or Blank. If Yes, Remove Infraction SDIP is mandatory.
"""

VEHICLE_SHEET_RULES = """
VEHICLE SHEET RULES:
- Test Case No format: TS-XX-XX-XX where last XX = vehicle number (01, 02, etc.)
- Transaction: Add for New Business, Edit for Endorsement/Renewal/Retroactive/Add From Cancel, Delete only if vehicle count > 1
- Territory: always Blank
- If only 1 vehicle: must be Private Passenger or Classic type only
- Vehicle Use: Pleasure, Business, Commute, Farming. Motorhome/Utility Trailer/Antique = Pleasure only. If Commute: Miles One-way mandatory.
- VIN: Use exact data from the provided VIN table. NO duplicate VINs in same test case.
- Model Year, Vehicle (VIN), Vehicle Make, Make/Model, Style: use exact rows from VIN table.
- Vehicle Make column = blank unless data provided in table.
- Pass/Rest, Anti-Theft: use one of the specified values exactly.
- Anti-Lock: always Yes
- Comp Ded: blank or 50/100/200/250/500/1,000 (use 50 very rarely)
- Full Glass: Yes only if Private Passenger type. No = blank.
- Coll Ded: blank or 50/100/200/250/500/1,000 (use 50 very rarely)
- Sub Trans: only if Comp Ded >= 100. Values: 30/900, 40/1,000, 50/1,500 or blank.
- Towing: only if Comp Ded >= 100. Values: 25, 50, 75, 100 or blank.
- Corporate Car Discount: Yes only if Comp Ded >= 100, else blank.
- Joint Ownership: Yes only if Comp Ded >= 100, else blank.
- Enhancement Endorsement: Yes only if Coll Ded >= 100, else blank.
- Trip Interrupt 600: Yes only if Comp Ded >= 100, else blank.
- Cost New: only for Private Passenger, Motorhome, Recreational Trailer. 1,000-100,000 no $ symbol.
- Stated Amt: only for Classic, Antique, Utility Trailer. 1,000-100,000 no $ symbol.
- Customized Amt: thousands format (2,000 / 5,000 / 10,000 etc.), no $ symbol.
- Delete Liability and Suspend Liability: must have same value (both Yes or both blank).
- Loan/Lease Coverage: always blank.
- RI UMPD: blank for ME state.
"""

ASSIGNMENT_SHEET_RULES = """
ASSIGNMENT SHEET RULES:
- Test Case No format: TS-XX-XX-XX (same suffix as driver number)
- Name: must match exactly the driver's Name from Driver sheet
- Driver Type: must match exactly the driver's Driver Type from Driver sheet
- If 1 Driver & 1 Vehicle: Veh #1 = "1st"
- Multiple drivers/vehicles: each driver must be assigned to at least one vehicle as "1st"
- Unassigned vehicle columns should be blank
"""

INFRACTION_SHEET_RULES = """
INFRACTION SHEET RULES:
- Only generate rows for drivers where Add Infraction = Yes in Driver sheet
- Test Case No: same as the driver's Test Case No from Driver sheet
- Infraction Type: "Accident/Loss" or "Moving Violation"
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


def _get_sheet_rules(sheet_type: str) -> str:
    """Return insurance-specific rules for a given sheet type."""
    rules_map = {
        "policy": POLICY_SHEET_RULES,
        "driver": DRIVER_SHEET_RULES,
        "vehicle": VEHICLE_SHEET_RULES,
        "assignment": ASSIGNMENT_SHEET_RULES,
        "infraction": INFRACTION_SHEET_RULES,
        "summary": SUMMARY_SHEET_RULES,
        "policy_info": POLICY_INFO_SHEET_RULES,
    }
    return rules_map.get(sheet_type, "")


# ---------------------------------------------------------------------------
# VIN table for vehicle data
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


def build_insurance_context(
    sheet_name: str,
    policy_data: list[dict[str, Any]] | None,
    driver_data: list[dict[str, Any]] | None,
    original_row_count: int,
) -> tuple[int, str]:
    """Build insurance-specific context for a sheet.

    Returns:
        (adjusted_row_count, additional_prompt_text)
    """
    sheet_type = detect_sheet_type(sheet_name)
    rules = _get_sheet_rules(sheet_type)

    if not rules:
        return original_row_count, ""

    extra_parts = [rules]

    # Add VIN table for vehicle sheet
    if sheet_type == "vehicle":
        extra_parts.append(VIN_TABLE)

    # Add infraction table
    if sheet_type == "infraction":
        extra_parts.append(INFRACTION_TABLE)

    # Calculate row count and build row mapping if we have policy structure
    adjusted_row_count = original_row_count
    if policy_data:
        policy_structure = parse_policy_structure(policy_data)

        expanded_count = calculate_expanded_row_count(sheet_type, policy_structure, driver_data)
        if expanded_count is not None and expanded_count > 0:
            adjusted_row_count = expanded_count
        elif expanded_count == 0 and sheet_type == "infraction":
            adjusted_row_count = 0

        mapping = build_row_mapping_instructions(sheet_type, policy_structure, driver_data)
        if mapping:
            extra_parts.append(mapping)

    return adjusted_row_count, "\n\n".join(extra_parts)


# ---------------------------------------------------------------------------
# Main generation function
# ---------------------------------------------------------------------------

def generate_test_data(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str = "",
    previous_sheets_data: dict[str, Any] | None = None,
    row_count_override: int | None = None,
    additional_rules: str = "",
) -> list[dict[str, Any]]:
    """Generate test data using OpenAI for US insurance domain."""

    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "")

    effective_row_count = row_count_override if row_count_override is not None else row_count

    # Handle zero-row sheets (e.g., infraction with no infractions)
    if effective_row_count == 0:
        return []

    # Build context from previously generated sheets
    previous_context = ""
    if previous_sheets_data:
        previous_context = """
CRITICAL - CROSS-SHEET DATA CONSISTENCY:
The following sheets have already been generated. You MUST reuse the SAME values for matching/similar fields across sheets.
Row-level correspondence is based on Test Case No matching, NOT row position.
For Driver/Vehicle/Assignment sheets, expand rows according to the row mapping below.

Fields that MUST stay consistent across sheets include (when the column exists):
- Test Case No (matching prefix determines the test case)
- Names (first driver Name = Insured Name from Policy)
- Addresses (Street, City, State, ZIP)
- Phone numbers, Email addresses
- Policy numbers, Account numbers
- Effective dates, Endorsement dates
- Transaction Type must match the Policy sheet for the same Test Case No prefix
- Rating State and all state-dependent fields

Previously generated data:
"""
        for prev_sheet_name, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- Sheet: '{prev_sheet_name}' ---\n"
            previous_context += json.dumps(prev_data, indent=2)
            previous_context += "\n"

    user_instruction_block = (special_instruction or "").strip()
    if not user_instruction_block:
        user_instruction_block = "No additional instructions provided."
    date_policy_summary = _build_date_policy_summary(special_instruction)

    # Additional insurance rules
    insurance_rules_block = ""
    if additional_rules:
        insurance_rules_block = f"""
INSURANCE DOMAIN RULES (MUST FOLLOW):
{additional_rules}
"""

    prompt = f"""You are a test data generator for US insurance applications.

You are generating data for sheet: "{sheet_name}"

Generate exactly {effective_row_count} rows of realistic test data for the following columns:
{json.dumps(headers)}

DUPLICATE COLUMN HANDLING:
Some columns may have a suffix like "(Col X)" — this means the original spreadsheet has multiple columns with the same name.
You MUST generate DIFFERENT values for each such column. For example, "Address (Col 3)" and "Address (Col 7)" are two separate columns that need different address values.
Use ALL column names exactly as provided (including suffixes) as JSON keys.

{previous_context}

{insurance_rules_block}

INSTRUCTION PRIORITY (highest to lowest):
1. INSURANCE DOMAIN RULES above (row mappings, sheet-specific rules).
2. USER SPECIAL INSTRUCTIONS (MUST be followed exactly when relevant):
    {user_instruction_block}
3. Cross-sheet consistency constraints from previously generated sheets.
4. Default important rules below (apply ONLY when they do not conflict with above).

DATE POLICY FOR EFFECTIVE/EXPIRATION:
{date_policy_summary}

IMPORTANT RULES:
1. Generate realistic US insurance test data
2. Use valid US formats:
   - Phone: 1111111111 (10 digits, no dashes or parentheses)
   - SSN: XXX-XX-XXXX (use fake but valid format)
   - ZIP: 5 digits or ZIP+4
   - State: 2-letter abbreviations
   - Dates: MMDDYYYY format (8 digits, NO slashes, NO dashes)
3. For policy numbers, use realistic formats used in United States.
4. For currency amounts, use realistic insurance values
5. Names should be diverse and realistic
6. Addresses should be realistic US addresses related to the Rating State
7. For vehicles, use exact data from the VIN table provided in the rules.
8. All the Dollar values should be numbers and not in string. But should have a dollar sign before them.
9. The effective date, endorsement date should match across all sheets for same test case.
10. CRITICAL: Generate EXACTLY {effective_row_count} rows. No more, no less.

Return a JSON object with a single key "data" whose value is an array of {effective_row_count} objects.
Each object must use the exact column names as keys.
No markdown, no explanation, just the JSON object.

Example format:
{{"data": [{{"column1": "value1", "column2": "value2"}}, ...]}}
"""

    # Retry up to 2 times for rate limits only. JSON mode guarantees valid output.
    max_retries = 2
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model_name,
                temperature=1,
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a precise JSON-only data generator for US insurance test datasets. "
                            "Always return a JSON object with a single key \"data\" containing an array of row objects. "
                            "If insurance domain rules specify exact row mappings with Test Case No values, follow them EXACTLY. "
                            "If user special instructions conflict with defaults, follow user special instructions."
                        ),
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
            )

            content = response.choices[0].message.content
            text: str = content if isinstance(content, str) else ""
            if not text.strip():
                raise ValueError("OpenAI returned an empty response")

            data = _parse_json_array(text)

            # Ensure all rows have all requested headers.
            normalized: list[dict[str, Any]] = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({header: row.get(header, "") for header in headers})
            if not normalized:
                raise ValueError("Model output did not contain valid row objects")

            normalized = _enforce_effective_expiration_date_range(
                normalized,
                special_instruction,
            )

            return normalized
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "quota" in error_msg.lower() or "rate" in error_msg.lower():
                wait_time = 60 * (attempt + 1)
                print(f"Rate limit hit for sheet '{sheet_name}'. Waiting {wait_time}s before retry {attempt + 1}/{max_retries}...")
                time.sleep(wait_time)
                if attempt == max_retries - 1:
                    raise
            else:
                raise

    raise RuntimeError("OpenAI response was not received after retries")
