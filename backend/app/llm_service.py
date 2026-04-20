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
    return value.strftime("%m%d%Y")


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
    from datetime import timedelta

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

        # Check transaction type — Renewal needs effective 300+ days in the past
        txn_type = ""
        for key, val in row.items():
            if "transaction type" in key.lower():
                txn_type = str(val or "").strip().lower()
                break
        is_renewal = "renewal" in txn_type

        if is_renewal:
            # Renewal: effective date must be >= 300 days in the past
            days_back = random.randint(300, 400)
            effective_dt = today - timedelta(days=days_back)
            expiration_dt = _add_one_year(effective_dt)
        elif not year_range and not _has_date_intent(special_instruction):
            # Default: current date ±60 days
            offset = random.randint(-60, 60)
            effective_dt = today + timedelta(days=offset)
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
            offset = random.randint(-60, 60)
            effective_dt = today + timedelta(days=offset)
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
# Policy type detection (PAP, MCA, etc.)
# ---------------------------------------------------------------------------

def detect_policy_type(filename: str) -> str:
    """Detect insurance policy type from the uploaded filename.

    Supports:
      PAP  – Personal Auto Policy
      MCA  – (future)

    Returns the policy-type string, defaulting to 'PAP' when unrecognised.
    """
    name = (filename or "").upper()
    if "PAP" in name:
        return "PAP"
    if "MCA" in name:
        return "MCA"
    # Treat any auto-related filename as PAP by default
    return "PAP"


# ---------------------------------------------------------------------------
# Valid driver/vehicle combinations (strict – only these 6 are allowed)
# ---------------------------------------------------------------------------

VALID_DRIVER_VEHICLE_COMBINATIONS: set[tuple[int, int]] = {
    (1, 1),
    (2, 1),
    (1, 2),
    (2, 2),
    (3, 1),
    (4, 6),
}


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
            d = int(m.group(1))
            v = int(m.group(2))
            if (d, v) in VALID_DRIVER_VEHICLE_COMBINATIONS:
                info["driver_count"] = d
                info["vehicle_count"] = v
            else:
                # Clamp to closest valid combination – default 1 Driver & 1 Vehicle
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
        from app.assignment_logic import to_ordinal as _to_ordinal

        row_num = 0
        for ps in policy_structure:
            base_no = ps["test_case_no"]
            txn_type = ps["transaction_type"]
            driver_count = ps["driver_count"]
            # N = vehicle_count from policy details (eligible count may be less, but
            # without actual vehicle type data here we use the declared total as a
            # safe approximation; the deterministic path in main.py uses actual types)
            N = ps["vehicle_count"]

            for i in range(driver_count):
                row_num += 1
                expanded_no = f"{base_no}-{i + 1:02d}"
                # Backward rotation: ordinal = (j - i) % N + 1
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
  For Renewal transactions: effective date must be 300+ days in the past.
- Endorsement Date: MMDDYYYY format (NO slashes). Same as Effective Date for New Business.
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
    If only 1 driver → must be Principal.
    If multiple drivers → first = Principal, others = Occasional.
    Principal is only valid for Private Passenger or Classic vehicle types.
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


# ---------------------------------------------------------------------------
# Personal Auto Policy (PAP) – hardcoded backend rules
# Always injected for PAP files regardless of rule-book availability
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


def _get_sheet_rules(sheet_type: str, policy_type: str = "PAP") -> str:
    """Return insurance-specific rules for a given sheet type, merged with policy-type core rules."""
    rules_map = {
        "policy": POLICY_SHEET_RULES,
        "driver": DRIVER_SHEET_RULES,
        "vehicle": VEHICLE_SHEET_RULES,
        "assignment": ASSIGNMENT_SHEET_RULES,
        "infraction": INFRACTION_SHEET_RULES,
        "summary": SUMMARY_SHEET_RULES,
        "policy_info": POLICY_INFO_SHEET_RULES,
    }
    sheet_rules = rules_map.get(sheet_type, "")
    if not sheet_rules:
        return ""

    # Prepend policy-type core rules for all supported types
    core = ""
    if policy_type == "PAP":
        core = PAP_CORE_RULES
    return f"{core}\n{sheet_rules}" if core else sheet_rules


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
    policy_type: str = "PAP",
) -> tuple[int, str]:
    """Build insurance-specific context for a sheet.

    Returns:
        (adjusted_row_count, additional_prompt_text)
    """
    sheet_type = detect_sheet_type(sheet_name)
    rules = _get_sheet_rules(sheet_type, policy_type)

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
