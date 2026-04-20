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

    Filename prefixes (case-insensitive) determine the handler:
      IMS*  → IMS   (commercial insurance, 9-sheet workbook)
      PAP*  → PAP   (Personal Auto Policy — Quincy)
      MCA*  → MCA   (future auto policy type)
      else  → GENERIC (fallback: preserves baseline main behavior)

    The returned string is used as the key in policies.get_handler().
    """
    name = (filename or "").strip().upper()
    # Strip any leading path separators
    name = name.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    if name.startswith("IMS"):
        return "IMS"
    if name.startswith("PAP"):
        return "PAP"
    if name.startswith("MCA"):
        return "MCA"
    return "GENERIC"



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
