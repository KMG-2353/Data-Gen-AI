import anthropic
import json
import os
import time


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


def configure_claude() -> anthropic.Anthropic:
    """Configure Anthropic client with API key from environment."""
    key = os.getenv("CLAUDE_API_KEY")
    if not key:
        raise ValueError("CLAUDE_API_KEY not found")
    return anthropic.Anthropic(api_key=key)


def generate_test_data(headers: list[str], row_count: int, special_instruction: str, sheet_name: str = "", previous_sheets_data: dict = None) -> list[dict]:
    """Generate test data using Claude for US insurance domain"""

    client = configure_claude()

    # Build context from previously generated sheets
    previous_context = ""
    if previous_sheets_data:
        previous_context = """
CRITICAL - CROSS-SHEET DATA CONSISTENCY:
The following sheets have already been generated. You MUST reuse the SAME values for matching/similar fields across sheets.
For example, if Sheet1 has an address "123 Main St, Dallas, TX 75201", and the current sheet also has address columns, you MUST use the exact same addresses for the corresponding rows.
Row 1 in this sheet corresponds to Row 1 in all other sheets (same insured/policy/location).

Fields that MUST stay consistent across sheets include (when the column exists):
- Names (Insured Name, Contact Name, Agent Name, etc.)
- Addresses (Street, City, State, ZIP, County)
- Phone numbers, Email addresses
- Policy numbers, Account numbers
- Effective dates, Expiration dates
- SSN, FEIN, Tax ID
- Any identifier that appears in multiple sheets
- Data Set (DS_1, DS_2, etc.) must match row-for-row across all sheets

Previously generated data:
"""
        for prev_sheet_name, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- Sheet: '{prev_sheet_name}' ---\n"
            previous_context += json.dumps(prev_data, indent=2)
            previous_context += "\n"

    prompt = f"""You are generating data for sheet: "{sheet_name}"

Generate exactly {row_count} rows of realistic test data for the following columns:
{json.dumps(headers)}

DUPLICATE COLUMN HANDLING:
Some columns may have a suffix like "(Col X)" — this means the original spreadsheet has multiple columns with the same name.
You MUST generate DIFFERENT values for each such column. For example, "Address (Col 3)" and "Address (Col 7)" are two separate columns that need different address values.
Use ALL column names exactly as provided (including suffixes) as JSON keys.

{previous_context}

Make sure you follow the below important rules and {special_instruction} given by users for particular sheet and column name.
The Special Instruction should be prioritized before the important rules.

IMPORTANT RULES:
1. Generate realistic US insurance test data
2. Use valid US formats:
   - Phone: (XXX) XXX-XXXX
   - SSN: XXX-XX-XXXX (use fake but valid format)
   - ZIP: 5 digits or ZIP+4
   - State: 2-letter abbreviations (CA, TX, NY, etc.)
   - Dates: MM/DD/YYYY format
3. For policy numbers, use realistic formats used in United States.
4. For currency amounts, use realistic insurance values
5. Names should be diverse and realistic
6. Addresses should be realistic US addresses
7. For vehicles no and other vehicle detail make sure you follow the United States formats.
8. The effective date column is of today's date (2026 year) or one day after today.
9. The expiration date column is complete one year gap from the effective date.
10. All the Dollar values should be numbers and not in string. But should have a dollar sign before them.
11. The Data Set column will always start with DS_1, DS_2. The number will increase for each test case.
12. The effective date, expiration date should match across all sheets. There should not be mismatch for each test case row.
13. Row 1 in this sheet corresponds to Row 1 in all other sheets (same insured/policy/location).

Return ONLY a valid JSON array of objects with the exact column names as keys.
No markdown, no explanation, just the JSON array.

Example format:
[{{"column1": "value1", "column2": "value2"}}, ...]
"""

    # Retry up to 3 times with delay for rate limits
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=8192,
                system="You are a precise JSON-only data generator for US insurance test datasets. Return ONLY a valid JSON array. No markdown fences, no explanation, just the raw JSON array.",
                messages=[
                    {
                        "role": "user",
                        "content": prompt,
                    }
                ],
            )

            text = response.content[0].text
            if not text.strip():
                raise ValueError("Claude returned an empty response")

            # Clean up any markdown formatting
            cleaned = _clean_response_text(text)
            data = json.loads(cleaned)

            # Handle case where model returns a dict wrapper
            if isinstance(data, dict):
                for value in data.values():
                    if isinstance(value, list) and value and isinstance(value[0], dict):
                        data = value
                        break
                else:
                    data = [data]

            if not isinstance(data, list) or not data:
                raise ValueError("Model output is not a valid JSON array")

            # Ensure all rows have all requested headers
            normalized = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({header: row.get(header, "") for header in headers})
            if not normalized:
                raise ValueError("Model output did not contain valid row objects")

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

    raise RuntimeError("Claude response was not received after retries")
