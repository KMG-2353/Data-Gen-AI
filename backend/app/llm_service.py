from openai import OpenAI
import json
import os
import time
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


def configure_openai() -> OpenAI:
    """Configure OpenAI client with API key from environment."""
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise ValueError("OPENAI_API_KEY not found")
    return OpenAI(api_key=key)


def analyze_sheet_patterns(
    sheet_name: str,
    headers: list[str],
    samples: dict[str, list[Any]]
) -> dict[str, Any]:
    """Analyze sample data to extract patterns and rules for each column."""
    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "gpt-4o")

    # Build sample data description
    columns_data = []
    for header in headers:
        column_samples = samples.get(header, [])
        # Convert to strings for display
        sample_strs = [str(v) for v in column_samples[:10]]
        columns_data.append({
            "column_name": header,
            "sample_values": sample_strs,
            "sample_count": len(column_samples)
        })

    prompt = f"""Analyze the following sample data from sheet "{sheet_name}" and extract generation rules.

SAMPLE DATA:
{json.dumps(columns_data, indent=2)}

For each column, determine:
1. rule_type: One of "format", "range", "enum", "pattern", "sequence", "text"
   - "format": Data follows a specific format (phone, SSN, date, email, etc.)
   - "range": Numeric or date range
   - "enum": Fixed list of values
   - "pattern": Custom regex pattern
   - "sequence": Sequential/incrementing values (IDs, policy numbers)
   - "text": Free-form text

2. description: Human-readable rule description (e.g., "US Phone: (XXX) XXX-XXXX")

3. pattern: Regex pattern if the data follows a specific format (optional)

4. constraints: JSON object with applicable constraints:
   - For "format": {{"format_example": "..."}}
   - For "range": {{"min": X, "max": Y}}
   - For "enum": {{"values": ["A", "B", "C"]}}
   - For "date": {{"format": "MM/DD/YYYY", "range_start": "...", "range_end": "..."}}
   - For "sequence": {{"prefix": "POL-", "start": 10001, "step": 1}}

5. examples: 3 representative examples (can be from the data or generated)

6. confidence: Your confidence in this rule (0.0 to 1.0)

7. llm_reasoning: Brief explanation of why you chose this rule

Also identify any CROSS-COLUMN RULES such as:
- Date relationships (effective_date < expiration_date)
- Calculated fields (total = quantity * price)
- Conditional values

Return a JSON object with this structure:
{{
  "rules": [
    {{
      "column_name": "...",
      "rule_type": "...",
      "description": "...",
      "pattern": null,
      "constraints": {{}},
      "examples": [],
      "confidence": 0.95,
      "llm_reasoning": "..."
    }}
  ],
  "cross_column_rules": [
    "effective_date must be before expiration_date"
  ]
}}"""

    response = client.chat.completions.create(
        model=model_name,
        temperature=0.3,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": "You are an expert data analyst. Analyze sample data to extract precise patterns and rules. Return valid JSON only."
            },
            {
                "role": "user",
                "content": prompt
            },
        ],
    )

    content = response.choices[0].message.content
    text: str = content if isinstance(content, str) else ""
    if not text.strip():
        raise ValueError("OpenAI returned an empty response")

    result = json.loads(_clean_response_text(text))
    return result


def refine_rule(
    current_rule: dict[str, Any],
    user_feedback: str,
    sample_values: list[Any]
) -> dict[str, Any]:
    """Refine a rule based on user feedback using LLM."""
    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "gpt-4o")

    sample_strs = [str(v) for v in sample_values[:10]]

    prompt = f"""You are refining a data generation rule based on user feedback.

CURRENT RULE:
Column: {current_rule.get('column_name', '')}
Type: {current_rule.get('rule_type', '')}
Description: {current_rule.get('description', '')}
Pattern: {current_rule.get('pattern', '')}
Constraints: {json.dumps(current_rule.get('constraints', {}))}

USER FEEDBACK:
"{user_feedback}"

ORIGINAL SAMPLE DATA:
{json.dumps(sample_strs)}

Generate an UPDATED rule that incorporates the user's feedback while remaining consistent with the data patterns where applicable.

Return a JSON object with the SAME structure as the input rule:
{{
  "column_name": "...",
  "rule_type": "...",
  "description": "...",
  "pattern": null,
  "constraints": {{}},
  "examples": [],
  "confidence": 0.95,
  "llm_reasoning": "Explanation of what changed based on feedback"
}}"""

    response = client.chat.completions.create(
        model=model_name,
        temperature=0.3,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": "You are a data rule refinement assistant. Update rules based on user feedback. Return valid JSON only."
            },
            {
                "role": "user",
                "content": prompt
            },
        ],
    )

    content = response.choices[0].message.content
    text: str = content if isinstance(content, str) else ""
    if not text.strip():
        raise ValueError("OpenAI returned an empty response")

    return json.loads(_clean_response_text(text))

def generate_test_data(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str = "",
    previous_sheets_data: dict[str, Any] | None = None,
    rules: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Generate test data using OpenAI, optionally following verified rules."""

    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "gpt-4o")

    # Build context from previously generated sheets
    previous_context = ""
    if previous_sheets_data:
        previous_context = """
CRITICAL - CROSS-SHEET DATA CONSISTENCY:
The following sheets have already been generated. You MUST reuse the SAME values for matching/similar fields across sheets.
Row 1 in this sheet corresponds to Row 1 in all other sheets (same insured/policy/location).

Fields that MUST stay consistent across sheets include (when the column exists):
- Names, Addresses, Phone numbers, Email addresses
- Policy numbers, Account numbers
- Effective dates, Expiration dates
- SSN, FEIN, Tax ID
- Data Set (DS_1, DS_2, etc.) must match row-for-row

Previously generated data:
"""
        for prev_sheet_name, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- Sheet: '{prev_sheet_name}' ---\n"
            previous_context += json.dumps(prev_data, indent=2)
            previous_context += "\n"

    # Build rules context if rules are provided
    rules_context = ""
    if rules and "rules" in rules:
        rules_context = """
VERIFIED COLUMN RULES (You MUST follow these exactly):
"""
        for rule in rules["rules"]:
            col_name = rule.get("column_name", "")
            description = rule.get("description", "")
            pattern = rule.get("pattern", "")
            constraints = rule.get("constraints", {})

            rules_context += f"\n- {col_name}: {description}"
            if pattern:
                rules_context += f" (Pattern: {pattern})"
            if constraints:
                rules_context += f" Constraints: {json.dumps(constraints)}"

        if rules.get("cross_column_rules"):
            rules_context += "\n\nCROSS-COLUMN RULES:\n"
            for cross_rule in rules["cross_column_rules"]:
                rules_context += f"- {cross_rule}\n"

    # Build prompt based on whether rules are provided
    if rules_context:
        prompt = f"""You are a test data generator for US insurance applications.

You are generating data for sheet: "{sheet_name}"

Generate exactly {row_count} rows of test data for the following columns:
{json.dumps(headers)}

{rules_context}

{previous_context}

Additional user instructions: {special_instruction}

IMPORTANT: Follow the verified rules EXACTLY. These rules were analyzed from sample data and verified by the user.

Return a JSON object with a single key "data" whose value is an array of {row_count} objects.
Each object must use the exact column names as keys.
No markdown, no explanation, just the JSON object.

Example format:
{{"data": [{{"column1": "value1", "column2": "value2"}}, ...]}}
"""
    else:
        # Legacy mode - use default US insurance rules
        prompt = f"""You are a test data generator for US insurance applications.

You are generating data for sheet: "{sheet_name}"

Generate exactly {row_count} rows of realistic test data for the following columns:
{json.dumps(headers)}

DUPLICATE COLUMN HANDLING:
Some columns may have a suffix like "(Col X)" — this means the original spreadsheet has multiple columns with the same name.
You MUST generate DIFFERENT values for each such column.
Use ALL column names exactly as provided (including suffixes) as JSON keys.

{previous_context}

Make sure you follow these rules and {special_instruction} given by users.

IMPORTANT RULES:
1. Generate realistic US insurance test data
2. Use valid US formats:
   - Phone: (XXX) XXX-XXXX
   - SSN: XXX-XX-XXXX (use fake but valid format)
   - ZIP: 5 digits or ZIP+4
   - State: 2-letter abbreviations (CA, TX, NY, etc.)
   - Dates: MM/DD/YYYY format
3. For policy numbers, use realistic formats
4. For currency amounts, use realistic insurance values
5. Names should be diverse and realistic
6. Addresses should be realistic US addresses
7. The effective date column is of today's date (2026 year) or one day after today
8. The expiration date column is complete one year gap from the effective date
9. All Dollar values should be numbers with a dollar sign before them
10. The Data Set column will always start with DS_1, DS_2, etc.

Return a JSON object with a single key "data" whose value is an array of {row_count} objects.
Each object must use the exact column names as keys.
No markdown, no explanation, just the JSON object.

Example format:
{{"data": [{{"column1": "value1", "column2": "value2"}}, ...]}}
"""

    # Retry up to 2 times for rate limits only
    max_retries = 2
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model_name,
                temperature=0.8 if rules else 1,  # Lower temp when following rules
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "system",
                        "content": "You are a precise JSON-only data generator for US insurance test datasets. Always return a JSON object with a single key \"data\" containing an array of row objects."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    },
                ],
            )

            content = response.choices[0].message.content
            text: str = content if isinstance(content, str) else ""
            if not text.strip():
                raise ValueError("OpenAI returned an empty response")

            data = _parse_json_array(text)

            # Ensure all rows have all requested headers
            normalized: list[dict[str, Any]] = []
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

    raise RuntimeError("OpenAI response was not received after retries")

