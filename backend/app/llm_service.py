import anthropic
import json
import os
import time
from datetime import datetime, timedelta
from typing import Any

from .pattern_detector import analyze_sheet_patterns_deterministic


# Default model - Claude Sonnet 4 is excellent for pattern recognition
DEFAULT_MODEL = "claude-sonnet-4-20250514"


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


def configure_anthropic() -> anthropic.Anthropic:
    """Configure Anthropic client with API key from environment."""
    key = os.getenv("CLAUDE_API_KEY")
    if not key:
        raise ValueError("CLAUDE_API_KEY not found in environment")
    return anthropic.Anthropic(api_key=key)


def analyze_sheet_patterns(
    sheet_name: str,
    headers: list[str],
    samples: dict[str, list[Any]]
) -> dict[str, Any]:
    """
    Analyze sample data to extract patterns and rules for each column.

    Uses a HYBRID approach:
    1. First, run deterministic pattern detection (fast, accurate for known patterns)
    2. For low-confidence rules, optionally enhance with LLM analysis
    """
    # Step 1: Deterministic pattern detection
    deterministic_result = analyze_sheet_patterns_deterministic(sheet_name, headers, samples)

    # Step 2: Check if any rules need LLM enhancement (confidence < 0.7)
    low_confidence_rules = [
        r for r in deterministic_result["rules"]
        if r.get("confidence", 1.0) < 0.7 and r.get("rule_type") == "text"
    ]

    # If all rules have high confidence, return deterministic results
    if not low_confidence_rules:
        return deterministic_result

    # Step 3: Use LLM only for low-confidence columns
    try:
        client = configure_anthropic()
        model_name = os.getenv("CLAUDE_MODEL", DEFAULT_MODEL)

        # Only send low-confidence columns to LLM
        columns_for_llm = []
        for rule in low_confidence_rules:
            col_name = rule["column_name"]
            col_samples = samples.get(col_name, [])
            sample_strs = [str(v) for v in col_samples[:10] if v is not None]
            columns_for_llm.append({
                "column_name": col_name,
                "sample_values": sample_strs,
                "current_detection": rule.get("description", "Unknown")
            })

        prompt = f"""Analyze ONLY these columns that need pattern clarification:

{json.dumps(columns_for_llm, indent=2)}

For each column, determine:
1. rule_type: One of "format", "range", "enum", "pattern", "text"
2. description: Human-readable rule description
3. constraints: Applicable constraints as JSON

Return JSON:
{{"rules": [{{"column_name": "...", "rule_type": "...", "description": "...", "constraints": {{}}}}]}}"""

        response = client.messages.create(
            model=model_name,
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
            system="You are a data analyst. Analyze only ambiguous columns. Return valid JSON only.",
        )

        text = "".join(block.text for block in response.content if block.type == "text")

        if text.strip():
            llm_result = json.loads(_clean_response_text(text))

            # Merge LLM results into deterministic results
            llm_rules_map = {r["column_name"]: r for r in llm_result.get("rules", [])}

            for i, rule in enumerate(deterministic_result["rules"]):
                if rule["column_name"] in llm_rules_map:
                    llm_rule = llm_rules_map[rule["column_name"]]
                    deterministic_result["rules"][i].update({
                        "rule_type": llm_rule.get("rule_type", rule["rule_type"]),
                        "description": llm_rule.get("description", rule["description"]),
                        "constraints": llm_rule.get("constraints", rule.get("constraints", {})),
                        "detection_method": "hybrid"
                    })
    except Exception as e:
        # If LLM fails, just use deterministic results
        print(f"LLM enhancement failed, using deterministic results: {e}")

    return deterministic_result


def refine_rule(
    current_rule: dict[str, Any],
    user_feedback: str,
    sample_values: list[Any]
) -> dict[str, Any]:
    """Refine a rule based on user feedback using LLM."""
    client = configure_anthropic()
    model_name = os.getenv("CLAUDE_MODEL", DEFAULT_MODEL)

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

    response = client.messages.create(
        model=model_name,
        max_tokens=2048,
        messages=[
            {
                "role": "user",
                "content": prompt
            },
        ],
        system="You are a data rule refinement assistant. Update rules based on user feedback. Return valid JSON only.",
    )

    # Extract text from response
    text = ""
    for block in response.content:
        if block.type == "text":
            text += block.text

    if not text.strip():
        raise ValueError("Claude returned an empty response")

    return json.loads(_clean_response_text(text))


def _generate_deterministic_value(rule: dict[str, Any], row_index: int, base_date: datetime | None = None) -> Any:
    """
    Generate a value deterministically based on the rule type.
    Returns None if the value should be generated by LLM instead.
    """
    rule_type = rule.get("rule_type", "text")
    constraints = rule.get("constraints", {})

    if rule_type == "empty":
        return None  # Column should be empty

    if rule_type == "sequence":
        seq_type = constraints.get("sequence_type", "integer")
        start = constraints.get("start", 1)
        step = constraints.get("step", 1)
        prefix = constraints.get("prefix", "")
        suffix = constraints.get("suffix", "")

        if seq_type == "prefixed" or prefix:
            # Preserve leading zeros if present in original
            num_digits = constraints.get("num_digits", len(str(start)))
            value = start + (row_index * step)
            return f"{prefix}{str(value).zfill(num_digits)}{suffix}"
        elif seq_type == "date":
            if base_date:
                step_days = constraints.get("step_days", step)
                # Return datetime object - will be formatted later
                return base_date + timedelta(days=row_index * step_days)
            return None
        else:
            return start + (row_index * step)

    if rule_type == "enum":
        values = constraints.get("values", [])
        if constraints.get("constant") and values:
            return values[0]  # Always return the constant value
        elif values:
            # Cycle through values or pick based on index
            return values[row_index % len(values)]

    if rule_type == "date":
        # Generate date values - return datetime object for later formatting
        if base_date:
            step_days = constraints.get("step_days", 1)
            return base_date + timedelta(days=row_index * step_days)

    # For other types, return None to indicate LLM should generate
    return None


def _generate_row_deterministic(
    headers: list[str],
    rules: list[dict[str, Any]],
    row_index: int,
    cross_column_rules: list[dict[str, Any]] | None = None
) -> tuple[dict[str, Any], list[str]]:
    """
    Generate as much of a row as possible deterministically.
    Returns (partial_row, columns_needing_llm)
    """
    rules_map = {r["column_name"]: r for r in rules}
    row = {}
    need_llm = []

    # First pass: generate independent deterministic values
    base_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    for header in headers:
        rule = rules_map.get(header, {"rule_type": "text"})
        value = _generate_deterministic_value(rule, row_index, base_date)

        if value is not None:
            row[header] = value
        elif rule.get("rule_type") != "empty":
            need_llm.append(header)

    # Second pass: handle cross-column relationships
    if cross_column_rules:
        for cross_rule in cross_column_rules:
            if cross_rule.get("relationship") == "exactly_one_year_apart":
                cols = cross_rule.get("columns", [])
                if len(cols) == 2 and cols[0] in row:
                    start_date = row[cols[0]]
                    if isinstance(start_date, datetime):
                        # Add exactly one year
                        try:
                            row[cols[1]] = start_date.replace(year=start_date.year + 1)
                        except ValueError:
                            # Handle Feb 29 -> Feb 28
                            row[cols[1]] = start_date.replace(year=start_date.year + 1, day=28)
                        if cols[1] in need_llm:
                            need_llm.remove(cols[1])

            elif cross_rule.get("type") == "equal_values":
                cols = cross_rule.get("columns", [])
                if len(cols) == 2 and cols[0] in row and cols[1] not in row:
                    row[cols[1]] = row[cols[0]]
                    if cols[1] in need_llm:
                        need_llm.remove(cols[1])

    return row, need_llm


def generate_test_data(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str = "",
    previous_sheets_data: dict[str, Any] | None = None,
    rules: dict[str, Any] | None = None,
    samples: dict[str, list[Any]] | None = None,
) -> list[dict[str, Any]]:
    """
    Generate test data using HYBRID approach:
    1. Generate deterministic values (sequences, enums, dates, empty) without LLM
    2. Use LLM only for text/complex fields that can't be generated deterministically
    """

    # If we have rules, use hybrid generation
    if rules and "rules" in rules:
        return _generate_hybrid(
            headers=headers,
            row_count=row_count,
            special_instruction=special_instruction,
            sheet_name=sheet_name,
            previous_sheets_data=previous_sheets_data,
            rules=rules,
            samples=samples
        )

    # Legacy mode - full LLM generation
    return _generate_full_llm(
        headers=headers,
        row_count=row_count,
        special_instruction=special_instruction,
        sheet_name=sheet_name,
        previous_sheets_data=previous_sheets_data,
        samples=samples
    )


def _generate_hybrid(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    rules: dict[str, Any],
    samples: dict[str, list[Any]] | None = None
) -> list[dict[str, Any]]:
    """Hybrid generation: deterministic where possible, LLM for the rest."""

    rules_list = rules.get("rules", [])
    cross_rules = rules.get("cross_column_rules_detailed", [])

    # Build a map of column -> date format from rules
    date_format_map = {}
    for rule in rules_list:
        col_name = rule.get("column_name")
        constraints = rule.get("constraints", {})
        if rule.get("rule_type") == "date":
            # Check format in constraints
            fmt = constraints.get("output_format", constraints.get("format", "date"))
            date_format_map[col_name] = fmt

    # Step 1: Generate deterministic parts for all rows
    all_rows = []
    all_need_llm = set()

    for row_idx in range(row_count):
        partial_row, need_llm = _generate_row_deterministic(
            headers, rules_list, row_idx, cross_rules
        )
        all_rows.append(partial_row)
        all_need_llm.update(need_llm)

    # Step 2: If some columns need LLM, generate them
    if all_need_llm:
        llm_columns = list(all_need_llm)
        llm_data = _generate_llm_columns(
            columns=llm_columns,
            row_count=row_count,
            rules_list=rules_list,
            special_instruction=special_instruction,
            sheet_name=sheet_name,
            previous_sheets_data=previous_sheets_data,
            existing_rows=all_rows,
            samples=samples
        )

        # Merge LLM data into rows
        for row_idx, row in enumerate(all_rows):
            if row_idx < len(llm_data):
                for col in llm_columns:
                    if col in llm_data[row_idx]:
                        row[col] = llm_data[row_idx][col]

    # Step 3: Ensure all headers are present and format dates correctly
    normalized = []
    for row in all_rows:
        formatted_row = {}
        for header in headers:
            value = row.get(header, "")
            # Format datetime objects based on detected format
            if isinstance(value, datetime):
                # Check if this column should be date-only
                date_fmt = date_format_map.get(header, "date")
                if date_fmt in ("date", "YYYY-MM-DD") or date_fmt != "datetime":
                    # Format as date string without time
                    formatted_row[header] = value.strftime("%Y-%m-%d")
                else:
                    # Keep full datetime
                    formatted_row[header] = value.strftime("%Y-%m-%d %H:%M:%S")
            else:
                formatted_row[header] = value
        normalized.append(formatted_row)

    return normalized


def _generate_llm_columns(
    columns: list[str],
    row_count: int,
    rules_list: list[dict[str, Any]],
    special_instruction: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    existing_rows: list[dict[str, Any]],
    samples: dict[str, list[Any]] | None = None
) -> list[dict[str, Any]]:
    """Use LLM to generate only specific columns following EXACT patterns from samples."""

    client = configure_anthropic()
    model_name = os.getenv("CLAUDE_MODEL", DEFAULT_MODEL)

    # Build rules context with SAMPLE VALUES for each column
    rules_map = {r["column_name"]: r for r in rules_list}
    columns_with_samples = []

    for col in columns:
        col_info = {"column": col}

        # Get the rule
        if col in rules_map:
            rule = rules_map[col]
            col_info["rule"] = rule.get('description', 'Generate similar values')

        # Get ORIGINAL sample values - this is critical for pattern matching
        if samples and col in samples:
            sample_values = [str(v) for v in samples[col][:10] if v is not None]
            col_info["original_samples"] = sample_values

        columns_with_samples.append(col_info)

    # Build context from existing deterministic values
    existing_context = ""
    if existing_rows:
        sample_rows = existing_rows[:3]
        existing_context = f"""
ALREADY GENERATED VALUES (stay consistent with these):
{json.dumps(sample_rows, indent=2, default=str)}
"""

    # Build context from previous sheets
    previous_context = ""
    if previous_sheets_data:
        previous_context = """
CROSS-SHEET CONSISTENCY - Reuse matching field values from:
"""
        for prev_name, prev_data in previous_sheets_data.items():
            if prev_data:
                previous_context += f"\n{prev_name}: {json.dumps(prev_data[:2], indent=2, default=str)}\n"

    prompt = f"""Generate NEW, UNIQUE values for these columns following the same PATTERN and FORMAT as the original samples.

Sheet: "{sheet_name}"
Number of rows needed: {row_count}

COLUMNS TO GENERATE WITH THEIR ORIGINAL SAMPLE VALUES:
{json.dumps(columns_with_samples, indent=2)}

CRITICAL INSTRUCTIONS:
1. FOLLOW the same FORMAT/STYLE as the samples (e.g., "First Last" name format, address structure)
2. Generate DIFFERENT values - DO NOT copy the exact sample values
3. If samples show names like "John Doe, Mark Doe", generate NEW names like "Sarah Smith, Michael Chen"
4. If samples show addresses like "123 Ny Street, 11115", generate NEW addresses in the SAME format
5. Match character patterns, capitalization, spacing from samples
6. For numeric values, stay within a similar range but use different numbers
7. Ensure variety - each row should have unique values

{existing_context}
{previous_context}

Additional instructions: {special_instruction}

Return a JSON object with key "data" containing an array of {row_count} objects.
Each object should ONLY have keys for: {json.dumps(columns)}

Example: {{"data": [{{"{columns[0]}": "new value matching sample format"}}, ...]}}"""

    max_retries = 2
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model_name,
                max_tokens=8192,
                messages=[{"role": "user", "content": prompt}],
                system="You are a realistic test data generator. Generate NEW, UNIQUE values that match the FORMAT and STYLE of samples but with DIFFERENT content. Never copy exact sample values. Return valid JSON only.",
            )

            text = "".join(block.text for block in response.content if block.type == "text")

            if not text.strip():
                raise ValueError("Claude returned an empty response")

            data = _parse_json_array(text)
            return data

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "rate" in error_msg.lower():
                wait_time = 60 * (attempt + 1)
                print(f"Rate limit for '{sheet_name}'. Waiting {wait_time}s...")
                time.sleep(wait_time)
                if attempt == max_retries - 1:
                    raise
            else:
                raise

    return [{col: "" for col in columns} for _ in range(row_count)]


def _generate_full_llm(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    samples: dict[str, list[Any]] | None = None
) -> list[dict[str, Any]]:
    """Legacy full LLM generation when no rules are provided."""

    client = configure_anthropic()
    model_name = os.getenv("CLAUDE_MODEL", DEFAULT_MODEL)

    # Build sample context
    sample_context = ""
    if samples:
        sample_data = []
        for header in headers[:15]:  # Limit to first 15 columns
            if header in samples:
                sample_values = [str(v) for v in samples[header][:5] if v is not None]
                if sample_values:
                    sample_data.append({"column": header, "samples": sample_values})

        if sample_data:
            sample_context = f"""
ORIGINAL SAMPLE DATA - FOLLOW THE FORMAT/STYLE BUT GENERATE NEW VALUES:
{json.dumps(sample_data, indent=2)}

CRITICAL: Generate NEW data that follows the same FORMAT and STYLE as these samples, but with DIFFERENT values. Do NOT copy the exact sample values.
"""

    previous_context = ""
    if previous_sheets_data:
        previous_context = """
CROSS-SHEET CONSISTENCY - Reuse matching field values from:
"""
        for prev_name, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- {prev_name} ---\n{json.dumps(prev_data[:2], indent=2, default=str)}\n"

    prompt = f"""Generate {row_count} rows of NEW, UNIQUE test data following the same FORMAT and STYLE as the samples.

Sheet: "{sheet_name}"
Columns: {json.dumps(headers)}

{sample_context}
{previous_context}

Instructions: {special_instruction}

RULES:
1. Follow the same FORMAT and STYLE as samples (naming conventions, address formats, etc.)
2. Generate NEW, DIFFERENT values - DO NOT copy exact sample values
3. Ensure variety across rows - each row should have unique values
4. Match character patterns, spacing, capitalization from samples
5. Data Set: DS_01, DS_02, etc.
6. Expiration = Effective + 1 year

Return: {{"data": [{{...}}, ...]}}"""

    max_retries = 2
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model_name,
                max_tokens=8192,
                messages=[{"role": "user", "content": prompt}],
                system="You are a realistic test data generator. Generate NEW, UNIQUE values that match the FORMAT and STYLE of samples but with DIFFERENT content. Never copy exact sample values. Return valid JSON only.",
            )

            text = "".join(block.text for block in response.content if block.type == "text")

            if not text.strip():
                raise ValueError("Claude returned an empty response")

            data = _parse_json_array(text)

            normalized = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({h: row.get(h, "") for h in headers})
            return normalized if normalized else []

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "rate" in error_msg.lower():
                wait_time = 60 * (attempt + 1)
                print(f"Rate limit for '{sheet_name}'. Waiting {wait_time}s...")
                time.sleep(wait_time)
                if attempt == max_retries - 1:
                    raise
            else:
                raise

    raise RuntimeError("Claude response was not received after retries")
