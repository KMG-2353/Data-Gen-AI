import json
import re
import time
from datetime import datetime, timedelta
from typing import Any, Callable
import google.generativeai as genai

from .pattern_detector import analyze_sheet_patterns_deterministic


# Default Gemini model
DEFAULT_MODEL = "gemini-2.5-pro"


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


def configure_google(api_key: str) -> Any:
    """Configure Gemini SDK with explicit API key (no environment reads)."""
    if not api_key or not api_key.strip():
        raise ValueError("Google API key is required")
    genai.configure(api_key=api_key)
    return genai


def _google_text_response(
    client: Any,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_tokens: int,
) -> str:
    """Get text response from Gemini model generate_content."""
    model = client.GenerativeModel(model_name=model_name, system_instruction=system_prompt)
    response = model.generate_content(
        user_prompt,
        generation_config=client.GenerationConfig(max_output_tokens=max_tokens),
    )

    text = getattr(response, "text", None)
    if text:
        return str(text)

    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        parts = getattr(content, "parts", None) if content else None
        if not parts:
            continue
        collected: list[str] = []
        for part in parts:
            part_text = getattr(part, "text", None)
            if part_text:
                collected.append(str(part_text))
        if collected:
            return "".join(collected)

    return ""


def _normalize_header_name(header: str) -> str:
    """Normalize header text for fuzzy matching across sheets."""
    cleaned = re.sub(r"[^a-z0-9]", "", header.lower())
    return cleaned


def _looks_like_test_case_header(header: str) -> bool:
    """Detect likely Test Case identifier columns across naming variations."""
    normalized = _normalize_header_name(header)
    if not normalized:
        return False

    # Common variants: Test Case #, Test Case No, TestCaseNumber, etc.
    # Avoid broad matching like "test case details".
    if normalized in {"testcase", "testcaseid", "testcaseno", "testcasenum", "testcasenumber"}:
        return True
    if normalized.startswith("testcase") and any(k in normalized for k in ("id", "no", "num", "number")):
        return True

    # Support short forms like TS Number / TS No, but avoid generic "type" etc.
    if normalized.startswith("ts") and any(k in normalized for k in ("number", "num", "no", "id", "case")):
        return True

    return False


def _looks_like_identifier_value(value: str) -> bool:
    """Heuristic check for identifier-like values to avoid contaminating ID pools."""
    if not value:
        return False
    if len(value) > 64:
        return False
    if " " in value:
        return False
    if not any(ch.isdigit() for ch in value):
        return False
    return bool(re.match(r"^[A-Za-z0-9][A-Za-z0-9._/-]*$", value))


def _header_category(header: str) -> str | None:
    """Map a column header to a semantic category used for cross-sheet consistency."""
    normalized = _normalize_header_name(header)
    if not normalized:
        return None

    if _looks_like_test_case_header(header):
        return "test_case_id"

    if "policy" in normalized and any(k in normalized for k in ("id", "number", "num", "no")):
        return "policy_id"
    if "driver" in normalized and any(k in normalized for k in ("id", "number", "num", "no")):
        return "driver_id"
    if "vehicle" in normalized and any(k in normalized for k in ("id", "number", "num", "no", "vin")):
        return "vehicle_id"

    if "dateofbirth" in normalized or normalized == "dob" or "birthdate" in normalized:
        return "dob"
    if "gender" in normalized or "sex" == normalized:
        return "gender"
    if "age" == normalized:
        return "age"
    if "maritalstatus" in normalized:
        return "marital_status"
    if "relationship" in normalized or "relation" in normalized:
        return "relationship"

    if "name" in normalized:
        if any(k in normalized for k in ("driver", "insured", "customer", "applicant", "member", "person")):
            return "person_name"
        # Keep plain name columns linked across sheets too.
        if normalized == "name" or normalized.endswith("name"):
            return "person_name"

    return None


def _is_strict_reference_category(category: str) -> bool:
    """Categories that should be reused across sheets as canonical references."""
    return category in {"test_case_id", "policy_id", "driver_id", "vehicle_id"}


def _is_insurance_context(headers: list[str], sheet_name: str, special_instruction: str) -> bool:
    """Detect whether planner/validator should enforce insurance lifecycle semantics."""
    haystack = " ".join(headers + [sheet_name, special_instruction]).lower()
    signals = (
        "policy",
        "transaction",
        "coverage",
        "premium",
        "driver",
        "vehicle",
        "insured",
        "renewal",
        "endorsement",
        "cancellation",
        "reinstatement",
    )
    return sum(1 for s in signals if s in haystack) >= 2


def _transaction_cycle_from_instruction(special_instruction: str) -> list[str]:
    """Build the transaction cycle requested by the instruction text."""
    text = (special_instruction or "").lower()
    cycle = ["NB"]

    if any(k in text for k in ("endorsement", "mid-term", "mta")):
        cycle.append("ENDORSEMENT")
    if any(k in text for k in ("cancel", "cancellation", "flat cancel")):
        cycle.append("CANCELLATION")
    if any(k in text for k in ("reinstate", "reinstatement")):
        cycle.append("REINSTATEMENT")
    if "renewal" in text:
        cycle.append("RENEWAL")

    if len(cycle) == 1:
        cycle.extend(["ENDORSEMENT", "RENEWAL"])

    return cycle


def _build_insurance_scenarios(
    row_count: int,
    previous_sheets_data: dict[str, Any] | None,
    special_instruction: str,
) -> list[dict[str, Any]]:
    """Build deterministic lifecycle scenarios keyed by TestCase/Policy IDs."""
    reference_by_category = _build_reference_values_by_category(previous_sheets_data)
    case_pool = reference_by_category.get("test_case_id", [])
    policy_pool = reference_by_category.get("policy_id", [])

    cycle = _transaction_cycle_from_instruction(special_instruction)
    cycle_len = max(1, len(cycle))
    case_count = max(1, row_count // cycle_len)

    if case_pool:
        case_count = min(case_count, len(case_pool))

    base_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    scenarios: list[dict[str, Any]] = []

    for case_idx in range(case_count):
        test_case_id = case_pool[case_idx] if case_idx < len(case_pool) else f"TS-{case_idx + 1:04d}"
        policy_id = policy_pool[case_idx] if case_idx < len(policy_pool) else f"POL-{case_idx + 1:06d}"

        inception = base_date + timedelta(days=case_idx)
        expiration = inception + timedelta(days=365)
        cancellation_date: datetime | None = None

        for step, txn_type in enumerate(cycle):
            if len(scenarios) >= row_count:
                return scenarios

            if txn_type == "NB":
                txn_date = inception
                effective = inception
                exp = expiration
            elif txn_type == "ENDORSEMENT":
                txn_date = inception + timedelta(days=30)
                effective = inception
                exp = expiration
            elif txn_type == "CANCELLATION":
                txn_date = inception + timedelta(days=120)
                cancellation_date = txn_date
                effective = inception
                exp = expiration
            elif txn_type == "REINSTATEMENT":
                txn_date = (cancellation_date or inception + timedelta(days=120)) + timedelta(days=5)
                effective = inception
                exp = expiration
            elif txn_type == "RENEWAL":
                txn_date = expiration
                effective = expiration
                exp = expiration + timedelta(days=365)
            else:
                txn_date = inception + timedelta(days=step * 15)
                effective = inception
                exp = expiration

            scenarios.append(
                {
                    "test_case_id": test_case_id,
                    "policy_id": policy_id,
                    "transaction_type": txn_type,
                    "transaction_date": txn_date,
                    "effective_date": effective,
                    "expiration_date": exp,
                    "cancellation_date": cancellation_date,
                    "reinstatement_date": txn_date if txn_type == "REINSTATEMENT" else None,
                    "driver_changed": txn_type == "ENDORSEMENT",
                    "vehicle_changed": txn_type == "ENDORSEMENT",
                }
            )

    while len(scenarios) < row_count:
        scenarios.append(scenarios[len(scenarios) % max(1, len(scenarios))])

    return scenarios[:row_count]


def _header_role(header: str) -> str | None:
    """Classify a header into deterministic insurance roles."""
    normalized = _normalize_header_name(header)
    category = _header_category(header)
    if category == "test_case_id":
        return "test_case_id"
    if category == "policy_id":
        return "policy_id"

    if "transactiontype" in normalized or normalized in {"trantype", "trxtype"}:
        return "transaction_type"
    if "transactiondate" in normalized or normalized in {"trandate", "trxdate"}:
        return "transaction_date"
    if "effectivedate" in normalized or normalized in {"effdate", "inceptiondate"}:
        return "effective_date"
    if "expirationdate" in normalized or "expirydate" in normalized or normalized == "expdate":
        return "expiration_date"
    if "cancellationdate" in normalized or "canceldate" in normalized:
        return "cancellation_date"
    if "reinstatementdate" in normalized or "reinstate" in normalized:
        return "reinstatement_date"
    if "status" in normalized:
        return "policy_status"
    if "driverchange" in normalized:
        return "driver_changed"
    if "vehiclechange" in normalized:
        return "vehicle_changed"
    if "drivercount" in normalized:
        return "driver_count"
    if "vehiclecount" in normalized:
        return "vehicle_count"

    return None


def _to_date(value: Any) -> datetime | None:
    """Parse incoming value to datetime when possible."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _fmt_date(value: datetime | None) -> str:
    return value.strftime("%Y-%m-%d") if value else ""


def _apply_insurance_scenario_projection(
    rows: list[dict[str, Any]],
    headers: list[str],
    sheet_name: str,
    special_instruction: str,
    previous_sheets_data: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Project deterministic scenario values into known insurance lifecycle columns."""
    if not rows or not _is_insurance_context(headers, sheet_name, special_instruction):
        return rows

    scenarios = _build_insurance_scenarios(len(rows), previous_sheets_data, special_instruction)
    for idx, row in enumerate(rows):
        scenario = scenarios[idx]
        for header in headers:
            role = _header_role(header)
            if role is None:
                continue

            if role in {"test_case_id", "policy_id", "transaction_type"}:
                row[header] = scenario[role]
            elif role in {
                "transaction_date",
                "effective_date",
                "expiration_date",
                "cancellation_date",
                "reinstatement_date",
            }:
                row[header] = _fmt_date(scenario.get(role))
            elif role in {"driver_changed", "vehicle_changed"}:
                row[header] = "Y" if scenario[role] else "N"
            elif role == "policy_status":
                row[header] = "Cancelled" if scenario["transaction_type"] == "CANCELLATION" else "Active"
            elif role == "driver_count":
                row[header] = 2 if scenario["driver_changed"] else 1
            elif role == "vehicle_count":
                row[header] = 2 if scenario["vehicle_changed"] else 1

    return rows


def _validate_and_repair_insurance_rows(rows: list[dict[str, Any]], headers: list[str]) -> list[dict[str, Any]]:
    """Enforce deterministic temporal and status consistency for insurance rows."""
    if not rows:
        return rows

    for row in rows:
        role_to_header: dict[str, str] = {}
        for header in headers:
            role = _header_role(header)
            if role:
                role_to_header[role] = header

        txn_type = str(row.get(role_to_header.get("transaction_type", ""), "")).strip().upper()

        eff = _to_date(row.get(role_to_header.get("effective_date", "")))
        exp = _to_date(row.get(role_to_header.get("expiration_date", "")))
        txn = _to_date(row.get(role_to_header.get("transaction_date", "")))
        cancel = _to_date(row.get(role_to_header.get("cancellation_date", "")))
        rein = _to_date(row.get(role_to_header.get("reinstatement_date", "")))

        if eff and not exp:
            exp = eff + timedelta(days=365)

        if eff and exp and (exp - eff).days < 300:
            exp = eff + timedelta(days=365)

        if txn_type == "CANCELLATION":
            if not cancel:
                cancel = txn or (eff + timedelta(days=120) if eff else None)
            if txn and cancel and cancel < txn:
                cancel = txn

        if txn_type == "REINSTATEMENT":
            if not cancel and eff:
                cancel = eff + timedelta(days=120)
            if not rein and cancel:
                rein = cancel + timedelta(days=5)
            if rein and cancel and rein <= cancel:
                rein = cancel + timedelta(days=1)

        if role_to_header.get("effective_date"):
            row[role_to_header["effective_date"]] = _fmt_date(eff)
        if role_to_header.get("expiration_date"):
            row[role_to_header["expiration_date"]] = _fmt_date(exp)
        if role_to_header.get("transaction_date") and txn:
            row[role_to_header["transaction_date"]] = _fmt_date(txn)
        if role_to_header.get("cancellation_date"):
            row[role_to_header["cancellation_date"]] = _fmt_date(cancel)
        if role_to_header.get("reinstatement_date"):
            row[role_to_header["reinstatement_date"]] = _fmt_date(rein)

        if role_to_header.get("driver_changed"):
            row[role_to_header["driver_changed"]] = "Y" if txn_type == "ENDORSEMENT" else "N"
        if role_to_header.get("vehicle_changed"):
            row[role_to_header["vehicle_changed"]] = "Y" if txn_type == "ENDORSEMENT" else "N"
        if role_to_header.get("policy_status"):
            row[role_to_header["policy_status"]] = "Cancelled" if txn_type == "CANCELLATION" else "Active"

    return rows


def _extract_reference_values_from_rows(
    rows: list[dict[str, Any]],
    headers_filter: Callable[[str], bool],
) -> list[str]:
    """Extract unique reference values from matching headers while preserving order."""
    seen: set[str] = set()
    values: list[str] = []

    for row in rows:
        for key, value in row.items():
            if not headers_filter(key):
                continue

            if value is None:
                continue
            value_str = str(value).strip()
            if not value_str:
                continue

            if value_str not in seen:
                seen.add(value_str)
                values.append(value_str)

    return values


def _build_reference_values_by_category(previous_sheets_data: dict[str, Any] | None) -> dict[str, list[str]]:
    """Collect unique cross-sheet values grouped by semantic category."""
    if not previous_sheets_data:
        return {}

    by_category: dict[str, list[str]] = {}
    seen_by_category: dict[str, set[str]] = {}

    for _, prev_rows in previous_sheets_data.items():
        if not isinstance(prev_rows, list):
            continue

        for row in prev_rows:
            if not isinstance(row, dict):
                continue
            for key, value in row.items():
                category = _header_category(key)
                if not category:
                    continue
                if value is None:
                    continue

                value_str = str(value).strip()
                if not value_str:
                    continue

                if _is_strict_reference_category(category) and not _looks_like_identifier_value(value_str):
                    continue

                if category not in by_category:
                    by_category[category] = []
                    seen_by_category[category] = set()

                if value_str not in seen_by_category[category]:
                    seen_by_category[category].add(value_str)
                    by_category[category].append(value_str)

    return by_category


def _build_test_case_reference_pool(previous_sheets_data: dict[str, Any] | None) -> list[str]:
    """Build cross-sheet pool of canonical test case IDs from already generated sheets."""
    if not previous_sheets_data:
        return []

    pool: list[str] = []
    seen: set[str] = set()

    for _, prev_rows in previous_sheets_data.items():
        if not isinstance(prev_rows, list):
            continue

        extracted = _extract_reference_values_from_rows(prev_rows, _looks_like_test_case_header)
        for value in extracted:
            if value not in seen:
                seen.add(value)
                pool.append(value)

    return pool


def _build_anchor_profiles(previous_sheets_data: dict[str, Any] | None) -> dict[str, dict[str, str]]:
    """Build stable entity profiles keyed by anchor IDs (test_case_id, policy_id)."""
    if not previous_sheets_data:
        return {}

    profiles: dict[str, dict[str, str]] = {}

    for _, prev_rows in previous_sheets_data.items():
        if not isinstance(prev_rows, list):
            continue

        for row in prev_rows:
            if not isinstance(row, dict):
                continue

            categorized: dict[str, str] = {}
            for key, value in row.items():
                category = _header_category(key)
                if not category or value is None:
                    continue
                value_str = str(value).strip()
                if value_str:
                    categorized[category] = value_str

            anchor = categorized.get("test_case_id") or categorized.get("policy_id")
            if not anchor:
                continue

            existing = profiles.get(anchor, {})
            for cat, val in categorized.items():
                if cat not in existing:
                    existing[cat] = val
            profiles[anchor] = existing

    return profiles


def _enforce_test_case_consistency(
    rows: list[dict[str, Any]],
    headers: list[str],
    previous_sheets_data: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Force current sheet test-case-like columns to reuse IDs from prior sheets."""
    if not rows or not previous_sheets_data:
        return rows

    pool = _build_test_case_reference_pool(previous_sheets_data)
    if not pool:
        return rows

    target_headers = [h for h in headers if _looks_like_test_case_header(h)]
    if not target_headers:
        return rows

    for idx, row in enumerate(rows):
        # Reuse prior IDs cyclically to match requested row count.
        ref_id = pool[idx % len(pool)]
        for header in target_headers:
            row[header] = ref_id

    return rows


def _enforce_cross_sheet_consistency(
    rows: list[dict[str, Any]],
    headers: list[str],
    previous_sheets_data: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Apply deterministic cross-sheet consistency for reference IDs and anchored attributes."""
    if not rows or not previous_sheets_data:
        return rows

    reference_by_category = _build_reference_values_by_category(previous_sheets_data)
    if not reference_by_category:
        return rows

    headers_by_category: dict[str, list[str]] = {}
    for header in headers:
        category = _header_category(header)
        if not category:
            continue
        headers_by_category.setdefault(category, []).append(header)

    # Step 1: Enforce strict reference categories (IDs) from previous sheets.
    for category, category_headers in headers_by_category.items():
        if not _is_strict_reference_category(category):
            continue

        pool = reference_by_category.get(category, [])
        if not pool:
            continue

        for idx, row in enumerate(rows):
            canonical_value = pool[idx % len(pool)]
            for header in category_headers:
                row[header] = canonical_value

    # Step 2: Enforce anchored profile consistency for person attributes.
    profiles = _build_anchor_profiles(previous_sheets_data)
    if not profiles:
        return rows

    for row in rows:
        # Find current row anchor from any matching header.
        anchor_value = ""
        for key, value in row.items():
            cat = _header_category(key)
            if cat in ("test_case_id", "policy_id") and value is not None:
                val = str(value).strip()
                if val:
                    anchor_value = val
                    break

        if not anchor_value:
            continue

        profile = profiles.get(anchor_value)
        if not profile:
            continue

        for key in list(row.keys()):
            cat = _header_category(key)
            if not cat:
                continue
            if cat in profile and cat not in ("test_case_id", "policy_id"):
                row[key] = profile[cat]

    return rows


def analyze_sheet_patterns(
    sheet_name: str,
    headers: list[str],
    samples: dict[str, list[Any]],
    special_instruction: str = "",
    api_key: str | None = None,
    model_name: str = DEFAULT_MODEL,
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

    # Step 3: Use LLM only for low-confidence columns when API key is available
    if not api_key or not api_key.strip():
        return deterministic_result

    try:
        client = configure_google(api_key)

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

        instruction_context = ""
        if special_instruction and special_instruction.strip():
            instruction_context = f"""
    SPECIAL INSTRUCTION CONTEXT (use this to infer domain meaning for ambiguous text fields):
    {special_instruction}
    """

        prompt = f"""Analyze ONLY these columns that need pattern clarification:

{json.dumps(columns_for_llm, indent=2)}
    {instruction_context}

For each column, determine:
1. rule_type: One of "format", "range", "enum", "pattern", "text"
2. description: Human-readable rule description
3. constraints: Applicable constraints as JSON

Return JSON:
{{"rules": [{{"column_name": "...", "rule_type": "...", "description": "...", "constraints": {{}}}}]}}"""

        text = _google_text_response(
            client=client,
            model_name=model_name,
            system_prompt="You are a data analyst. Analyze only ambiguous columns. Return valid JSON only.",
            user_prompt=prompt,
            max_tokens=2048,
        )

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
    sample_values: list[Any],
    api_key: str,
    model_name: str = DEFAULT_MODEL,
) -> dict[str, Any]:
    """Refine a rule based on user feedback using LLM."""
    client = configure_google(api_key)

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

    text = _google_text_response(
        client=client,
        model_name=model_name,
        system_prompt="You are a data rule refinement assistant. Update rules based on user feedback. Return valid JSON only.",
        user_prompt=prompt,
        max_tokens=2048,
    )

    if not text.strip():
        raise ValueError("Model returned an empty response")

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
    api_key: str,
    model_name: str = DEFAULT_MODEL,
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

    # If we have rules, always use hybrid generation so deterministic pattern
    # constraints remain grounded while LLM fills only complex text fields.
    if rules and "rules" in rules:
        return _generate_hybrid(
            headers=headers,
            row_count=row_count,
            special_instruction=special_instruction,
            api_key=api_key,
            model_name=model_name,
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
        api_key=api_key,
        model_name=model_name,
        sheet_name=sheet_name,
        previous_sheets_data=previous_sheets_data,
        samples=samples
    )


def _generate_instruction_first_with_rules(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    api_key: str,
    model_name: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    rules: dict[str, Any],
    samples: dict[str, list[Any]] | None = None,
) -> list[dict[str, Any]]:
    """Generate rows with explicit priority: special instruction first, rules second."""

    client = configure_google(api_key)

    rules_list = rules.get("rules", [])
    cross_rules = rules.get("cross_column_rules_detailed", [])

    rules_summary = []
    for r in rules_list:
        rules_summary.append(
            {
                "column_name": r.get("column_name", ""),
                "rule_type": r.get("rule_type", "text"),
                "description": r.get("description", ""),
                "constraints": r.get("constraints", {}),
                "examples": r.get("examples", [])[:3],
            }
        )

    sample_context = ""
    if samples:
        sample_data = []
        for header in headers[:20]:
            if header in samples:
                sample_values = [str(v) for v in samples[header][:5] if v is not None]
                if sample_values:
                    sample_data.append({"column": header, "samples": sample_values})
        if sample_data:
            sample_context = f"""
ORIGINAL SAMPLE STYLE REFERENCE:
{json.dumps(sample_data, indent=2)}
"""

    previous_context = ""
    reference_pool = _build_test_case_reference_pool(previous_sheets_data)
    reference_by_category = _build_reference_values_by_category(previous_sheets_data)
    anchor_profiles = _build_anchor_profiles(previous_sheets_data)
    if previous_sheets_data:
        previous_context = "\nCROSS-SHEET CONTEXT (reuse related entities where sensible):\n"
        for prev_name, prev_data in previous_sheets_data.items():
            if prev_data:
                previous_context += f"\n--- {prev_name} ---\n{json.dumps(prev_data[:2], indent=2, default=str)}\n"

    reference_context = ""
    if reference_pool:
        reference_context = f"""
CROSS-SHEET REFERENCE IDS (MUST REUSE for test case columns):
{json.dumps(reference_pool, indent=2)}
"""

    semantic_reference_context = ""
    if reference_by_category:
        semantic_reference_context = f"""
CROSS-SHEET REFERENCE VALUES BY ENTITY (MUST STAY CONSISTENT):
{json.dumps(reference_by_category, indent=2)}
"""

    anchor_profile_context = ""
    if anchor_profiles:
        sample_profiles = dict(list(anchor_profiles.items())[:20])
        anchor_profile_context = f"""
ANCHORED ENTITY PROFILES (for same TestCase/Policy IDs, keep attributes consistent):
{json.dumps(sample_profiles, indent=2)}
"""

    prompt = f"""Generate {row_count} rows of realistic synthetic data.

Sheet: "{sheet_name}"
Columns: {json.dumps(headers)}

SPECIAL INSTRUCTION (HIGHEST PRIORITY):
{special_instruction}

DETECTED RULES (SECOND PRIORITY, use when not conflicting with special instruction):
{json.dumps(rules_summary, indent=2)}

CROSS-COLUMN RULES (SECOND PRIORITY):
{json.dumps(cross_rules, indent=2)}

{sample_context}
{previous_context}
{reference_context}
{semantic_reference_context}
{anchor_profile_context}

PRIORITY ORDER (MANDATORY):
1. Follow SPECIAL INSTRUCTION first.
2. Follow detected rules/patterns where they do not conflict with the special instruction.
3. Match sample style/format while producing NEW values (never copy exact samples).
4. For test-case identifier columns (e.g., Test Case # / Test Case No), reuse IDs from CROSS-SHEET REFERENCE IDS.
5. Keep entity consistency across sheets: same Test Case / Policy should map to same person attributes where applicable.

OUTPUT REQUIREMENTS:
1. Return valid JSON only.
2. Output object must be: {{"data": [{{...}}, ...]}}.
3. Include exactly these columns in every row: {json.dumps(headers)}.
4. Produce exactly {row_count} rows.
"""

    max_retries = 2
    for attempt in range(max_retries):
        try:
            text = _google_text_response(
                client=client,
                model_name=model_name,
                system_prompt="You are a synthetic data generator. Respect instruction priority strictly: special instruction first, detected rules second. Return valid JSON only.",
                user_prompt=prompt,
                max_tokens=8192,
            )

            if not text.strip():
                raise ValueError("Model returned an empty response")

            data = _parse_json_array(text)

            normalized = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({h: row.get(h, "") for h in headers})

            if len(normalized) > row_count:
                normalized = normalized[:row_count]
            elif len(normalized) < row_count:
                normalized.extend([{h: "" for h in headers} for _ in range(row_count - len(normalized))])

            normalized = _enforce_cross_sheet_consistency(
                rows=normalized,
                headers=headers,
                previous_sheets_data=previous_sheets_data,
            )

            normalized = _apply_insurance_scenario_projection(
                rows=normalized,
                headers=headers,
                sheet_name=sheet_name,
                special_instruction=special_instruction,
                previous_sheets_data=previous_sheets_data,
            )

            normalized = _validate_and_repair_insurance_rows(normalized, headers)

            return normalized

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


def _generate_hybrid(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    api_key: str,
    model_name: str,
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
            cross_rules=cross_rules,
            special_instruction=special_instruction,
            api_key=api_key,
            model_name=model_name,
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

    normalized = _enforce_cross_sheet_consistency(
        rows=normalized,
        headers=headers,
        previous_sheets_data=previous_sheets_data,
    )

    normalized = _apply_insurance_scenario_projection(
        rows=normalized,
        headers=headers,
        sheet_name=sheet_name,
        special_instruction=special_instruction,
        previous_sheets_data=previous_sheets_data,
    )

    return _validate_and_repair_insurance_rows(normalized, headers)


def _generate_llm_columns(
    columns: list[str],
    row_count: int,
    rules_list: list[dict[str, Any]],
    cross_rules: list[dict[str, Any]] | None,
    special_instruction: str,
    api_key: str,
    model_name: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    existing_rows: list[dict[str, Any]],
    samples: dict[str, list[Any]] | None = None
) -> list[dict[str, Any]]:
    """Use LLM to generate only specific columns following EXACT patterns from samples."""

    client = configure_google(api_key)

    # Build rules context with constraints and examples for each LLM column
    rules_map = {r["column_name"]: r for r in rules_list}
    columns_with_samples: list[dict[str, Any]] = []
    llm_column_rules: list[dict[str, Any]] = []

    for col in columns:
        col_info: dict[str, Any] = {"column": col}

        # Get the rule
        if col in rules_map:
            rule = rules_map[col]
            col_info["rule"] = rule.get('description', 'Generate similar values')
            llm_column_rules.append(
                {
                    "column_name": col,
                    "rule_type": rule.get("rule_type", "text"),
                    "description": rule.get("description", ""),
                    "constraints": rule.get("constraints", {}),
                    "examples": rule.get("examples", [])[:3],
                }
            )

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
    reference_pool = _build_test_case_reference_pool(previous_sheets_data)
    reference_by_category = _build_reference_values_by_category(previous_sheets_data)
    anchor_profiles = _build_anchor_profiles(previous_sheets_data)
    if previous_sheets_data:
        previous_context = """
CROSS-SHEET CONSISTENCY - Reuse matching field values from:
"""
        for prev_name, prev_data in previous_sheets_data.items():
            if prev_data:
                previous_context += f"\n{prev_name}: {json.dumps(prev_data[:2], indent=2, default=str)}\n"

    reference_context = ""
    if reference_pool:
        reference_context = f"""
MANDATORY TEST CASE ID REUSE:
If current columns include any test-case identifier field, values must come from this list:
{json.dumps(reference_pool, indent=2)}
"""

    semantic_reference_context = ""
    if reference_by_category:
        semantic_reference_context = f"""
REFERENCE VALUES BY ENTITY CATEGORY (MUST STAY CONSISTENT WHERE APPLICABLE):
{json.dumps(reference_by_category, indent=2)}
"""

    anchor_profile_context = ""
    if anchor_profiles:
        sample_profiles = dict(list(anchor_profiles.items())[:20])
        anchor_profile_context = f"""
ANCHORED ENTITY PROFILES (same TestCase/Policy -> same person attributes):
{json.dumps(sample_profiles, indent=2)}
"""

    prompt = f"""Generate NEW, UNIQUE values for these columns following special instruction and analyzed rules.

Sheet: "{sheet_name}"
Number of rows needed: {row_count}

SPECIAL INSTRUCTION (HIGHEST PRIORITY):
{special_instruction}

COLUMNS TO GENERATE WITH THEIR ORIGINAL SAMPLE VALUES:
{json.dumps(columns_with_samples, indent=2)}

ANALYZED RULES FOR THESE COLUMNS:
{json.dumps(llm_column_rules, indent=2)}

CROSS-COLUMN RULES:
{json.dumps(cross_rules or [], indent=2)}

CRITICAL INSTRUCTIONS:
1. Follow SPECIAL INSTRUCTION first.
2. Follow ANALYZED RULES and CROSS-COLUMN RULES for these columns.
3. Match sample format/style while generating NEW values.
4. Do not copy exact sample values.
5. Keep row-to-row data coherent with already generated deterministic fields.

{existing_context}
{previous_context}
{reference_context}
{semantic_reference_context}
{anchor_profile_context}

Return a JSON object with key "data" containing an array of {row_count} objects.
Each object should ONLY have keys for: {json.dumps(columns)}

Example: {{"data": [{{"{columns[0]}": "new value matching sample format"}}, ...]}}"""

    max_retries = 2
    for attempt in range(max_retries):
        try:
            text = _google_text_response(
                client=client,
                model_name=model_name,
                system_prompt="You are a strict synthetic data generator. Obey priority order: special instruction first, then analyzed rules/cross-rules, then sample style. Return valid JSON only.",
                user_prompt=prompt,
                max_tokens=8192,
            )

            if not text.strip():
                raise ValueError("Model returned an empty response")

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
    api_key: str,
    model_name: str,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    samples: dict[str, list[Any]] | None = None
) -> list[dict[str, Any]]:
    """Legacy full LLM generation when no rules are provided."""

    client = configure_google(api_key)

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
    reference_pool = _build_test_case_reference_pool(previous_sheets_data)
    reference_by_category = _build_reference_values_by_category(previous_sheets_data)
    anchor_profiles = _build_anchor_profiles(previous_sheets_data)
    if previous_sheets_data:
        previous_context = """
CROSS-SHEET CONSISTENCY - Reuse matching field values from:
"""
        for prev_name, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- {prev_name} ---\n{json.dumps(prev_data[:2], indent=2, default=str)}\n"

    reference_context = ""
    if reference_pool:
        reference_context = f"""
MANDATORY TEST CASE ID REUSE:
If this sheet has any test-case identifier column, values must be selected from:
{json.dumps(reference_pool, indent=2)}
"""

    semantic_reference_context = ""
    if reference_by_category:
        semantic_reference_context = f"""
REFERENCE VALUES BY ENTITY CATEGORY (MUST STAY CONSISTENT WHERE APPLICABLE):
{json.dumps(reference_by_category, indent=2)}
"""

    anchor_profile_context = ""
    if anchor_profiles:
        sample_profiles = dict(list(anchor_profiles.items())[:20])
        anchor_profile_context = f"""
ANCHORED ENTITY PROFILES (same TestCase/Policy -> same person attributes):
{json.dumps(sample_profiles, indent=2)}
"""

    prompt = f"""Generate {row_count} rows of NEW, UNIQUE test data following the same FORMAT and STYLE as the samples.

Sheet: "{sheet_name}"
Columns: {json.dumps(headers)}

{sample_context}
{previous_context}
{reference_context}
{semantic_reference_context}
{anchor_profile_context}

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
            text = _google_text_response(
                client=client,
                model_name=model_name,
                system_prompt="You are a realistic test data generator. Generate NEW, UNIQUE values that match the FORMAT and STYLE of samples but with DIFFERENT content. Never copy exact sample values. Return valid JSON only.",
                user_prompt=prompt,
                max_tokens=8192,
            )

            if not text.strip():
                raise ValueError("Model returned an empty response")

            data = _parse_json_array(text)

            normalized = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({h: row.get(h, "") for h in headers})
            if not normalized:
                return []

            normalized = _enforce_cross_sheet_consistency(
                rows=normalized,
                headers=headers,
                previous_sheets_data=previous_sheets_data,
            )

            normalized = _apply_insurance_scenario_projection(
                rows=normalized,
                headers=headers,
                sheet_name=sheet_name,
                special_instruction=special_instruction,
                previous_sheets_data=previous_sheets_data,
            )

            return _validate_and_repair_insurance_rows(normalized, headers)

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

    raise RuntimeError("Google response was not received after retries")
