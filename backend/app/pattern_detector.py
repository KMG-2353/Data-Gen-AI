"""
Deterministic Pattern Detection Module

This module provides accurate pattern detection using statistical analysis
instead of relying solely on LLM inference. It detects:
- Data types (int, float, datetime, string)
- Sequences (incrementing numbers, dates)
- Enums (fixed set of values)
- Empty/null columns
- Common formats (phone, SSN, email, ZIP, policy numbers)
- Cross-column relationships
"""

import re
from datetime import datetime, timedelta
from typing import Any
from collections import Counter


# Common format patterns with regex
FORMAT_PATTERNS = {
    "us_phone": {
        "regex": r"^\(\d{3}\)\s?\d{3}-\d{4}$",
        "description": "US Phone: (XXX) XXX-XXXX",
        "example": "(555) 123-4567"
    },
    "us_phone_dots": {
        "regex": r"^\d{3}\.\d{3}\.\d{4}$",
        "description": "US Phone: XXX.XXX.XXXX",
        "example": "555.123.4567"
    },
    "ssn": {
        "regex": r"^\d{3}-\d{2}-\d{4}$",
        "description": "SSN: XXX-XX-XXXX",
        "example": "123-45-6789"
    },
    "email": {
        "regex": r"^[\w.-]+@[\w.-]+\.\w{2,}$",
        "description": "Email address",
        "example": "user@example.com"
    },
    "zip5": {
        "regex": r"^\d{5}$",
        "description": "US ZIP Code (5 digits)",
        "example": "12345"
    },
    "zip9": {
        "regex": r"^\d{5}-\d{4}$",
        "description": "US ZIP+4 Code",
        "example": "12345-6789"
    },
    "date_mdy_slash": {
        "regex": r"^\d{1,2}/\d{1,2}/\d{4}$",
        "description": "Date: MM/DD/YYYY",
        "example": "01/15/2025"
    },
    "date_ymd_dash": {
        "regex": r"^\d{4}-\d{2}-\d{2}$",
        "description": "Date: YYYY-MM-DD",
        "example": "2025-01-15"
    },
    "currency_dollar": {
        "regex": r"^\$[\d,]+(\.\d{2})?$",
        "description": "Currency: $X,XXX.XX",
        "example": "$1,234.56"
    },
    "percentage": {
        "regex": r"^\d+(\.\d+)?%$",
        "description": "Percentage: X.X%",
        "example": "15.5%"
    },
    "us_state_abbrev": {
        "regex": r"^(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)$",
        "description": "US State Abbreviation",
        "example": "CA"
    },
}


def detect_date_format(values: list[Any]) -> str | None:
    """Detect the format of date values from samples."""
    for val in values:
        if val is None:
            continue

        # Check string representations
        val_str = str(val).strip()

        # If it's a datetime object, check if original string had time
        if isinstance(val, datetime):
            # Check if the string representation has time (not midnight)
            if val.hour != 0 or val.minute != 0 or val.second != 0:
                return "datetime"  # Has meaningful time
            else:
                return "date"  # Date only (midnight time)

        # Check string format patterns
        if re.match(r"^\d{4}-\d{2}-\d{2}$", val_str):
            return "YYYY-MM-DD"
        elif re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", val_str):
            return "MM/DD/YYYY"
        elif re.match(r"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$", val_str):
            return "datetime"
        elif re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", val_str):
            return "datetime_iso"

    return "date"  # Default to date-only


def detect_data_type(values: list[Any]) -> dict[str, Any]:
    """Detect the primary data type of a column."""
    non_null = [v for v in values if v is not None and str(v).strip() != ""]

    if not non_null:
        return {"type": "empty", "nullable": True, "all_null": True}

    type_counts = Counter()
    date_format = None

    for val in non_null:
        if isinstance(val, datetime):
            type_counts["datetime"] += 1
            if date_format is None:
                date_format = detect_date_format([val])
        elif isinstance(val, bool):
            type_counts["bool"] += 1
        elif isinstance(val, int):
            type_counts["int"] += 1
        elif isinstance(val, float):
            type_counts["float"] += 1
        elif isinstance(val, str):
            # Try to infer more specific types from strings
            val_str = val.strip()
            if val_str.lower() in ("yes", "no", "true", "false", "y", "n"):
                type_counts["bool_string"] += 1
            elif re.match(r"^-?\d+$", val_str):
                type_counts["int_string"] += 1
            elif re.match(r"^-?\d+\.\d+$", val_str):
                type_counts["float_string"] += 1
            else:
                type_counts["string"] += 1
        else:
            type_counts["string"] += 1

    # Determine primary type (most common)
    if type_counts:
        primary_type = type_counts.most_common(1)[0][0]
    else:
        primary_type = "string"

    result = {
        "type": primary_type,
        "nullable": len(non_null) < len(values),
        "null_count": len(values) - len(non_null),
        "type_distribution": dict(type_counts)
    }

    # Add date format if detected
    if date_format:
        result["date_format"] = date_format

    return result


def detect_sequence(values: list[Any]) -> dict[str, Any] | None:
    """Detect if values form a sequence (incrementing numbers or dates)."""
    non_null = [v for v in values if v is not None]

    if len(non_null) < 3:
        return None

    # Check for integer sequence
    if all(isinstance(v, (int, float)) or (isinstance(v, str) and re.match(r"^-?\d+$", str(v).strip())) for v in non_null):
        nums = [int(v) if isinstance(v, (int, float)) else int(str(v).strip()) for v in non_null]
        diffs = [nums[i+1] - nums[i] for i in range(len(nums)-1)]

        if len(set(diffs)) == 1 and diffs[0] != 0:
            return {
                "detected": True,
                "sequence_type": "integer",
                "start": nums[0],
                "step": diffs[0],
                "description": f"Sequential integers starting at {nums[0]}, step {diffs[0]}"
            }

    # Check for date sequence
    if all(isinstance(v, datetime) for v in non_null):
        dates = sorted(non_null)
        diffs = [(dates[i+1] - dates[i]).days for i in range(len(dates)-1)]

        if len(set(diffs)) == 1:
            return {
                "detected": True,
                "sequence_type": "date",
                "start": dates[0].isoformat(),
                "step_days": diffs[0],
                "description": f"Sequential dates, {diffs[0]} day(s) apart"
            }

    # Check for prefixed sequence (e.g., DS_01, DS_02, POL-1001, POL-1002)
    if all(isinstance(v, str) for v in non_null):
        # Try to extract prefix and number
        pattern = r"^([A-Za-z_-]+)(\d+)([A-Za-z_-]*)$"
        matches = [re.match(pattern, str(v).strip()) for v in non_null]

        if all(m is not None for m in matches):
            prefixes = [m.group(1) for m in matches]
            numbers = [int(m.group(2)) for m in matches]
            suffixes = [m.group(3) for m in matches]

            # Check if prefix and suffix are consistent
            if len(set(prefixes)) == 1 and len(set(suffixes)) == 1:
                diffs = [numbers[i+1] - numbers[i] for i in range(len(numbers)-1)]

                if len(set(diffs)) == 1:
                    return {
                        "detected": True,
                        "sequence_type": "prefixed",
                        "prefix": prefixes[0],
                        "suffix": suffixes[0],
                        "start": numbers[0],
                        "step": diffs[0],
                        "num_digits": len(str(numbers[0])),
                        "description": f"Prefixed sequence: {prefixes[0]}XXXX{suffixes[0]}"
                    }

    return None


def detect_enum(values: list[Any], threshold: float = 0.8) -> dict[str, Any] | None:
    """Detect if values are from a fixed set (enum)."""
    non_null = [v for v in values if v is not None and str(v).strip() != ""]

    if not non_null:
        return None

    unique_values = set(str(v).strip() for v in non_null)

    # If few unique values relative to total, likely an enum
    # Also check absolute count - more than 10 unique values is probably not an enum
    if len(unique_values) <= 10 and len(unique_values) / len(non_null) < 0.5:
        value_counts = Counter(str(v).strip() for v in non_null)

        return {
            "detected": True,
            "values": sorted(list(unique_values)),
            "value_counts": dict(value_counts),
            "description": f"Enum with {len(unique_values)} values: {sorted(list(unique_values))[:5]}{'...' if len(unique_values) > 5 else ''}"
        }

    # Special case: single constant value
    if len(unique_values) == 1:
        return {
            "detected": True,
            "constant": True,
            "values": list(unique_values),
            "description": f"Constant value: {list(unique_values)[0]}"
        }

    return None


def detect_format(values: list[Any]) -> dict[str, Any] | None:
    """Detect if values match a known format pattern."""
    non_null = [str(v).strip() for v in values if v is not None and str(v).strip() != ""]

    if not non_null:
        return None

    for format_name, format_info in FORMAT_PATTERNS.items():
        pattern = format_info["regex"]
        matches = sum(1 for v in non_null if re.match(pattern, v, re.IGNORECASE))
        match_ratio = matches / len(non_null)

        if match_ratio >= 0.8:  # 80% threshold
            return {
                "detected": True,
                "format_name": format_name,
                "pattern": pattern,
                "description": format_info["description"],
                "example": format_info["example"],
                "confidence": match_ratio
            }

    # Try to detect custom patterns (like policy numbers)
    # Look for consistent structure: prefix + number + suffix
    if len(non_null) >= 3:
        # Check for patterns like "PKGN1000001-00"
        custom_pattern = detect_custom_pattern(non_null)
        if custom_pattern:
            return custom_pattern

    return None


def detect_custom_pattern(values: list[str]) -> dict[str, Any] | None:
    """Try to detect custom alphanumeric patterns."""
    if not values:
        return None

    # Analyze structure by character type
    def get_structure(s: str) -> str:
        result = []
        for c in s:
            if c.isalpha():
                result.append("A")
            elif c.isdigit():
                result.append("D")
            else:
                result.append(c)
        return "".join(result)

    structures = [get_structure(v) for v in values]

    # If all values have the same structure
    if len(set(structures)) == 1:
        structure = structures[0]

        # Simplify structure (collapse repeated chars)
        simplified = re.sub(r"A+", "A+", re.sub(r"D+", "D+", structure))

        return {
            "detected": True,
            "format_name": "custom",
            "structure": structure,
            "simplified": simplified,
            "description": f"Custom pattern: {values[0][:20]}",
            "example": values[0],
            "confidence": 1.0
        }

    return None


def detect_cross_column_rules(
    headers: list[str],
    samples: dict[str, list[Any]]
) -> list[dict[str, Any]]:
    """Detect relationships between columns."""
    rules = []

    # Find date columns
    date_columns = []
    for header in headers:
        values = samples.get(header, [])
        non_null = [v for v in values if v is not None]
        if non_null and all(isinstance(v, datetime) for v in non_null):
            date_columns.append(header)

    # Check for date relationships
    for i, col1 in enumerate(date_columns):
        for col2 in date_columns[i+1:]:
            vals1 = [v for v in samples.get(col1, []) if v is not None]
            vals2 = [v for v in samples.get(col2, []) if v is not None]

            if len(vals1) == len(vals2) and vals1 and vals2:
                # Check if one is always before the other
                diffs = [(v2 - v1).days for v1, v2 in zip(vals1, vals2)]

                if len(set(diffs)) == 1:
                    diff_days = diffs[0]
                    if diff_days == 365 or diff_days == 366:
                        rules.append({
                            "type": "date_relationship",
                            "columns": [col1, col2],
                            "relationship": "exactly_one_year_apart",
                            "description": f"{col2} = {col1} + 1 year"
                        })
                    elif diff_days > 0:
                        rules.append({
                            "type": "date_relationship",
                            "columns": [col1, col2],
                            "relationship": f"exactly_{diff_days}_days_apart",
                            "description": f"{col2} = {col1} + {diff_days} days"
                        })
                elif all(d > 0 for d in diffs):
                    rules.append({
                        "type": "date_relationship",
                        "columns": [col1, col2],
                        "relationship": "after",
                        "description": f"{col2} is always after {col1}"
                    })

    # Check for equal columns (same values)
    for i, col1 in enumerate(headers):
        for col2 in headers[i+1:]:
            vals1 = samples.get(col1, [])
            vals2 = samples.get(col2, [])

            if vals1 and vals2 and len(vals1) == len(vals2):
                if all(v1 == v2 for v1, v2 in zip(vals1, vals2) if v1 is not None and v2 is not None):
                    rules.append({
                        "type": "equal_values",
                        "columns": [col1, col2],
                        "description": f"{col1} always equals {col2}"
                    })

    return rules


def analyze_column(header: str, values: list[Any]) -> dict[str, Any]:
    """Comprehensive analysis of a single column."""
    result = {
        "column_name": header,
        "sample_count": len(values),
        "data_type": detect_data_type(values)
    }

    # Check for empty/null column
    if result["data_type"]["type"] == "empty" or result["data_type"].get("all_null"):
        result["rule_type"] = "empty"
        result["description"] = "Column is always empty/null"
        result["constraints"] = {"always_null": True}
        result["confidence"] = 1.0
        return result

    # Check for sequence
    sequence = detect_sequence(values)
    if sequence and sequence.get("detected"):
        result["rule_type"] = "sequence"
        result["sequence_info"] = sequence
        result["description"] = sequence["description"]
        result["constraints"] = {
            "sequence_type": sequence["sequence_type"],
            "start": sequence.get("start"),
            "step": sequence.get("step", sequence.get("step_days", 1)),
            "prefix": sequence.get("prefix", ""),
            "suffix": sequence.get("suffix", ""),
        }
        result["confidence"] = 0.95
        return result

    # Check for enum
    enum_info = detect_enum(values)
    if enum_info and enum_info.get("detected"):
        result["rule_type"] = "enum"
        result["enum_info"] = enum_info
        result["description"] = enum_info["description"]
        result["constraints"] = {
            "values": enum_info["values"],
            "constant": enum_info.get("constant", False)
        }
        result["confidence"] = 0.95
        return result

    # Check for format
    format_info = detect_format(values)
    if format_info and format_info.get("detected"):
        result["rule_type"] = "format"
        result["format_info"] = format_info
        result["description"] = format_info["description"]
        result["pattern"] = format_info.get("pattern")
        result["constraints"] = {
            "format_name": format_info["format_name"],
            "format_example": format_info.get("example", "")
        }
        result["confidence"] = format_info.get("confidence", 0.9)
        return result

    # Default: free text or numeric
    if result["data_type"]["type"] in ("int", "int_string"):
        non_null = [v for v in values if v is not None]
        nums = []
        for v in non_null:
            try:
                if isinstance(v, int):
                    nums.append(v)
                elif isinstance(v, float):
                    nums.append(int(v))
                elif isinstance(v, str) and v.strip():
                    # Try to parse as int, handle floats too
                    cleaned = v.strip()
                    if '.' in cleaned:
                        nums.append(int(float(cleaned)))
                    else:
                        nums.append(int(cleaned))
            except (ValueError, TypeError):
                continue  # Skip non-numeric values
        if nums:
            result["rule_type"] = "range"
            result["description"] = f"Integer values ({min(nums)} to {max(nums)})"
            result["constraints"] = {"min": min(nums), "max": max(nums), "type": "integer"}
            result["confidence"] = 0.8
        else:
            result["rule_type"] = "text"
            result["description"] = "Free-form text"
            result["confidence"] = 0.5
    elif result["data_type"]["type"] in ("float", "float_string"):
        non_null = [v for v in values if v is not None]
        nums = []
        for v in non_null:
            try:
                if isinstance(v, (int, float)):
                    nums.append(float(v))
                elif isinstance(v, str) and v.strip():
                    nums.append(float(v.strip()))
            except (ValueError, TypeError):
                continue  # Skip non-numeric values
        if nums:
            result["rule_type"] = "range"
            result["description"] = f"Decimal values ({min(nums):.2f} to {max(nums):.2f})"
            result["constraints"] = {"min": min(nums), "max": max(nums), "type": "float"}
            result["confidence"] = 0.8
        else:
            result["rule_type"] = "text"
            result["description"] = "Free-form text"
            result["confidence"] = 0.5
    elif result["data_type"]["type"] == "datetime":
        non_null = [v for v in values if v is not None and isinstance(v, datetime)]
        if non_null:
            # Detect the date format from original data
            date_fmt = result["data_type"].get("date_format", "date")
            result["rule_type"] = "date"
            result["description"] = f"Date values ({min(non_null).strftime('%Y-%m-%d')} to {max(non_null).strftime('%Y-%m-%d')})"
            result["constraints"] = {
                "format": date_fmt,  # Use detected format (date vs datetime)
                "output_format": "YYYY-MM-DD" if date_fmt == "date" else "datetime",
                "range_start": min(non_null).isoformat(),
                "range_end": max(non_null).isoformat()
            }
            result["confidence"] = 0.9
        else:
            result["rule_type"] = "text"
            result["description"] = "Free-form text"
            result["confidence"] = 0.5
    else:
        result["rule_type"] = "text"
        result["description"] = "Free-form text"
        result["confidence"] = 0.5

    return result


def _format_value_for_display(value: Any) -> str:
    """Format a value for display, handling datetime objects properly."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        # Check if time component is meaningful (not midnight)
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%Y-%m-%d")  # Date only
        else:
            return value.strftime("%Y-%m-%d %H:%M:%S")  # Full datetime
    return str(value)


def analyze_sheet_patterns_deterministic(
    sheet_name: str,
    headers: list[str],
    samples: dict[str, list[Any]]
) -> dict[str, Any]:
    """
    Main entry point: Analyze all columns in a sheet using deterministic methods.

    Returns a structure compatible with the existing rule format.
    """
    rules = []

    for header in headers:
        values = samples.get(header, [])
        analysis = analyze_column(header, values)

        # Convert to rule format
        rule = {
            "column_name": analysis["column_name"],
            "rule_type": analysis.get("rule_type", "text"),
            "description": analysis.get("description", ""),
            "pattern": analysis.get("pattern"),
            "constraints": analysis.get("constraints", {}),
            "examples": [_format_value_for_display(v) for v in values[:3] if v is not None],
            "confidence": analysis.get("confidence", 0.5),
            "data_type": analysis.get("data_type", {}).get("type", "string"),
            "detection_method": "deterministic"
        }
        rules.append(rule)

    # Detect cross-column rules
    cross_rules = detect_cross_column_rules(headers, samples)

    return {
        "rules": rules,
        "cross_column_rules": [r["description"] for r in cross_rules],
        "cross_column_rules_detailed": cross_rules
    }
