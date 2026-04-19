from openai import OpenAI
import json
import os
import time
import random
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any


# ---------------------------------------------------------------------------
# Response parsing helpers
# ---------------------------------------------------------------------------

def _clean_response_text(text: str) -> str:
    """Strip common fenced-code wrappers from model output."""
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
        for value in parsed.values():
            if isinstance(value, list) and value and isinstance(value[0], dict):
                return value
        return [parsed]
    raise ValueError(f"Model output is not a JSON array or wrapped array. Got: {type(parsed)}")


# ---------------------------------------------------------------------------
# Date utilities
# ---------------------------------------------------------------------------

def _format_mmddyyyy(value: date) -> str:
    return value.strftime("%m/%d/%Y")


def _add_one_year(value: date) -> date:
    try:
        return value.replace(year=value.year + 1)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + 1)


def _find_header_key(row: dict[str, Any], candidates: list[str]) -> str | None:
    """Return the first row key whose lowercase form contains any candidate substring."""
    for key in row.keys():
        key_lower = key.lower()
        if any(c in key_lower for c in candidates):
            return key
    return None


# ---------------------------------------------------------------------------
# Special-instruction parser
# ---------------------------------------------------------------------------

@dataclass
class ParsedInstructions:
    """
    All structured config extracted from a free-text special instruction string.

    Every attribute has a sensible default so callers always get a valid object
    even when the user provides no (or partial) instructions.
    """

    # ---- date range --------------------------------------------------------
    date_range_start: int | None = None
    date_range_end: int | None = None
    distribute_all_years: bool = False
    cap_expiration_to_range: bool = True

    # ---- data set numbering ------------------------------------------------
    ds_prefix: str = "DS_"
    ds_start: int = 1

    # ---- row distribution / state filtering --------------------------------
    allowed_states: list[str] = field(default_factory=list)

    # ---- LOB flags ---------------------------------------------------------
    active_lobs: list[str] = field(default_factory=list)

    # ---- address uniqueness ------------------------------------------------
    unique_addresses: bool = True

    # ---- submitted-date behaviour -----------------------------------------
    submitted_date_mode: str = "current"
    submitted_date_static: date | None = None

    # ---- percentage format -------------------------------------------------
    pct_format: str = "symbol"

    # ---- raw pass-through --------------------------------------------------
    raw: str = ""


def parse_special_instructions(text: str, row_count: int = 0) -> ParsedInstructions:
    """Parse a free-text special instruction string into a structured ParsedInstructions object.

    `row_count` is used to auto-enable year distribution when a multi-year range
    is detected and the user has asked for enough rows to cover it.
    """
    pi = ParsedInstructions(raw=text)
    if not text:
        return pi

    t = text.lower()

    # ---- date range --------------------------------------------------------
    date_range_match = re.search(
        r"(20\d{2})\s*(?:-|to|–|—|through|thru|till|until)\s*(20\d{2})", t
    )
    if date_range_match:
        s, e = int(date_range_match.group(1)), int(date_range_match.group(2))
        pi.date_range_start = min(s, e)
        pi.date_range_end = max(s, e)

    if pi.date_range_start is None:
        single_year = re.search(r"\b(20\d{2})\b", t)
        if single_year and any(kw in t for kw in ["effective", "expir", "date"]):
            yr = int(single_year.group(1))
            pi.date_range_start = yr
            pi.date_range_end = yr

    # Keyword detection: matches "all years", "every year", "each year",
    # "include all 7 years", "all 5 years", "inclusive", etc.
    all_years_patterns = [
        r"\ball\s+years\b",
        r"\ball\s+\d+\s+years?\b",       # "all 7 years"
        r"\bevery\s+year\b",
        r"\beach\s+year\b",
        r"\byears?\s+from\b",
        r"\binclude\s+all\s+years?\b",
        r"\binclude\s+all\s+\d+\s+years?\b",  # "include all 7 years"
        r"\bto\s+be\s+included\b",
        r"\bone\s+per\s+year\b",
        r"\bspread\s+(?:across|over)\s+years?\b",
        r"\bdistribute\s+(?:across|over)\s+years?\b",
        r"\binclusive\)?",                # "(inclusive)"
    ]
    keyword_hit = any(re.search(p, t) for p in all_years_patterns)

    # Auto-enable distribution whenever a multi-year range is detected AND
    # the user either used a distribution keyword OR asked for at least as
    # many rows as years in the span. This makes "2020-2026" + row_count>=7
    # behave intuitively even without explicit phrasing.
    if pi.date_range_start is not None and pi.date_range_end is not None:
        span = pi.date_range_end - pi.date_range_start + 1
        if span >= 2 and (keyword_hit or row_count >= span):
            pi.distribute_all_years = True
    else:
        pi.distribute_all_years = keyword_hit

    if any(phrase in t for phrase in [
        "expiration can exceed", "allow expiration outside", "don't cap expiration",
        "no cap on expiration",
    ]):
        pi.cap_expiration_to_range = False

    # ---- DS numbering ------------------------------------------------------
    ds_prefix_match = re.search(r"ds\s+prefix[:\s]+([A-Za-z0-9_]+)", t)
    if ds_prefix_match:
        pi.ds_prefix = ds_prefix_match.group(1).upper() + "_"

    ds_start_match = re.search(r"ds\s+start(?:ing)?\s+(?:from\s+)?(\d+)", t)
    if ds_start_match:
        pi.ds_start = int(ds_start_match.group(1))

    # ---- state filtering ---------------------------------------------------
    state_match = re.search(
        r"(?:only|for|states?)[:\s]+([A-Z]{2}(?:[,\s/]+[A-Z]{2})*)",
        text,
    )
    if state_match:
        raw_states = re.findall(r"[A-Z]{2}", state_match.group(1))
        us_states = {
            "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
            "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
            "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
            "TX","UT","VT","VA","WA","WV","WI","WY","DC",
        }
        pi.allowed_states = [s for s in raw_states if s in us_states]

    # ---- submitted date mode -----------------------------------------------
    if "submitted" in t:
        if any(p in t for p in ["submitted = effective", "submitted match effective",
                                 "submitted same as effective"]):
            pi.submitted_date_mode = "match"
        static_match = re.search(r"submitted\s+(?:date\s+)?(?:is\s+|=\s*)(\d{4}-\d{2}-\d{2})", t)
        if static_match:
            try:
                pi.submitted_date_static = datetime.strptime(
                    static_match.group(1), "%Y-%m-%d"
                ).date()
                pi.submitted_date_mode = "static"
            except ValueError:
                pass

    # ---- percentage format -------------------------------------------------
    if any(p in t for p in ["percent as decimal", "pct as decimal",
                              "percentage as decimal", "use decimal for percent"]):
        pi.pct_format = "decimal"

    # ---- address uniqueness ------------------------------------------------
    if any(p in t for p in ["reuse address", "allow duplicate address",
                              "same address allowed"]):
        pi.unique_addresses = False

    return pi


# ---------------------------------------------------------------------------
# Date-enforcement post-processor
# ---------------------------------------------------------------------------

def _build_date_policy_summary(pi: ParsedInstructions) -> str:
    if pi.date_range_start is not None:
        s, e = pi.date_range_start, pi.date_range_end
        span = e - s + 1
        if pi.distribute_all_years:
            years_list = ", ".join(str(y) for y in range(s, e + 1))
            return (
                f"Effective dates MUST span ALL {span} years from {s} through {e} "
                f"inclusive ({years_list}). Every year in this list must appear at "
                f"least once. Distribute evenly across rows. Expiration Date is "
                f"exactly 1 year after Effective Date"
                f"{', capped at ' + str(e) + '/12/31 if +1y exceeds the range.' if pi.cap_expiration_to_range else '.'}"
            )
        return (
            f"Effective dates must fall within {s}–{e} (inclusive). "
            f"Expiration Date is exactly 1 year after Effective"
            f"{', capped at ' + str(e) + '/12/31 if the +1 year exceeds the range.' if pi.cap_expiration_to_range else '.'}"
        )
    return (
        "Use today's date for Effective Date; "
        "Expiration Date is exactly 1 year later."
    )


def _enforce_effective_expiration_date_range(
    rows: list[dict[str, Any]],
    pi: ParsedInstructions,
) -> list[dict[str, Any]]:
    """Post-process rows to enforce date policy deterministically."""
    effective_candidates = ["effective date", "effetive date", "effetive", "effective"]
    expiration_candidates = ["expiration date", "expiry date", "expiry", "expiration", "exp date"]
    submitted_candidates = ["submitted"]
    rates_rules_candidates = ["rates & rules", "rates and rules", "rate and rule", "rate & rule"]

    today = date.today()
    has_range = pi.date_range_start is not None

    if has_range:
        start_yr, end_yr = pi.date_range_start, pi.date_range_end
        year_pool = list(range(start_yr, end_yr + 1)) if end_yr > start_yr else [start_yr]
    else:
        year_pool = []

    for idx, row in enumerate(rows):
        eff_key = _find_header_key(row, effective_candidates)
        exp_key = _find_header_key(row, expiration_candidates)
        sub_key = _find_header_key(row, submitted_candidates)
        rr_key = _find_header_key(row, rates_rules_candidates)

        if eff_key or exp_key:
            if has_range:
                if pi.distribute_all_years and year_pool:
                    eff_year = year_pool[idx % len(year_pool)]
                else:
                    eff_year = random.randint(start_yr, end_yr)
                eff_month = random.randint(1, 12)
                eff_day = random.randint(1, 28)
                eff_dt = date(eff_year, eff_month, eff_day)
            else:
                eff_dt = today

            exp_dt = _add_one_year(eff_dt)

            # Expiration is always effective + 1 year regardless of range cap.
            # The end_yr constraint applies only to effective dates, not expirations.
            # (capping expiration to 12/31 of end_yr produces sub-1-year policies)

            if eff_key:
                row[eff_key] = _format_mmddyyyy(eff_dt)
            if exp_key:
                row[exp_key] = _format_mmddyyyy(exp_dt)

            # Rates & Rules as of = same as Effective Date
            if rr_key:
                row[rr_key] = _format_mmddyyyy(eff_dt)

        # ---- submitted date ------------------------------------------------
        if sub_key:
            if pi.submitted_date_mode == "current":
                sub_dt = today
            elif pi.submitted_date_mode == "match" and eff_key:
                sub_dt = eff_dt if (eff_key or exp_key) else today
            elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
                sub_dt = pi.submitted_date_static
            else:
                sub_dt = today
            row[sub_key] = _format_mmddyyyy(sub_dt)

    return rows


# ---------------------------------------------------------------------------
# Property sheet business rule enforcer
# ---------------------------------------------------------------------------

_VALID_STORIES = ["1", "2", "3"]
_VALID_COINSURANCE = ["80", "90", "100"]
_VALID_INFLATION_GUARD = ["2%", "3%", "4%", "6%", "8%", "10%", "N/A"]
_VALID_CAUSE_OF_LOSS = ["Basic", "Special", "Broad"]

# ---------------------------------------------------------------------------
# Exact unique-header-name → forced value for the Property sheet.
# These are derived from KMG_IMS_Data-Set v1.0.xlsx column names.
# Unique headers (appear only once) keep their plain name; duplicate headers
# are disambiguated by the "(Col N)" suffix added in main.py.
# ---------------------------------------------------------------------------
_PROPERTY_EXACT: dict[str, Any] = {
    # Building Ordinance section (AF, AG, AH)
    "Building Ordinance":                "No",
    "Building Ordinance (B)":            0,
    "Building Ordinance (C)":            0,
    # Personal property detail section (AI–AO all blank)
    "Personal Property Sub-Description": "",
    "Property of Insured":               "",
    "Property of Others":                "",
    "Improvements":                      "",   # col AM — different from "Tenant Improvements"
    "Stock":                             "",   # col AN
    # Alarm / security (AZ, BA)
    "Watchman Service":                  "None",
    "Alarm":                             "None",
    # Business Income section fixed fields
    "Civil Authority":                   "N/A",
    "Income with extra Expense":         "No",
    # Dependent Properties section (BP–BR)
    "Dependent Properties Limit":        "0",
    "Dependent Properties Form":         "N/A",
    "Limit Percentage":                  "100%",
    # Business Income deductible/waiting period fields (BK)
    "Days in Deductible":               "N/A",
    # Extra Expense / Utility section (BS–BW)
    "Extra Expense Limit":               "$0",
    "Liability":                         "N/A",
    "Limits on Loss Payment":            "N/A",
    "Utility Service Coverage Provided": "N/A",
    "Public or Other":                   "N/A",
    # Sub-limits (BX–BZ)
    "Building Sub-Limit":                "$0",
    "Personal Property Sub-Limit":       "$0",
    "Earthquake Sub-Limit":              "$0",
    # Utility / supply lines (CA–CE)
    "Water Supply":                      "No",
    "Communication Supply":              "No",
    "Power Supply":                      "No",
    "Power Lines":                       "No",
    "Communication Lines":               "No",
    # Pollution / debris / fire legal etc. (CF–CN)
    "Pollutant Cleanup and Removal":     "$0",
    "Pollutant Cleanup Deductible":      "N/A",
    "Debris Removal Over 10K":           "$0",
    "Fire Legal Building":               "$0",
    "Fire Legal Contents":               "$0",
    "Spoilage Coverage Limit":           "$0",
    "Peak Season Additional Limit 1":    "$0",
    "Peak Season Additional Limit 2":    "$0",
    "Vacancy Permit Exposure":           "$0",
    # Earthquake / masonry / sublimit tail (CQ–CX)
    "Earthquake Coverage":               "N/A",
    "Masonry Veneer Limit":              "N/A",
    "Roof Tank on Building":             "No",
    "Sublimit Amount":                   "$0",
    "Sublimit Percentage":               "N/A",
    "Sprinkler Leakage":                 "N/A",
    # Final descriptive fields (CY, CZ) — always blank
    "Building Construction":             "",
    "Building Description":              "",
}

# Duplicate-header rules: match by "(Col N)" position suffix.
# Format: col_number (1-indexed, matching main.py idx+1) → forced value.
_PROPERTY_BY_COL: dict[int, Any] = {
    # Number of Stories (CT) = col 98 → always 0 (personal property section)
    98: 0,
    # Deductible (CR) = col 96 → always "5%"  (Earthquake deductible)
    96: "5%",
}

# Columns whose value must be COPIED from an earlier sibling (AQ-AW mirrors W-AC).
# Format: source_col_number → dest_col_number (both 1-indexed)
_PROPERTY_COPY_PAIRS: list[tuple[int, int]] = [
    (23, 43),  # Cause of Loss W → AQ
    (24, 44),  # Coinsurance X → AR
    (25, 45),  # Deductible Y → AS
    (26, 46),  # Wind Hail Deductible Z → AT
    (27, 47),  # Theft Deductible AA → AU
    (28, 48),  # Inflation Guard AB → AV
    (29, 49),  # Valuation AC → AW
]


def _col_num(key: str) -> int | None:
    """Extract the 1-based column number from a '(Col N)' suffix, or None."""
    m = re.search(r"\(col\s+(\d+)\)", key.lower())
    return int(m.group(1)) if m else None


def _enforce_property_sheet_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deterministically fix the most common LLM mistakes on the Property sheet.

    Uses exact unique-header names from KMG_IMS_Data-Set v1.0.xlsx so that
    fixes are precise and never clobber unrelated columns.
    """
    for row in rows:
        keys = list(row.keys())

        # ---- Apply exact-name fixed values ----------------------------------
        for k in keys:
            # Strip any "(Col N)" suffix for the name-only lookup
            bare = re.sub(r"\s*\(col\s+\d+\)", "", k, flags=re.IGNORECASE).strip()
            if k in _PROPERTY_EXACT:
                row[k] = _PROPERTY_EXACT[k]
            elif bare in _PROPERTY_EXACT and bare != k:
                # Only apply if this key is NOT a duplicate that needs
                # different treatment (e.g. "Other" vs "Public or Other")
                if bare not in ("Other",):   # "Other" col AO is handled below
                    row[k] = _PROPERTY_EXACT[bare]

        # ---- "Other" (col AO = 41): blank — exact key only ------------------
        for k in keys:
            if k.strip() == "Other":
                row[k] = ""

        # ---- Apply col-number-based fixed values ----------------------------
        for k in keys:
            col_n = _col_num(k)
            if col_n in _PROPERTY_BY_COL:
                row[k] = _PROPERTY_BY_COL[col_n]

        # ---- Number of Stories (col L = 12): must be 1, 2, or 3 ------------
        # Col CT (98) is handled by _PROPERTY_BY_COL above (forced to 0).
        for k in keys:
            if "number of stories" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n >= 90:
                continue   # That's the CT copy (handled above)
            try:
                n = int(str(row[k]).strip())
                if str(n) not in _VALID_STORIES:
                    row[k] = random.choice(_VALID_STORIES)
            except (ValueError, TypeError):
                row[k] = random.choice(_VALID_STORIES)

        # ---- Coinsurance (col X = 24): must be 80, 90, or 100 ---------------
        for k in keys:
            if "coinsurance" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue   # AR (col 44) is a copy — handled by copy-pairs below
            val = str(row[k]).strip()
            if val not in _VALID_COINSURANCE:
                row[k] = random.choice(_VALID_COINSURANCE)

        # ---- Inflation Guard (col AB = 28): valid percentage or N/A ---------
        for k in keys:
            if "inflation guard" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue   # AV (col 48) is a copy
            val = str(row[k]).strip()
            if val not in _VALID_INFLATION_GUARD:
                row[k] = random.choice(_VALID_INFLATION_GUARD)

        # ---- Cause of Loss (col W = 23): Basic, Special, or Broad ----------
        for k in keys:
            if "cause of loss" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue   # AQ (col 43) is a copy
            val = str(row[k]).strip()
            if val not in _VALID_CAUSE_OF_LOSS:
                row[k] = random.choice(_VALID_CAUSE_OF_LOSS)

        # ---- Tenant Improvements (col U): always $0 — never blank ----------
        for k in keys:
            if "tenant improvement" in k.lower():
                val = str(row.get(k, "")).strip()
                if val in ("", "nan", "None", "none"):
                    row[k] = "$0"

        # ---- Total Exposure (col V = 22): must equal Building Exposure (T = 20)
        bldg_exp_key = None
        total_v_key = None
        for k in keys:
            kl = k.lower()
            if "building exposure" in kl:
                bldg_exp_key = k
            elif "total exposure" in kl:
                col_n = _col_num(k)
                # Col V = 22 (first Total Exposure), Col AP = 42 (personal prop total)
                if col_n is None or col_n <= 30:
                    total_v_key = k
        if bldg_exp_key and total_v_key:
            row[total_v_key] = row.get(bldg_exp_key, row.get(total_v_key, ""))

        # ---- AQ–AW: copy values from matching W–AC columns ------------------
        # Build col_number → key lookup once per row
        col_to_key: dict[int, str] = {}
        for k in keys:
            col_n = _col_num(k)
            if col_n is not None:
                col_to_key[col_n] = k

        for src_col, dst_col in _PROPERTY_COPY_PAIRS:
            src_k = col_to_key.get(src_col)
            dst_k = col_to_key.get(dst_col)
            if src_k and dst_k:
                row[dst_k] = row.get(src_k, row.get(dst_k, ""))

        # ---- Agreed Value (AE=col31, AY=col51): must be Yes/No, not dollar amounts
        for k in keys:
            if "agreed value" not in k.lower():
                continue
            col_n = _col_num(k)
            # col 67 = BO (Business Income Agreed Value) → always No
            if col_n is not None and col_n >= 65:
                row[k] = "No"
                continue
            val = str(row.get(k, "")).strip()
            # If it's a dollar amount or numeric, convert to Yes
            if val and val not in ("Yes", "No", ""):
                row[k] = "Yes"
            elif not val:
                row[k] = "No"

        # ---- Civil Authority (BL=col64): always N/A
        for k in keys:
            if "civil authority" in k.lower():
                row[k] = "N/A"

        # ---- Ordinance or Law - Increase Period of Restoration (BJ=col62): Yes/No only
        for k in keys:
            kl = k.lower()
            if "ordinance or law" in kl and "increase period" in kl:
                val = str(row.get(k, "")).strip()
                if val not in ("Yes", "No"):
                    row[k] = random.choice(["Yes", "No"])

        # ---- Days in Deductible (BK=col63): always N/A
        for k in keys:
            if "days in deductible" in k.lower():
                row[k] = "N/A"

        # ---- Ordinary Payroll (BN=col66): IMS-valid waiting-period values only
        _VALID_ORD_PAYROLL = ["No Limitation", "0 days", "90 days", "180 days"]
        for k in keys:
            if "ordinary payroll" in k.lower():
                val = str(row.get(k, "")).strip()
                if val not in _VALID_ORD_PAYROLL:
                    row[k] = random.choice(_VALID_ORD_PAYROLL)

        # ---- Business Income Type (BG=col59): valid IMS values only
        _VALID_BI_TYPE = ["Mercantile", "Manufacturing", "Rental Properties"]
        for k in keys:
            if "business income type" in k.lower():
                val = str(row.get(k, "")).strip()
                if val not in _VALID_BI_TYPE:
                    row[k] = random.choice(_VALID_BI_TYPE)

        # ---- Time Period (BM=col65): must be exact IMS waiting-period values
        _TP_VALID = {"72-Hour Waiting", "48-Hour Waiting", "24-Hour Waiting",
                     "12-Hour Waiting", "No Waiting"}
        _TP_MAP = {
            "90 days": "72-Hour Waiting", "180 days": "72-Hour Waiting",
            "365 days": "No Waiting", "30 days": "72-Hour Waiting",
            "60 days": "72-Hour Waiting", "0 days": "No Waiting",
            "0": "No Waiting", "no limitation": "No Waiting",
            "none": "No Waiting",
        }
        for k in keys:
            if "time period" in k.lower():
                val = str(row.get(k, "")).strip()
                if val in _TP_VALID:
                    pass  # already valid
                elif val.lower() in _TP_MAP:
                    row[k] = _TP_MAP[val.lower()]
                elif val:
                    row[k] = "72-Hour Waiting"

        # ---- Personal Property Total Exposure (AP=col42) must equal Contents (AJ=col36)
        pp_contents_key = None
        pp_total_key = None
        for k in keys:
            kl = k.lower()
            col_n = _col_num(k)
            if kl.strip() == "contents" or (kl == "contents" and (col_n is None or 34 <= (col_n or 0) <= 38)):
                pp_contents_key = k
            elif "total exposure" in kl and col_n is not None and 40 <= col_n <= 44:
                pp_total_key = k
        if pp_contents_key and pp_total_key:
            row[pp_total_key] = row.get(pp_contents_key, row.get(pp_total_key, ""))

    return rows


# ---------------------------------------------------------------------------
# Crime sheet conditional field enforcer
# ---------------------------------------------------------------------------

# Maps coverage parent name keywords → list of dependent field keywords that
# must be blanked when the parent is "No".
_CRIME_PARENT_DEPS: list[tuple[list[str], list[str]]] = [
    (
        ["employee theft and forgery", "employee theft"],
        ["ratable employees", "additional premises", "faithful performance",
         "limit of insurance", "deductible amount", "deductible limit",
         "include expenses", "credit card", "clients property"],
    ),
    (
        ["fraud impersonation"],
        ["ratable employees", "additional premises", "employees for",
         "verification option", "in excess of limit", "limit of insurance",
         "deductible amount"],
    ),
    (
        ["robbery"],
        ["type for", "limit of insurance", "deductible amount"],
    ),
    (
        ["outside the premises"],
        ["limit of insurance", "deductible amount"],
    ),
    (
        ["destruction of electronic data", "electronic data"],
        ["ratable employees", "additional premises", "limit of insurance", "deductible amount"],
    ),
    (
        ["guests property"],
        ["safe deposit", "number of rooms", "room or apartment", "increased limit per guest",
         "food or liquor", "laundry", "articles for sale"],
    ),
    (
        ["theft of money", "money or securities"],
        ["limit of insurance", "deductible", "increased limit", "number of premises",
         "date beginning", "date ending"],
    ),
    (
        ["money orders and counterfeit", "money orders"],
        # Match the disambiguated column headers: "Limit (Col 56)" / "Deductible (Col 57)"
        # or bare "Limit"/"Deductible" if they happen to be unique in the sheet.
        # "(col 56)" and "(col 57)" pin exactly to Money Orders' two child columns (BD, BE).
        ["(col 56)", "(col 57)"],
    ),
    (
        ["computer and fund transfer"],
        ["ratable employees", "additional premises", "limit of insurance", "deductible for"],
    ),
    (
        ["identity fraud expense"],
        ["ratable employees", "additional premises", "limit of insurance", "deductible"],
    ),
]


def _enforce_crime_conditional_fields(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Enforce Crime coverage conditional field rules.

    Pass order matters:
      1. YES fill  — when parent=Yes and a child is blank, populate a sensible default.
      2. NO blank  — when parent=No, blank ALL dependent children (authoritative override).

    This ensures a parent=No always produces blank children regardless of LLM output,
    and YES fill only ever fills genuinely blank children (never overwrites real values).
    """

    def _crime_child_matches(kl: str, child_keywords: list[str]) -> bool:
        """Match a column name against child_keywords.

        For the bare keyword 'deductible' we use an exact/anchored pattern so that
        columns like 'Safe Deposit Deductible' are NOT matched by Employee Theft's
        deductible child rule.  All other keywords use plain substring matching.
        """
        for ck in child_keywords:
            if ck == "deductible":
                # Match "deductible" or "deductible (col N)" but NOT "* deductible" or "deductible *"
                if re.match(r"^deductible(\s+\(col\s+\d+\))?$", kl):
                    return True
            else:
                if ck in kl:
                    return True
        return False

    # -----------------------------------------------------------------------
    # YES fill spec: parent_keywords → [(child_keywords, fill_lambda), ...]
    # -----------------------------------------------------------------------
    _CRIME_YES_FILL: list[tuple[list[str], list[tuple[list[str], Any]]]] = [
        (
            ["employee theft and forgery", "employee theft"],
            [
                (["ratable employees"], lambda: str(random.randint(10, 75))),
                (["additional premises"], lambda: str(random.randint(1, 3))),
                (["limit of insurance"], lambda: random.choice(["$50,000", "$75,000", "$100,000", "$150,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000", "2500"])),
            ],
        ),
        (
            ["fraud impersonation"],
            [
                (["ratable employees"], lambda: str(random.randint(10, 75))),
                (["additional premises"], lambda: str(random.randint(1, 3))),
                (["limit of insurance"], lambda: random.choice(["$50,000", "$75,000", "$100,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000", "2500"])),
            ],
        ),
        (
            ["robbery"],
            [
                (["limit of insurance"], lambda: random.choice(["$25,000", "$50,000", "$75,000", "$100,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000"])),
            ],
        ),
        (
            ["outside the premises"],
            [
                (["limit of insurance"], lambda: random.choice(["$15,000", "$25,000", "$50,000"])),
                (["deductible"], lambda: random.choice(["250", "500"])),
            ],
        ),
        (
            ["destruction of electronic data", "electronic data"],
            [
                (["ratable employees"], lambda: str(random.randint(10, 60))),
                (["additional premises"], lambda: str(random.randint(1, 3))),
                (["limit of insurance"], lambda: random.choice(["$10,000", "$20,000", "$30,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000"])),
            ],
        ),
        (
            ["theft of money", "money or securities"],
            [
                (["limit of insurance"], lambda: random.choice(["$10,000", "$25,000", "$50,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000"])),
                (["increased limit"], lambda: random.choice(["$5,000", "$10,000", "$15,000"])),
            ],
        ),
        (
            ["money orders and counterfeit", "money orders"],
            [
                # Use column-number-anchored keywords to avoid matching other Limit/Deductible cols
                (["(col 56)"], lambda: random.choice(["$2,500", "$5,000", "$7,500"])),
                (["(col 57)"], lambda: random.choice(["100", "250", "500"])),
            ],
        ),
        (
            ["computer and fund transfer"],
            [
                (["ratable employees"], lambda: str(random.randint(10, 60))),
                (["additional premises"], lambda: str(random.randint(1, 3))),
                (["limit of insurance"], lambda: random.choice(["$25,000", "$50,000", "$75,000"])),
                (["deductible"], lambda: random.choice(["500", "1000", "2500"])),
            ],
        ),
        (
            ["identity fraud expense"],
            [
                (["ratable employees"], lambda: str(random.randint(10, 50))),
                (["additional premises"], lambda: str(random.randint(1, 2))),
                (["limit of insurance"], lambda: random.choice(["$10,000", "$15,000", "$25,000"])),
                (["deductible"], lambda: random.choice(["250", "500", "1000"])),
            ],
        ),
    ]

    # ---- Pass 1: YES fill (must run BEFORE NO blank) -----------------------
    for row in rows:
        keys = list(row.keys())
        for parent_keywords, child_rules in _CRIME_YES_FILL:
            parent_key = None
            for k in keys:
                kl = k.lower()
                if any(pk in kl for pk in parent_keywords):
                    parent_key = k
                    break
            if parent_key is None:
                continue
            val = str(row.get(parent_key, "")).strip().lower()
            if val != "yes":
                continue
            for child_keywords, fill_fn in child_rules:
                for k in keys:
                    if k == parent_key:
                        continue
                    kl = k.lower()
                    if _crime_child_matches(kl, child_keywords):
                        existing = row.get(k)
                        if existing is None or str(existing).strip() in ("", "None", "none"):
                            row[k] = fill_fn()

    # ---- Pass 2: NO blank (authoritative — runs AFTER YES fill) -----------
    for row in rows:
        keys = list(row.keys())
        for parent_keywords, dep_keywords in _CRIME_PARENT_DEPS:
            parent_key = None
            for k in keys:
                kl = k.lower()
                if any(pk in kl for pk in parent_keywords):
                    parent_key = k
                    break
            if parent_key is None:
                continue
            val = str(row.get(parent_key, "")).strip().lower()
            if val != "no":
                continue
            # Parent is No → blank every matching dependent column
            for k in keys:
                if k == parent_key:
                    continue
                kl = k.lower()
                if any(dk in kl for dk in dep_keywords):
                    row[k] = ""

    # Validate Coverage is Written — must be Primary, Excess, or Concurrent
    for row in rows:
        for k in list(row.keys()):
            if "coverage is written" in k.lower():
                val = str(row.get(k, "") or "").strip()
                if val not in ("Primary", "Excess", "Concurrent"):
                    row[k] = "Primary"

    return rows


# ---------------------------------------------------------------------------
# Per-sheet business rule dispatcher
# ---------------------------------------------------------------------------

_VALID_POLICY_TYPES = [
    "N/A", "Apartment", "Contractor", "Mercantile",
    "Motel/Hotel", "Office", "Institutional", "Industrial/Processing"
]


def _enforce_policy_sheet_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for row in rows:
        for k in list(row.keys()):
            # Deductible Property Damage Only → always blank
            if "deductible property damage" in k.lower():
                row[k] = ""
            # Any Vehicles used in TNC / On-Demand → always blank (field must be empty per spec)
            if "transportation network" in k.lower() or "on-demand capacity" in k.lower():
                row[k] = ""
            # GarageKeepers → always No
            if k.strip().lower() == "garagekeepers":
                row[k] = "No"
            # Type field → normalize spacing and validate
            if k.strip().lower() == "type":
                val = str(row.get(k, "") or "").strip()
                # Normalize spacing around slash
                val = re.sub(r'\s*/\s*', '/', val)
                if val not in _VALID_POLICY_TYPES:
                    row[k] = "Office"
                else:
                    row[k] = val
    return rows


_VALID_IMS_PRODUCERS = [
    "Acme Producer- NewYork - Jericho, NY",
    "Ash Group Prospect -Ewing, NJ",
    "Cigna Group -Washington, DC",
    "CRC Insurance Services Excess MPL -Jericho, NY",
    "Dextorville -Brooklyn, NY",
    "Evergreen Shield Insurance -New York, NY",
    "Iron Gate Insurance Group -Brooklyn, NY",
    "James Smith - NewYork - Jericho, NY",
    "Mutual's -Coral Gables, FL",
    "Pinnacle Assurance -Washington, DC",
    "PrimeCover Associates -Beverly Hills, CA",
    "SecureEdge Covers - Jericho, NY",
    "SSI Producer - Brooklyn, NY",
    "Statewide's -New York, NY",
    "Summit Ridge Risk Solutions -Las Vegas, NV",
    "TCI Boston -Boston, MA",
    "Test Producer_01_17 -Brooklyn, NY",
    "Test Producer_2 -Shreveport, LA",
    "The Armour's -Las Vegas, NV",
    "Ul Agency -Brooklyn, NY",
    "Unity Risk Management - Shreveport, LA",
]


def _enforce_ims_screen_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deterministically fix IMS Screen fields the LLM commonly gets wrong.

    Current fixes:
    - ZIP (col F): zero-pad to exactly 5 digits (e.g. 8618 → 08618).
    - Address (col E): strip city/state/ZIP if present — keep street only.
    - Producer by Location (col G): must be one of the exact valid producer values.
    """
    for row in rows:
        for k in list(row.keys()):
            kl = k.lower()
            if "zip" in kl:
                val = str(row[k]).strip()
                if val.isdigit() and 1 <= len(val) <= 4:
                    row[k] = val.zfill(5)
            elif "address" in kl:
                val = str(row.get(k, "") or "").strip()
                if val and "," in val:
                    # Strip everything after the first comma (city/state/zip)
                    row[k] = val.split(",")[0].strip()
                # Also strip trailing 5-digit ZIP
                elif val:
                    row[k] = re.sub(r'\s+\d{5}(-\d{4})?\s*$', '', val).strip()
            elif "producer" in kl and "location" in kl:
                val = str(row.get(k, "") or "").strip()
                if val not in _VALID_IMS_PRODUCERS:
                    row[k] = random.choice(_VALID_IMS_PRODUCERS)
    return rows


def _enforce_sheet_business_rules(
    rows: list[dict[str, Any]],
    sheet_name: str,
) -> list[dict[str, Any]]:
    """Route rows through the appropriate sheet-specific enforcer."""
    sn = sheet_name.lower().strip()
    if "ims" in sn and "screen" in sn:
        rows = _enforce_ims_screen_values(rows)
    if "policy" in sn:
        rows = _enforce_policy_sheet_values(rows)
    if "property" in sn and "sub" not in sn and "inland" not in sn:
        rows = _enforce_property_sheet_values(rows)
    if "crime" in sn:
        rows = _enforce_crime_conditional_fields(rows)
    return rows


# ---------------------------------------------------------------------------
# Cross-sheet post-processing (structural, not rule-hardcoded)
# ---------------------------------------------------------------------------

def enforce_cross_sheet_consistency(
    rows: list[dict[str, Any]],
    sheet_name: str,
    previous_sheets_data: dict[str, list[dict[str, Any]]] | None,
) -> list[dict[str, Any]]:
    """
    Enforce structural cross-sheet relationships that the LLM may miss.
    These are structural rules (data must match between sheets), not business
    rules (those come from the special instructions / prompt).
    """
    if not previous_sheets_data or not rows:
        return rows

    sn = sheet_name.lower().strip()

    # Find reference sheets (tolerant of NETRATE prefixes and misspellings)
    ims_data = None
    policy_data = None
    property_data = None
    location_data = None

    for k, v in previous_sheets_data.items():
        kl = k.lower().strip()
        if "ims" in kl and "screen" in kl:
            ims_data = v
        elif "policy" in kl:
            policy_data = v
        elif "property" in kl and "sub" not in kl:
            property_data = v
        elif "location" in kl:
            location_data = v

    for idx, row in enumerate(rows):
        # ---- DS, dates, submitted must be identical across all sheets ------
        if ims_data and idx < len(ims_data):
            ref = ims_data[idx]
            for cand in [
                ["effective", "effetive"],
                ["expiration", "expiry"],
                ["submitted"],
                ["data set"],
            ]:
                src_key = _find_header_key(ref, cand)
                dst_key = _find_header_key(row, cand)
                if src_key and dst_key:
                    row[dst_key] = ref[src_key]

        # ---- Rates & Rules as of = Effective Date (must re-sync AFTER the
        #      cross-sheet copy above, which may have updated the effective date
        #      to a different value than what the date-enforcer originally set).
        eff_key = _find_header_key(row, ["effective", "effetive"])
        rr_key  = _find_header_key(row, ["rates & rules", "rates and rules",
                                         "rate and rule", "rate & rule"])
        if eff_key and rr_key:
            row[rr_key] = row[eff_key]

        # ---- Location sheet: Address column C must mirror IMS column E -----
        # IMS stores street-only address ("1450 Northern Blvd"); Location
        # needs the full address with city/state/zip. If the LLM already
        # produced a full address on the Location row, keep it; otherwise
        # fall back to the IMS street so at least the street matches.
        if "location" in sn and ims_data and idx < len(ims_data):
            ims_row = ims_data[idx]
            ims_addr_key = _find_header_key(ims_row, ["address"])
            loc_addr_key = _find_header_key(row, ["address"])
            if ims_addr_key and loc_addr_key:
                ims_street = str(ims_row.get(ims_addr_key, "")).strip()
                loc_addr = str(row.get(loc_addr_key, "")).strip()
                if ims_street and ims_street.lower() not in loc_addr.lower():
                    # LLM produced a different address — prepend IMS street so
                    # the street matches; keep any city/state/zip tail if present.
                    tail = ""
                    if "," in loc_addr:
                        tail = loc_addr.split(",", 1)[1].strip()
                    row[loc_addr_key] = (
                        f"{ims_street}, {tail}" if tail else ims_street
                    )

        # ---- Other sub-sheets: "Location" column inherits full Location addr
        if "location" not in sn and location_data and idx < len(location_data):
            loc_ref = location_data[idx]
            loc_addr_key = _find_header_key(loc_ref, ["address"])
            row_loc_key = _find_header_key(row, ["location"])
            if loc_addr_key and row_loc_key and "address" not in row_loc_key.lower():
                row[row_loc_key] = loc_ref.get(loc_addr_key, row.get(row_loc_key, ""))

        # ---- Inland Marine: copy Class Lookup/Code/Construction from Property
        if "inland marine" in sn or "inland" in sn:
            if property_data and idx < len(property_data):
                prop_row = property_data[idx]
                for cand in [["class lookup"], ["class code"], ["construction"]]:
                    src_key = _find_header_key(prop_row, cand)
                    dst_key = _find_header_key(row, cand)
                    if src_key and dst_key:
                        row[dst_key] = prop_row[src_key]

        # ---- Property sheet: H (Building Subaddress) = street-only IMS addr
        #      S column mirrors H; V (Total exposure) = T column value.
        if "property" in sn and "sub" not in sn:
            if ims_data and idx < len(ims_data):
                ims_row = ims_data[idx]
                ims_addr_key = _find_header_key(ims_row, ["address"])
                # H = "building subaddress" or similar
                sub_key = _find_header_key(row, ["building subaddress", "subaddress"])
                if ims_addr_key and sub_key:
                    street = str(ims_row.get(ims_addr_key, "")).strip()
                    # Strip state/zip tail if LLM supplied full address
                    if "," in street:
                        street = street.split(",")[0].strip()
                    row[sub_key] = street

        # ---- Loss Payee/Mortgagee consistency with Policy ------------------
        if policy_data and idx < len(policy_data):
            policy_row = policy_data[idx]
            lp_key = _find_header_key(row, ["loss payee", "mortgagee"])
            name_key = _find_header_key(policy_row, ["name"])
            if lp_key and name_key:
                policy_name = str(policy_row.get(name_key, "")).strip()
                if policy_name and policy_name.lower() not in ["nan", ""]:
                    row[lp_key] = policy_name

    return rows


# ---------------------------------------------------------------------------
# LOB conditional filtering
# ---------------------------------------------------------------------------

LOB_SHEET_MAP = {
    "property": ["property"],
    "crime": ["crime"],
    "inland marine": ["inland marine", "inland"],
    "commercial auto": ["commercial auto", "comm auto", "auto"],
}

# All header keyword variants the LLM might use for each LOB column in the
# Policy sheet.  Order matters — more specific terms first.
_LOB_HEADER_VARIANTS: dict[str, list[str]] = {
    "property": [
        "property",
    ],
    "crime": [
        "crime",
    ],
    "inland marine": [
        "inland marine", "inland",
    ],
    "commercial auto": [
        "commercial auto", "comm auto",
        # "auto" alone is intentionally last and only used when nothing else
        # matched, because "auto" appears in other column names too.
    ],
}

# Headers that look like "auto" but are NOT the Commercial Auto LOB toggle
# (e.g. "Auto-Renewal", "Automobile", column suffixes like "(Col 4)").
_AUTO_FALSE_POSITIVES = {"auto-", "automobile", "automatic"}


def _find_lob_key(row: dict[str, Any], lob_name: str) -> str | None:
    """Return the Policy-row key that represents the on/off toggle for *lob_name*.

    Tries variants from most-specific to least-specific so that e.g. the
    'Commercial Auto' column is found before a generic 'auto' substring match.
    For the 'commercial auto' LOB we apply extra guards to avoid false-positive
    matches on unrelated headers that happen to contain 'auto'.
    """
    variants = _LOB_HEADER_VARIANTS.get(lob_name, [lob_name])
    row_keys = list(row.keys())

    for variant in variants:
        for k in row_keys:
            kl = k.lower().strip()
            if variant not in kl:
                continue
            # Guard: skip false positives for the 'auto' short variant
            if variant == "auto" and any(fp in kl for fp in _AUTO_FALSE_POSITIVES):
                continue
            # Guard: the key should not contain unrelated LOB names that would
            # make this a wrong match (e.g. "General Liability" when looking for
            # "property").
            return k

    return None


def extract_lob_flags(policy_data: list[dict[str, Any]]) -> dict[int, dict[str, bool]]:
    """From POLICY sheet data, extract LOB flags per DS row index.

    Only an explicit "no" (case-insensitive) disables the LOB.  Blanks, N/A,
    or any other value default to True — this prevents a sub-sheet from being
    wiped out just because the LLM left a Policy LOB cell empty.
    """
    flags: dict[int, dict[str, bool]] = {}
    for idx, row in enumerate(policy_data):
        row_flags: dict[str, bool] = {}
        for lob_name in _LOB_HEADER_VARIANTS:
            key = _find_lob_key(row, lob_name)
            if key:
                val = str(row.get(key, "")).strip().lower()
                row_flags[lob_name] = val != "no"
                if val == "no":
                    print(f"  LOB flags row {idx}: '{lob_name}' = No (key='{key}')")
            else:
                # Column not found → assume LOB is active (safe default)
                row_flags[lob_name] = True
        flags[idx] = row_flags
    return flags


def filter_rows_by_lob(
    rows: list[dict[str, Any]],
    sheet_name: str,
    lob_flags: dict[int, dict[str, bool]],
) -> list[dict[str, Any]]:
    """Remove rows from a sub-sheet where the corresponding LOB flag is No in Policy.

    If ALL rows are filtered out the sheet is still written to the workbook
    with just the header row, which is the correct behaviour (the LOB is
    disabled for this dataset).
    """
    sn = sheet_name.lower().strip()
    target_lob = None
    for lob_name, keywords in LOB_SHEET_MAP.items():
        if any(kw in sn for kw in keywords):
            target_lob = lob_name
            break

    if target_lob is None:
        return rows

    filtered = []
    for idx, row in enumerate(rows):
        if idx in lob_flags:
            if lob_flags[idx].get(target_lob, True):
                filtered.append(row)
            # else: LOB=No for this row → skip it
        else:
            # Row index beyond what Policy sheet had → include by default
            filtered.append(row)

    return filtered


# ---------------------------------------------------------------------------
# Prompt-building helpers
# ---------------------------------------------------------------------------

def _per_sheet_reminder(sheet_name: str) -> str:
    """Short, surgical reminders for the handful of high-defect sheets.

    These don't duplicate the full rule book (the user already supplies that
    via special instructions). They call out the most frequent rule-book
    violations we've seen in output.
    """
    sn = sheet_name.lower().strip()
    notes: list[str] = []

    if "policy" in sn:
        notes += [
            "• Company and Program columns: ALWAYS 'NetRate'.",
            "• Rate and Rules columns: ALWAYS equal the Effective Date value.",
            "• General Liability column: ALWAYS 'Yes'. Excess Liability: ALWAYS 'No'. Umbrella: ALWAYS 'No'.",
            "• Property / Crime / Inland Marine / Commercial Auto LOB columns: default to 'Yes' unless the user says otherwise. Never leave these blank.",
            "• Corporation column: ALWAYS 'Corporation'.",
            "• Protection Class: MUST be an integer from 1 to 10 (1 = best protection, 10 = unprotected). Never use Yes/No or text.",
            "• Property Damage Deductible: pick from dropdown 250, 500, 1000, 2500, 5000, 10000, 25000, 50000, 100000 — no other values.",
            "• Liability Limit: random integer between 10000 and 10000000. Liability Deductible: random integer between 250 and 100000.",
            "• Liability Symbol: random integer between 2 and 9.",
            "• Deductible Property Damage Only (col M): ALWAYS blank — never populate with any value.",
            "• Type (col V): must be exactly one of: N/A, Apartment, Contractor, Mercantile, Motel/Hotel, Office, Institutional, Industrial/Processing — no spaces around the slash.",
        ]
    elif "ims" in sn and "screen" in sn:
        notes += [
            "• Business Name and Name on Policy: ALWAYS the same authentic company name for a given row.",
            "• Address (col E): STREET ONLY — '[House Number] [Street Name] [Street Type] [Optional Directional]'. NO city, state, or ZIP.",
            "• ZIP (col F): MUST be exactly 5 digits as a STRING, zero-padded — e.g. '08618' not '8618', '02116' not '2116'. NEVER drop the leading zero.",
            "• State (col J): full state name, not abbreviation (e.g. 'New York' not 'NY'). Must match the state implied by the address/ZIP.",
            "• Line: ALWAYS 'Package'. Company: ALWAYS 'Acme'. Billing Type: ALWAYS 'Agency Bill'.",
            "• Type (col B): ALWAYS 'Corporation'.",
            "• Producer by Location (col G): must be one of the exact dropdown values. Valid examples: 'Acme Producer- NewYork - Jericho, NY', 'Ash Group Prospect -Ewing, NJ', 'Cigna Group -Washington, DC', 'CRC Insurance Services Excess MPL -Jericho, NY', 'Dextorville -Brooklyn, NY', 'Evergreen Shield Insurance -New York, NY', 'Iron Gate Insurance Group -Brooklyn, NY', 'James Smith - NewYork - Jericho, NY', 'Mutual\\'s -Coral Gables, FL', 'Pinnacle Assurance -Washington, DC', 'PrimeCover Associates -Beverly Hills, CA', 'SecureEdge Covers - Jericho, NY', 'SSI Producer - Brooklyn, NY', 'Statewide\\'s -New York, NY', 'Summit Ridge Risk Solutions -Las Vegas, NV', 'TCI Boston -Boston, MA', 'Test Producer_01_17 -Brooklyn, NY', 'Test Producer_2 -Shreveport, LA', 'The Armour\\'s -Las Vegas, NV', 'Ul Agency -Brooklyn, NY', 'Unity Risk Management - Shreveport, LA'. NEVER invent new producer names.",
        ]
    elif "location" in sn:
        notes += [
            "• Description: ALWAYS 'loc #1'.",
            "• Address (col C): FULL address in this exact format: '[House Number] [Street Name] [Street Type] [City], [ST] [ZIP]' — NO comma between street type and city (e.g. '218 Snowy Ridge Rd Fairbanks, AK 99709'). Street part must match the IMS sheet address for the same row.",
            "• Insured Location: ALWAYS 'Yes'.",
        ]
    elif "property" in sn and "sub" not in sn and "inland" not in sn:
        notes += [
            "• Class Lookup (D) and Class Code (E) MUST come from the SAME tuple in the user's dropdown — the code in E must appear inside the parentheses of the value in D.",
            "• Number of Stories (L): MUST be exactly 1, 2, or 3 — NEVER any other number, NEVER Yes/No.",
            "• Coinsurance (X): MUST be exactly 80, 90, or 100 — NEVER Yes/No, NEVER a percentage string.",
            "• Inflation Guard (AB): MUST be exactly one of: 2%, 3%, 4%, 6%, 8%, 10%, N/A — NEVER Yes/No.",
            "• Cause of Loss (W): MUST be exactly 'Basic', 'Special', or 'Broad' — NEVER 'FIRE' or any other value.",
            "• Deductible (Y): pick ONE of 250 / 500 / 1000 / 2500 / 5000 / 10000 — integer only, no '%'.",
            "• Watchman Service (AZ): ALWAYS 'None'. Alarm (BA): ALWAYS 'None'. NEVER Yes/No for these fields.",
            "• Tenant Improvements: ALWAYS '$0' — NEVER blank.",
            "• AF: ALWAYS 'No'. AG and AH: ALWAYS 0. AI: ALWAYS empty.",
            "• AK–AO: ALWAYS empty. AP = sum of AJ through AO (AK–AO are empty so AP = AJ).",
            "• V (Total Exposure) MUST equal T (Building Value) exactly — copy the same dollar value.",
            "• BP: 0 or blank. BQ: 'N/A'. BR: '100%'. BS: '$0'. BT–BW: 'N/A'. BX–BZ: '$0'. CA–CE: 'No'.",
            "• CF: '$0'. CG: 'N/A'. CH–CN: '$0' (literally '$0', nothing else). CQ: 'N/A'. CR: '5%'. CS: 'N/A'. CT: 0. CU: 'No'. CV: '$0'. CW, CX: 'N/A'. CY, CZ: empty.",
            "• Agreed Value (col AE, col AY): must be 'Yes' or 'No' only — NEVER a dollar amount or numeric value.",
            "• Agreed Value (col BO): ALWAYS 'No'.",
            "• Business Income Type (col BG): must be exactly one of: Mercantile, Manufacturing, Rental Properties (these are IMS values, NOT Coinsurance/Monthly Indemnity/etc.).",
            "• Civil Authority (col BL): always N/A.",
            "• Time Period (col BM): must be one of: 72-Hour Waiting, 48-Hour Waiting, 24-Hour Waiting, 12-Hour Waiting, No Waiting — NEVER '90 Days', '180 Days', '365 Days', or 'No Limitation'.",
            "• Ordinary Payroll (col BN): must be one of: Included, Excluded, Excluded - 30 Days, Excluded - 60 Days, Excluded - 90 Days (NEVER Yes or No).",
            "• Personal Property Total Exposure (col AP) must exactly equal Contents (col AJ) — they MUST match.",
        ]
    elif "inland" in sn:
        notes += [
            "• Class Lookup (C), Class Code (D), Construction (E): COPY EXACTLY from the PROPERTY sheet for the same row.",
            "• Every coverage toggle (G, P, X, AF, AL, AT, BM, BU, BZ, CC): when 'No', every dependent limit/deductible column MUST be blank — do NOT put any value in dependent columns when the parent is 'No'.",
        ]
    elif "crime" in sn:
        notes += [
            "• Columns D (Description) and E (Class Code) auto-derive from C (Class Lookup) — pick one tuple consistently.",
            "• Coverage is Written (col F): must be exactly one of: Primary, Excess, Concurrent (NOT Coinsurance or any other value).",
            "• CONDITIONAL BLANKING — this is mandatory: for EVERY coverage checkbox (G, S, AA, AE, AH, AM, AV, BC, BF, BL), when its value is 'No', ALL of its dependent columns (limits, deductibles, employees, premises, etc.) MUST be empty string ''. Do NOT populate dependent fields for disabled coverages.",
            "• When any coverage checkbox = Yes, ALL dependent child fields must be populated with valid values. NEVER leave a child field blank when the parent is 'Yes'.",
            "• Specifically: if 'Employee Theft and Forgery' = No → blank Ratable Employees, Additional Premises, Faithful Performance, Limit of Insurance (Theft), Deductible (Theft), Include Expenses, Credit Card columns, Limit of Insurance (Forgery), Deductible (Forgery), Clients Property columns.",
            "• If 'Fraud Impersonation' = No → blank all its Ratable Employees, Employees, Verification, Limit, Deductible columns.",
            "• If 'Robbery' = No → blank Type, Limit, Deductible for Robbery.",
            "• If 'Theft of Money or Securities' = No → blank its Limit, Deductible, Increased Limit, Number of Premises, Date Beginning, Date Ending.",
            "• Apply same logic to all other coverage checkboxes and their dependents.",
        ]
    elif "general li" in sn:
        notes += [
            "• Sheet name may be misspelled 'LIBIALITY' in the template — treat as General Liability.",
            "• Class of Business (M): MUST be an EXACT string from the user's dropdown list, including the code in parentheses e.g. 'Bakeries (10100)'. NEVER invent values.",
            "• Class Code (N): MUST be the numeric code extracted from the same M tuple e.g. if M='Bakeries (10100)' then N=10100.",
            "• Every value in C, D, E, F, I, J, K is a strict dropdown — pick exact strings from the user's instruction list. No invented values.",
            "• C (Each Occurrence) and D (General Aggregate): D must be ≥ C. E (Products Aggregate): same valid values as D.",
        ]
    elif "commercial auto" in sn or "comm auto" in sn:
        notes += [
            "• Registration State (AN): MUST match the 2-letter state abbreviation of the Location address for that row.",
            "• Columns D–I only populate when Composite Basis (C) is NOT 'No Composite'.",
            "• Vehicle Type (AQ): pick EXACT string from the user's dropdown — NEVER invent vehicle types.",
            "• Deductible fields (AS, AB, AC): pick EXACT values from the user's dropdown lists.",
            "• Uninsured (AD, AZ): pick EXACT values from the user's dropdown lists — e.g. '500,000' not '50,000'.",
        ]

    if not notes:
        return ""
    return "\nPER-SHEET REMINDERS (from the rule book — do not violate):\n" + "\n".join(notes)


def _build_configurable_constraints(pi: ParsedInstructions) -> str:
    lines: list[str] = []

    lines.append(f"DATE POLICY: {_build_date_policy_summary(pi)}")

    lines.append(
        f"DATA SET IDs: Use prefix '{pi.ds_prefix}' starting at {pi.ds_start} "
        f"(e.g. {pi.ds_prefix}{pi.ds_start}, {pi.ds_prefix}{pi.ds_start + 1}, ...)."
    )

    if pi.allowed_states:
        lines.append(
            f"STATES: Only generate data for these US states: {', '.join(pi.allowed_states)}."
        )

    if pi.unique_addresses:
        lines.append(
            "ADDRESSES: Every Data Set row must have a unique street address. "
            "Do not reuse the same address across DS rows."
        )
    else:
        lines.append("ADDRESSES: Address reuse across DS rows is permitted.")

    if pi.submitted_date_mode == "match":
        lines.append("SUBMITTED DATE: Must equal the Effective Date for each row.")
    elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
        lines.append(f"SUBMITTED DATE: Must always be {_format_mmddyyyy(pi.submitted_date_static)}.")
    else:
        lines.append("SUBMITTED DATE: Must be today's date at time of generation.")

    if pi.pct_format == "symbol":
        lines.append(
            "PERCENTAGES: Write percentage values as strings with the % symbol "
            "(e.g. '2%', '5%', '10%'). Do NOT use decimal fractions (0.02, 0.05)."
        )
    else:
        lines.append(
            "PERCENTAGES: Write percentage values as decimal fractions (e.g. 0.02 for 2%)."
        )

    return "\n".join(f"  • {line}" for line in lines)


# ---------------------------------------------------------------------------
# OpenAI client
# ---------------------------------------------------------------------------

def configure_openai() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise ValueError("OPENAI_API_KEY not found")
    return OpenAI(api_key=key)


# ---------------------------------------------------------------------------
# Main generation entry point
# ---------------------------------------------------------------------------

def generate_test_data(
    headers: list[str],
    row_count: int,
    special_instruction: str,
    sheet_name: str = "",
    previous_sheets_data: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """
    Generate test data using OpenAI for US insurance domain.

    All business rules come from the user's special instructions — nothing is
    hardcoded. The code only enforces structural correctness: dates, DS numbering,
    cross-sheet consistency, and LOB conditional filtering.
    """
    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "")

    pi = parse_special_instructions(special_instruction, row_count=row_count)

    # ---- cross-sheet context -----------------------------------------------
    previous_context = ""
    if previous_sheets_data:
        previous_context = (
            "\nCRITICAL — CROSS-SHEET DATA CONSISTENCY:\n"
            "The following sheets have already been generated. You MUST reuse the SAME values "
            "for matching fields across sheets. Row N in this sheet corresponds to Row N in all "
            "other sheets (same insured / policy / location).\n\n"
            "Fields that MUST be identical across sheets (when the column exists):\n"
            "  - Data Set IDs (DS_1, DS_2, etc.) — must match row-for-row\n"
            "  - Effective dates, Expiration dates — IDENTICAL across sheets for the same DS\n"
            "  - Names (Insured Name, Contact Name)\n"
            "  - Addresses (Street, City, State, ZIP)\n"
            "  - Class Lookup, Class Code — for Inland Marine, copy EXACTLY from PROPERTY sheet\n"
            "  - Construction — for Inland Marine, copy EXACTLY from PROPERTY sheet\n"
            "  - Loss Payee/Mortgagee — must match the Name in POLICY sheet\n"
            "  - Location — must match Location sheet address\n\n"
            "Previously generated data:\n"
        )
        for prev_sheet, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- Sheet: '{prev_sheet}' ---\n"
            previous_context += json.dumps(prev_data, indent=2)
            previous_context += "\n"

    # ---- For very large column sets, split into column groups ---------------
    # Aggressive chunking: use 30 cols/chunk for >100-col sheets (PROPERTY has 105)
    # to keep each LLM call well inside the JSON-mode completion budget.
    if len(headers) > 100:
        max_cols = 30
    elif len(headers) > 50:
        max_cols = 40
    else:
        max_cols = 50

    if len(headers) > max_cols:
        return _generate_chunked(
            client, model_name, headers, row_count, pi, sheet_name,
            previous_sheets_data, previous_context, max_cols,
        )

    # ---- assemble prompt ---------------------------------------------------
    user_instruction_block = pi.raw.strip() or "No additional instructions provided."
    configurable_constraints = _build_configurable_constraints(pi)

    sheet_focus = _per_sheet_reminder(sheet_name)

    prompt = f"""You are a test data generator for US insurance applications (IMS / NetRate system).

Sheet: "{sheet_name}"

CRITICAL: Return EXACTLY {row_count} row objects in the "data" array. Not fewer.
Generate realistic test data for these columns:
{json.dumps(headers)}

DUPLICATE COLUMN HANDLING:
Columns with a "(Col X)" suffix are separate columns that happen to share a name in the original
spreadsheet. Generate DIFFERENT values for each. Use ALL column names (with suffixes) as JSON keys.

{previous_context}

INSTRUCTION PRIORITY (strictly followed in order):
1. USER SPECIAL INSTRUCTIONS — these are the authoritative rules; every "always X",
   dropdown list, and cross-sheet directive in this block is NON-NEGOTIABLE:
{user_instruction_block}

2. STRUCTURED CONSTRAINTS (derived automatically from the instructions above):
{configurable_constraints}
{sheet_focus}

3. CROSS-SHEET CONSISTENCY — row N in this sheet = row N in previously generated
   sheets. Copy DS, Effective, Expiration, Submitted, Name, Address exactly.

4. DEFAULT RULES — apply only when not overridden by anything above:
  • Generate realistic US insurance test data.
  • Phone: (XXX) XXX-XXXX | SSN: XXX-XX-XXXX | ZIP: 5-digit or ZIP+4
  • State: 2-letter abbreviation (CA, TX, NY …)
  • Dates: MM/DD/YYYY format.
  • Dollar values: comma-formatted strings with $ sign (e.g. $1,000,000 not $1000000)
  • Data Set IDs: {pi.ds_prefix}{pi.ds_start}, {pi.ds_prefix}{pi.ds_start + 1} … (zero-padded to 2 digits if < 100 rows)
  • When a field says "ALWAYS X" or "always X" in the special instructions — use that exact value, no exceptions.
  • When a field says "EMPTY" or "empty" — leave it as empty string "".
  • Dropdown/enum fields: ONLY use values from the specified list. NEVER invent new values.
  • Percentages: DO NOT use "%" for fields that expect integer dollar amounts or counts — read the instruction for that column.

Return ONLY a JSON object with this exact shape:
{{"data": [{{<row 1>}}, {{<row 2>}}, ... exactly {row_count} objects ...]}}
No markdown, no explanation, no commentary.
"""

    return _call_llm_and_post_process(
        client, model_name, prompt, headers, sheet_name, pi, previous_sheets_data
    )


def _generate_chunked(
    client: Any,
    model_name: str,
    headers: list[str],
    row_count: int,
    pi: ParsedInstructions,
    sheet_name: str,
    previous_sheets_data: dict[str, Any] | None,
    previous_context: str,
    max_cols: int,
) -> list[dict[str, Any]]:
    """
    For sheets with many columns (e.g. PROPERTY with 105 cols), generate data
    in multiple LLM calls, then merge. Each chunk shares the Data Set column
    to keep rows aligned.
    """
    # Always include Data Set in every chunk
    ds_header = None
    for h in headers:
        if "data set" in h.lower():
            ds_header = h
            break

    remaining = [h for h in headers if h != ds_header]
    chunks: list[list[str]] = []
    for i in range(0, len(remaining), max_cols - (1 if ds_header else 0)):
        chunk = remaining[i : i + max_cols - (1 if ds_header else 0)]
        if ds_header:
            chunk = [ds_header] + chunk
        chunks.append(chunk)

    print(f"Sheet '{sheet_name}' has {len(headers)} cols → splitting into {len(chunks)} chunks")

    user_instruction_block = pi.raw.strip() or "No additional instructions provided."
    configurable_constraints = _build_configurable_constraints(pi)
    sheet_focus = _per_sheet_reminder(sheet_name)

    merged_rows: list[dict[str, Any]] = [{} for _ in range(row_count)]

    for chunk_idx, chunk_headers in enumerate(chunks):
        prompt = f"""You are a test data generator for US insurance applications (IMS / NetRate system).

Sheet: "{sheet_name}" — CHUNK {chunk_idx + 1} of {len(chunks)}

CRITICAL: Return EXACTLY {row_count} row objects. Not fewer. Not more.
Generate realistic test data for ONLY these columns:
{json.dumps(chunk_headers)}

This is part of a larger sheet. Row N here corresponds to row N in other chunks.

{previous_context}

INSTRUCTION PRIORITY:
1. USER SPECIAL INSTRUCTIONS (authoritative rules):
{user_instruction_block}

2. STRUCTURED CONSTRAINTS:
{configurable_constraints}
{sheet_focus}

3. CROSS-SHEET CONSISTENCY — use values from previously generated sheets for matching rows.

4. DEFAULT RULES:
  • Dates: MM/DD/YYYY. Effective = today. Expiration = Effective + 1 year.
  • Dollar values: comma-formatted with $ (e.g. $1,000,000)
  • Data Set IDs: {pi.ds_prefix}{pi.ds_start}, {pi.ds_prefix}{pi.ds_start + 1} …
  • Follow dropdown/enum constraints from special instructions exactly.

Return ONLY a JSON object with this exact shape:
{{"data": [{{<row 1>}}, {{<row 2>}}, ...]}}
The "data" array MUST contain exactly {row_count} objects. No markdown, no commentary.
"""

        # Retry the chunk up to 3 times if the model returns fewer rows than asked.
        chunk_data: list[dict[str, Any]] = []
        for attempt in range(3):
            chunk_data = _call_llm_and_post_process(
                client, model_name, prompt, chunk_headers, sheet_name, pi,
                previous_sheets_data, skip_cross_sheet=True,
            )
            if len(chunk_data) >= row_count:
                break
            print(
                f"  chunk {chunk_idx + 1}: got {len(chunk_data)}/{row_count} rows "
                f"(attempt {attempt + 1}/3) — retrying"
            )

        # Never silently drop rows. Pad with empty-value rows if the model
        # still fell short after retries so downstream code stays aligned.
        if len(chunk_data) < row_count:
            print(
                f"  WARNING: chunk {chunk_idx + 1} of '{sheet_name}' "
                f"returned {len(chunk_data)}/{row_count} rows after retries; padding."
            )
            while len(chunk_data) < row_count:
                chunk_data.append({h: "" for h in chunk_headers})

        # Merge chunk data into the merged rows
        for row_idx, chunk_row in enumerate(chunk_data[:row_count]):
            merged_rows[row_idx].update(chunk_row)

    # Fill in any missing headers with empty string so every row has every key
    for row in merged_rows:
        for h in headers:
            if h not in row:
                row[h] = ""

    # Final sanity check: we MUST have exactly row_count rows
    if len(merged_rows) != row_count:
        raise RuntimeError(
            f"Chunked generation for '{sheet_name}' produced "
            f"{len(merged_rows)} rows, expected {row_count}"
        )

    # Now run date enforcement, sheet-specific rules, and cross-sheet on the full merged result
    merged_rows = _enforce_effective_expiration_date_range(merged_rows, pi)
    merged_rows = _enforce_sheet_business_rules(merged_rows, sheet_name)
    merged_rows = enforce_cross_sheet_consistency(
        merged_rows, sheet_name, previous_sheets_data
    )

    return merged_rows


def _call_llm_and_post_process(
    client: Any,
    model_name: str,
    prompt: str,
    headers: list[str],
    sheet_name: str,
    pi: ParsedInstructions,
    previous_sheets_data: dict[str, Any] | None,
    skip_cross_sheet: bool = False,
) -> list[dict[str, Any]]:
    """Make the LLM call, parse response, normalize, and post-process."""
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
                            "Always return a JSON object with a single key \"data\" containing an array "
                            "of row objects. User special instructions override all defaults. "
                            "Follow dropdown/enum constraints EXACTLY — never invent values not in the list."
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
            )

            content = response.choices[0].message.content
            text: str = content if isinstance(content, str) else ""
            if not text.strip():
                raise ValueError("OpenAI returned an empty response")

            data = _parse_json_array(text)

            # Normalise: ensure every header key exists in every row
            normalized: list[dict[str, Any]] = []
            for row in data:
                if isinstance(row, dict):
                    normalized.append({h: row.get(h, "") for h in headers})
            if not normalized:
                raise ValueError("Model output did not contain valid row objects")

            # Post-process: enforce dates
            normalized = _enforce_effective_expiration_date_range(normalized, pi)

            # Post-process: sheet-specific business rules (dropdown enforcement,
            # conditional blanking, fixed-value fields, etc.)
            normalized = _enforce_sheet_business_rules(normalized, sheet_name)

            # Post-process: cross-sheet consistency
            if not skip_cross_sheet:
                normalized = enforce_cross_sheet_consistency(
                    normalized, sheet_name, previous_sheets_data
                )

            return normalized

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "quota" in error_msg.lower() or "rate" in error_msg.lower():
                wait_time = 60 * (attempt + 1)
                print(
                    f"Rate limit hit for sheet '{sheet_name}'. "
                    f"Waiting {wait_time}s before retry {attempt + 1}/{max_retries}…"
                )
                time.sleep(wait_time)
                if attempt == max_retries - 1:
                    raise
            else:
                raise

    raise RuntimeError("OpenAI response was not received after retries")
