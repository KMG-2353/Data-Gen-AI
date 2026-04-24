"""IMS (commercial insurance) handler.

Ported from origin/IMS-new-business. Owns:
  - ParsedInstructions dataclass + parse_special_instructions
  - LOB conditional filtering (Policy sheet toggles sub-sheets)
  - Sheet enforcers (IMS screen, Policy, Property, Crime)
  - Cross-sheet consistency enforcement
  - Configurable-constraints + per-sheet reminder prompt augmentation
  - Sheet-name canonicalization (misspelling fixes)
  - Date-range enforcement (IMS flavor — uses ParsedInstructions)
"""
from __future__ import annotations

import random
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any


# ---------------------------------------------------------------------------
# Sheet-name canonicalization
# ---------------------------------------------------------------------------

SHEET_NAME_FIXES = {
    "netrate general libiality": "NETRATE GENERAL LIABILITY",
    "general libiality": "GENERAL LIABILITY",
    "netrate commericial auto": "NETRATE COMMERCIAL AUTO",
    "commericial auto": "COMMERCIAL AUTO",
}


def canonical_sheet_name(name: str) -> str:
    key = name.lower().strip()
    return SHEET_NAME_FIXES.get(key, name)


# ---------------------------------------------------------------------------
# ParsedInstructions + parser
# ---------------------------------------------------------------------------

@dataclass
class ParsedInstructions:
    date_range_start: int | None = None
    date_range_end: int | None = None
    distribute_all_years: bool = False
    cap_expiration_to_range: bool = True

    ds_prefix: str = "DS_"
    ds_start: int = 1

    allowed_states: list[str] = field(default_factory=list)
    active_lobs: list[str] = field(default_factory=list)
    unique_addresses: bool = True

    submitted_date_mode: str = "current"
    submitted_date_static: date | None = None

    pct_format: str = "symbol"
    raw: str = ""


def parse_special_instructions(text: str, row_count: int = 0) -> ParsedInstructions:
    pi = ParsedInstructions(raw=text)
    if not text:
        return pi

    t = text.lower()

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

    all_years_patterns = [
        r"\ball\s+years\b",
        r"\ball\s+\d+\s+years?\b",
        r"\bevery\s+year\b",
        r"\beach\s+year\b",
        r"\byears?\s+from\b",
        r"\binclude\s+all\s+years?\b",
        r"\binclude\s+all\s+\d+\s+years?\b",
        r"\bto\s+be\s+included\b",
        r"\bone\s+per\s+year\b",
        r"\bspread\s+(?:across|over)\s+years?\b",
        r"\bdistribute\s+(?:across|over)\s+years?\b",
        r"\binclusive\)?",
    ]
    keyword_hit = any(re.search(p, t) for p in all_years_patterns)

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

    ds_prefix_match = re.search(r"ds\s+prefix[:\s]+([A-Za-z0-9_]+)", t)
    if ds_prefix_match:
        pi.ds_prefix = ds_prefix_match.group(1).upper() + "_"

    ds_start_match = re.search(r"ds\s+start(?:ing)?\s+(?:from\s+)?(\d+)", t)
    if ds_start_match:
        pi.ds_start = int(ds_start_match.group(1))

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

    if any(p in t for p in ["percent as decimal", "pct as decimal",
                              "percentage as decimal", "use decimal for percent"]):
        pi.pct_format = "decimal"

    if any(p in t for p in ["reuse address", "allow duplicate address",
                              "same address allowed"]):
        pi.unique_addresses = False

    # Parse LOB selection injected by the UI.
    # Format: "SELECTED LOBs (Policy sheet): General Liability = No, Crime = Yes, ..."
    # Each LOB has its own "LOBName = Yes/No" token so we can parse exactly.
    _KNOWN_IMS_LOBS = [
        "general liability", "property", "crime",
        "inland marine", "commercial auto", "optional coverage",
    ]
    if "selected lob" in t:
        for lob_name in _KNOWN_IMS_LOBS:
            pattern = re.escape(lob_name) + r"\s*=\s*yes"
            if re.search(pattern, t):
                pi.active_lobs.append(lob_name)

    return pi


# ---------------------------------------------------------------------------
# Date helpers & enforcement
# ---------------------------------------------------------------------------

def _format_mmddyyyy(value: date) -> str:
    return value.strftime("%m/%d/%Y")


def _add_one_year(value: date) -> date:
    try:
        return value.replace(year=value.year + 1)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + 1)


def _find_header_key(row: dict[str, Any], candidates: list[str]) -> str | None:
    for key in row.keys():
        key_lower = key.lower()
        if any(c in key_lower for c in candidates):
            return key
    return None


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


def enforce_effective_expiration_date_range(
    rows: list[dict[str, Any]],
    pi: ParsedInstructions,
) -> list[dict[str, Any]]:
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

        eff_dt = today
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

            if eff_key:
                row[eff_key] = _format_mmddyyyy(eff_dt)
            if exp_key:
                row[exp_key] = _format_mmddyyyy(exp_dt)

            if rr_key:
                row[rr_key] = _format_mmddyyyy(eff_dt)

        if sub_key:
            if pi.submitted_date_mode == "current":
                sub_dt = today
            elif pi.submitted_date_mode == "match" and eff_key:
                sub_dt = eff_dt
            elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
                sub_dt = pi.submitted_date_static
            else:
                sub_dt = today
            row[sub_key] = _format_mmddyyyy(sub_dt)

    return rows


# ---------------------------------------------------------------------------
# Property sheet enforcer
# ---------------------------------------------------------------------------

_VALID_STORIES = ["1", "2", "3"]
_VALID_COINSURANCE = ["80", "90", "100"]
_VALID_INFLATION_GUARD = ["2%", "3%", "4%", "6%", "8%", "10%", "N/A"]
_VALID_CAUSE_OF_LOSS = ["Basic", "Special", "Broad"]

_PROPERTY_EXACT: dict[str, Any] = {
    "Building Ordinance":                "No",
    "Building Ordinance (B)":            0,
    "Building Ordinance (C)":            0,
    "Personal Property Sub-Description": "",
    "Property of Insured":               "",
    "Property of Others":                "",
    "Improvements":                      "",
    "Stock":                             "",
    "Watchman Service":                  "None",
    "Alarm":                             "None",
    "Civil Authority":                   "N/A",
    "Income with extra Expense":         "No",
    "Dependent Properties Limit":        "0",
    "Dependent Properties Form":         "N/A",
    "Limit Percentage":                  "100%",
    "Days in Deductible":               "N/A",
    "Extra Expense Limit":               "$0",
    "Liability":                         "N/A",
    "Limits on Loss Payment":            "N/A",
    "Utility Service Coverage Provided": "N/A",
    "Public or Other":                   "N/A",
    "Building Sub-Limit":                "$0",
    "Personal Property Sub-Limit":       "$0",
    "Earthquake Sub-Limit":              "$0",
    "Water Supply":                      "No",
    "Communication Supply":              "No",
    "Power Supply":                      "No",
    "Power Lines":                       "No",
    "Communication Lines":               "No",
    "Pollutant Cleanup and Removal":     "$0",
    "Pollutant Cleanup Deductible":      "N/A",
    "Debris Removal Over 10K":           "$0",
    "Fire Legal Building":               "$0",
    "Fire Legal Contents":               "$0",
    "Spoilage Coverage Limit":           "$0",
    "Peak Season Additional Limit 1":    "$0",
    "Peak Season Additional Limit 2":    "$0",
    "Vacancy Permit Exposure":           "$0",
    "Earthquake Coverage":               "N/A",
    "Masonry Veneer Limit":              "N/A",
    "Roof Tank on Building":             "No",
    "Sublimit Amount":                   "$0",
    "Sublimit Percentage":               "N/A",
    "Sprinkler Leakage":                 "N/A",
    "Building Construction":             "",
    "Building Description":              "",
}

_PROPERTY_BY_COL: dict[int, Any] = {
    98: 0,
    96: "5%",
}

_PROPERTY_COPY_PAIRS: list[tuple[int, int]] = [
    (23, 43), (24, 44), (25, 45), (26, 46), (27, 47), (28, 48), (29, 49),
]


def _col_num(key: str) -> int | None:
    m = re.search(r"\(col\s+(\d+)\)", key.lower())
    return int(m.group(1)) if m else None


def _enforce_property_sheet_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for row in rows:
        keys = list(row.keys())

        for k in keys:
            bare = re.sub(r"\s*\(col\s+\d+\)", "", k, flags=re.IGNORECASE).strip()
            if k in _PROPERTY_EXACT:
                row[k] = _PROPERTY_EXACT[k]
            elif bare in _PROPERTY_EXACT and bare != k:
                if bare not in ("Other",):
                    row[k] = _PROPERTY_EXACT[bare]

        for k in keys:
            if k.strip() == "Other":
                row[k] = ""

        for k in keys:
            col_n = _col_num(k)
            if col_n in _PROPERTY_BY_COL:
                row[k] = _PROPERTY_BY_COL[col_n]

        for k in keys:
            if "number of stories" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n >= 90:
                continue
            try:
                n = int(str(row[k]).strip())
                if str(n) not in _VALID_STORIES:
                    row[k] = random.choice(_VALID_STORIES)
            except (ValueError, TypeError):
                row[k] = random.choice(_VALID_STORIES)

        for k in keys:
            if "coinsurance" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue
            val = str(row[k]).strip()
            if val not in _VALID_COINSURANCE:
                row[k] = random.choice(_VALID_COINSURANCE)

        for k in keys:
            if "inflation guard" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue
            val = str(row[k]).strip()
            if val not in _VALID_INFLATION_GUARD:
                row[k] = random.choice(_VALID_INFLATION_GUARD)

        for k in keys:
            if "cause of loss" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n > 30:
                continue
            val = str(row[k]).strip()
            if val not in _VALID_CAUSE_OF_LOSS:
                row[k] = random.choice(_VALID_CAUSE_OF_LOSS)

        for k in keys:
            if "tenant improvement" in k.lower():
                val = str(row.get(k, "")).strip()
                if val in ("", "nan", "None", "none"):
                    row[k] = "$0"

        bldg_exp_key = None
        total_v_key = None
        for k in keys:
            kl = k.lower()
            if "building exposure" in kl:
                bldg_exp_key = k
            elif "total exposure" in kl:
                col_n = _col_num(k)
                if col_n is None or col_n <= 30:
                    total_v_key = k
        if bldg_exp_key and total_v_key:
            row[total_v_key] = row.get(bldg_exp_key, row.get(total_v_key, ""))

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

        for k in keys:
            if "agreed value" not in k.lower():
                continue
            col_n = _col_num(k)
            if col_n is not None and col_n >= 65:
                row[k] = "No"
                continue
            val = str(row.get(k, "")).strip()
            if val and val not in ("Yes", "No", ""):
                row[k] = "Yes"
            elif not val:
                row[k] = "No"

        for k in keys:
            if "civil authority" in k.lower():
                row[k] = "N/A"

        for k in keys:
            kl = k.lower()
            if "ordinance or law" in kl and "increase period" in kl:
                val = str(row.get(k, "")).strip()
                if val not in ("Yes", "No"):
                    row[k] = random.choice(["Yes", "No"])

        for k in keys:
            if "days in deductible" in k.lower():
                row[k] = "N/A"

        _VALID_ORD_PAYROLL = ["No Limitation", "0 days", "90 days", "180 days"]
        for k in keys:
            if "ordinary payroll" in k.lower():
                val = str(row.get(k, "")).strip()
                if val not in _VALID_ORD_PAYROLL:
                    row[k] = random.choice(_VALID_ORD_PAYROLL)

        _VALID_BI_TYPE = ["Mercantile", "Manufacturing", "Rental Properties"]
        for k in keys:
            if "business income type" in k.lower():
                val = str(row.get(k, "")).strip()
                if val not in _VALID_BI_TYPE:
                    row[k] = random.choice(_VALID_BI_TYPE)

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
                    pass
                elif val.lower() in _TP_MAP:
                    row[k] = _TP_MAP[val.lower()]
                elif val:
                    row[k] = "72-Hour Waiting"

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
# Crime sheet enforcer
# ---------------------------------------------------------------------------

_CRIME_PARENT_DEPS: list[tuple[list[str], list[str]]] = [
    (["employee theft and forgery", "employee theft"],
     ["ratable employees", "additional premises", "faithful performance",
      "limit of insurance", "deductible amount", "deductible limit",
      "include expenses", "credit card", "clients property"]),
    (["fraud impersonation"],
     ["ratable employees", "additional premises", "employees for",
      "verification option", "in excess of limit", "limit of insurance",
      "deductible amount"]),
    (["robbery"], ["type for", "limit of insurance", "deductible amount"]),
    (["outside the premises"], ["limit of insurance", "deductible amount"]),
    (["destruction of electronic data", "electronic data"],
     ["ratable employees", "additional premises", "limit of insurance", "deductible amount"]),
    (["guests property"],
     ["safe deposit", "number of rooms", "room or apartment", "increased limit per guest",
      "food or liquor", "laundry", "articles for sale"]),
    (["theft of money", "money or securities"],
     ["limit of insurance", "deductible", "increased limit", "number of premises",
      "date beginning", "date ending"]),
    (["money orders and counterfeit", "money orders"],
     ["(col 56)", "(col 57)"]),
    (["computer and fund transfer"],
     ["ratable employees", "additional premises", "limit of insurance", "deductible for"]),
    (["identity fraud expense"],
     ["ratable employees", "additional premises", "limit of insurance", "deductible"]),
]


def _enforce_crime_conditional_fields(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    def _crime_child_matches(kl: str, child_keywords: list[str]) -> bool:
        for ck in child_keywords:
            if ck == "deductible":
                if re.match(r"^deductible(\s+\(col\s+\d+\))?$", kl):
                    return True
            else:
                if ck in kl:
                    return True
        return False

    _CRIME_YES_FILL: list[tuple[list[str], list[tuple[list[str], Any]]]] = [
        (["employee theft and forgery", "employee theft"], [
            (["ratable employees"], lambda: str(random.randint(10, 75))),
            (["additional premises"], lambda: str(random.randint(1, 3))),
            (["limit of insurance"], lambda: random.choice(["$50,000", "$75,000", "$100,000", "$150,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000", "2500"])),
        ]),
        (["fraud impersonation"], [
            (["ratable employees"], lambda: str(random.randint(10, 75))),
            (["additional premises"], lambda: str(random.randint(1, 3))),
            (["limit of insurance"], lambda: random.choice(["$50,000", "$75,000", "$100,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000", "2500"])),
        ]),
        (["robbery"], [
            (["limit of insurance"], lambda: random.choice(["$25,000", "$50,000", "$75,000", "$100,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000"])),
        ]),
        (["outside the premises"], [
            (["limit of insurance"], lambda: random.choice(["$15,000", "$25,000", "$50,000"])),
            (["deductible"], lambda: random.choice(["250", "500"])),
        ]),
        (["destruction of electronic data", "electronic data"], [
            (["ratable employees"], lambda: str(random.randint(10, 60))),
            (["additional premises"], lambda: str(random.randint(1, 3))),
            (["limit of insurance"], lambda: random.choice(["$10,000", "$20,000", "$30,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000"])),
        ]),
        (["theft of money", "money or securities"], [
            (["limit of insurance"], lambda: random.choice(["$10,000", "$25,000", "$50,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000"])),
            (["increased limit"], lambda: random.choice(["$5,000", "$10,000", "$15,000"])),
        ]),
        (["money orders and counterfeit", "money orders"], [
            (["(col 56)"], lambda: random.choice(["$2,500", "$5,000", "$7,500"])),
            (["(col 57)"], lambda: random.choice(["100", "250", "500"])),
        ]),
        (["computer and fund transfer"], [
            (["ratable employees"], lambda: str(random.randint(10, 60))),
            (["additional premises"], lambda: str(random.randint(1, 3))),
            (["limit of insurance"], lambda: random.choice(["$25,000", "$50,000", "$75,000"])),
            (["deductible"], lambda: random.choice(["500", "1000", "2500"])),
        ]),
        (["identity fraud expense"], [
            (["ratable employees"], lambda: str(random.randint(10, 50))),
            (["additional premises"], lambda: str(random.randint(1, 2))),
            (["limit of insurance"], lambda: random.choice(["$10,000", "$15,000", "$25,000"])),
            (["deductible"], lambda: random.choice(["250", "500", "1000"])),
        ]),
    ]

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
            for k in keys:
                if k == parent_key:
                    continue
                kl = k.lower()
                if any(dk in kl for dk in dep_keywords):
                    row[k] = ""

    for row in rows:
        for k in list(row.keys()):
            if "coverage is written" in k.lower():
                val = str(row.get(k, "") or "").strip()
                if val not in ("Primary", "Excess", "Concurrent"):
                    row[k] = "Primary"

    return rows


# ---------------------------------------------------------------------------
# Policy & IMS Screen enforcers
# ---------------------------------------------------------------------------

_VALID_POLICY_TYPES = [
    "N/A", "Apartment", "Contractor", "Mercantile",
    "Motel/Hotel", "Office", "Institutional", "Industrial/Processing"
]


def _enforce_policy_sheet_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for row in rows:
        for k in list(row.keys()):
            if "deductible property damage" in k.lower():
                row[k] = ""
            if "transportation network" in k.lower() or "on-demand capacity" in k.lower():
                row[k] = ""
            if k.strip().lower() == "garagekeepers":
                row[k] = "No"
            if k.strip().lower() == "type":
                val = str(row.get(k, "") or "").strip()
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


_STATE_ABBREV_TO_NAME: dict[str, str] = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District of Columbia",
}


def enforce_state_selection(
    rows: list[dict[str, Any]],
    sheet_name: str,
    selected_states: list[str],
) -> list[dict[str, Any]]:
    """Deterministically overwrite state columns with the user-selected states.

    IMS SCREEN uses full state names (col J); other sheets use 2-letter codes.
    States are distributed round-robin across rows so coverage is even.
    """
    if not selected_states or not rows:
        return rows

    sn = sheet_name.lower().strip()
    is_ims_screen = "ims" in sn and "screen" in sn

    # Normalise selected_states to uppercase abbreviations
    abbrevs = [s.strip().upper() for s in selected_states if s.strip()]
    if not abbrevs:
        return rows

    # For IMS SCREEN we need full names; for other sheets keep abbreviations
    if is_ims_screen:
        state_values = [_STATE_ABBREV_TO_NAME.get(a, a) for a in abbrevs]
    else:
        state_values = abbrevs

    # Find the state column key in the row (avoid "registration state", "state code", etc.)
    def _find_state_key(row: dict) -> str | None:
        # Priority 1: exact "state" key
        for k in row:
            if k.lower().strip() == "state":
                return k
        # Priority 2: key that IS "state" with a column suffix like "State (Col 10)"
        for k in row:
            kl = k.lower().strip()
            if kl.startswith("state (col "):
                return k
        # Priority 3: key that starts with "state" but is not a compound like "statewide"
        for k in row:
            kl = k.lower().strip()
            if kl.startswith("state") and "registration" not in kl and len(kl) <= 20:
                return k
        return None

    result = []
    for idx, row in enumerate(rows):
        new_row = dict(row)
        state_key = _find_state_key(new_row)
        if state_key:
            new_row[state_key] = state_values[idx % len(state_values)]
        result.append(new_row)
    return result


def _enforce_ims_screen_values(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
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
                    row[k] = val.split(",")[0].strip()
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
# Cross-sheet consistency
# ---------------------------------------------------------------------------

def enforce_cross_sheet_consistency(
    rows: list[dict[str, Any]],
    sheet_name: str,
    previous_sheets_data: dict[str, list[dict[str, Any]]] | None,
) -> list[dict[str, Any]]:
    if not previous_sheets_data or not rows:
        return rows

    sn = sheet_name.lower().strip()

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

        eff_key = _find_header_key(row, ["effective", "effetive"])
        rr_key  = _find_header_key(row, ["rates & rules", "rates and rules",
                                         "rate and rule", "rate & rule"])
        if eff_key and rr_key:
            row[rr_key] = row[eff_key]

        if "location" in sn and ims_data and idx < len(ims_data):
            ims_row = ims_data[idx]
            ims_addr_key = _find_header_key(ims_row, ["address"])
            loc_addr_key = _find_header_key(row, ["address"])
            if ims_addr_key and loc_addr_key:
                ims_street = str(ims_row.get(ims_addr_key, "")).strip()
                loc_addr = str(row.get(loc_addr_key, "")).strip()
                if ims_street and ims_street.lower() not in loc_addr.lower():
                    tail = ""
                    if "," in loc_addr:
                        tail = loc_addr.split(",", 1)[1].strip()
                    row[loc_addr_key] = (
                        f"{ims_street}, {tail}" if tail else ims_street
                    )

        if "location" not in sn and location_data and idx < len(location_data):
            loc_ref = location_data[idx]
            loc_addr_key = _find_header_key(loc_ref, ["address"])
            row_loc_key = _find_header_key(row, ["location"])
            if loc_addr_key and row_loc_key and "address" not in row_loc_key.lower():
                row[row_loc_key] = loc_ref.get(loc_addr_key, row.get(row_loc_key, ""))

        if "inland marine" in sn or "inland" in sn:
            if property_data and idx < len(property_data):
                prop_row = property_data[idx]
                for cand in [["class lookup"], ["class code"], ["construction"]]:
                    src_key = _find_header_key(prop_row, cand)
                    dst_key = _find_header_key(row, cand)
                    if src_key and dst_key:
                        row[dst_key] = prop_row[src_key]

        if "property" in sn and "sub" not in sn:
            if ims_data and idx < len(ims_data):
                ims_row = ims_data[idx]
                ims_addr_key = _find_header_key(ims_row, ["address"])
                sub_key = _find_header_key(row, ["building subaddress", "subaddress"])
                if ims_addr_key and sub_key:
                    street = str(ims_row.get(ims_addr_key, "")).strip()
                    if "," in street:
                        street = street.split(",")[0].strip()
                    row[sub_key] = street

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
# LOB filtering
# ---------------------------------------------------------------------------

LOB_SHEET_MAP = {
    "property": ["property"],
    "crime": ["crime"],
    "inland marine": ["inland marine", "inland"],
    "commercial auto": ["commercial auto", "comm auto", "auto"],
    "general liability": ["general liability", "netrate general liability"],
    "optional coverage": ["optional coverage", "opt coverage"],
}

_LOB_HEADER_VARIANTS: dict[str, list[str]] = {
    "property": ["property"],
    "crime": ["crime"],
    "inland marine": ["inland marine", "inland"],
    "commercial auto": ["commercial auto", "comm auto"],
    "general liability": ["general liability"],
    "optional coverage": ["optional coverage"],
}

# Frontend LOB names → canonical internal names (case-insensitive matching)
_FRONTEND_LOB_MAP: dict[str, str] = {
    "inland marine": "inland marine",
    "crime": "crime",
    "general liability": "general liability",
    "gl": "general liability",
    "optional coverage": "optional coverage",
    "commercial auto": "commercial auto",
    "property": "property",
}

_AUTO_FALSE_POSITIVES = {"auto-", "automobile", "automatic"}


def _find_lob_key(row: dict[str, Any], lob_name: str) -> str | None:
    variants = _LOB_HEADER_VARIANTS.get(lob_name, [lob_name])
    row_keys = list(row.keys())

    # Pass 1: exact match or exact match with "(Col N)" deduplication suffix.
    # This prevents "property" from matching "deductible property damage", etc.
    for variant in variants:
        for k in row_keys:
            kl = k.lower().strip()
            if kl == variant:
                return k
            if kl.startswith(variant + " (col "):
                return k

    # Pass 2: substring fallback (kept for edge cases with unusual header names)
    for variant in variants:
        for k in row_keys:
            kl = k.lower().strip()
            if variant not in kl:
                continue
            if variant == "auto" and any(fp in kl for fp in _AUTO_FALSE_POSITIVES):
                continue
            return k
    return None


def extract_lob_flags(policy_data: list[dict[str, Any]]) -> dict[int, dict[str, bool]]:
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
                row_flags[lob_name] = True
        flags[idx] = row_flags
    return flags


def filter_rows_by_lob(
    rows: list[dict[str, Any]],
    sheet_name: str,
    lob_flags: dict[int, dict[str, bool]],
) -> list[dict[str, Any]]:
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
        else:
            filtered.append(row)
    return filtered


def enforce_lob_selection(
    policy_data: list[dict[str, Any]],
    selected_lobs: list[str],
    special_instructions: str = "",
) -> list[dict[str, Any]]:
    """Override LOB Yes/No columns in Policy rows based on UI selection.

    selected_lobs: LOB names from the frontend (e.g., ["Crime", "General Liability"]).
    Special instructions only block a LOB override when they contain an EXPLICIT
    Yes/No directive for that LOB (e.g. "crime = no", "property: yes").
    If selected_lobs is empty, returns data unchanged (random LLM output is kept).
    """
    if not selected_lobs:
        return policy_data

    # Map frontend LOB names → canonical internal keys
    selected_canonical: set[str] = set()
    for lob in selected_lobs:
        canon = _FRONTEND_LOB_MAP.get(lob.lower().strip())
        if canon:
            selected_canonical.add(canon)

    # Determine which LOBs have an EXPLICIT override in special instructions
    # (e.g. "crime = no", "property: yes") — only those are left untouched.
    si_lower = (special_instructions or "").lower()
    si_overridden: set[str] = set()
    for lob_name in _LOB_HEADER_VARIANTS:
        pattern = re.escape(lob_name) + r"\s*[=:]\s*(yes|no)"
        if re.search(pattern, si_lower):
            si_overridden.add(lob_name)

    # Build a flat lookup: lowercased column name → target "Yes"/"No"
    # Covers every variant in _LOB_HEADER_VARIANTS that is NOT overridden.
    col_value_map: dict[str, str] = {}
    for lob_name, variants in _LOB_HEADER_VARIANTS.items():
        if lob_name in si_overridden:
            continue
        value = "Yes" if lob_name in selected_canonical else "No"
        for v in variants:
            col_value_map[v] = value  # exact lowercased variant → Yes/No

    result = []
    for row in policy_data:
        new_row = dict(row)
        for k in list(new_row.keys()):
            kl = k.lower().strip()
            # Exact match (e.g. "general liability")
            if kl in col_value_map:
                new_row[k] = col_value_map[kl]
                continue
            # Deduplicated header with "(Col N)" suffix
            for variant, val in col_value_map.items():
                if kl.startswith(variant + " (col "):
                    new_row[k] = val
                    break
        result.append(new_row)
    return result


# ---------------------------------------------------------------------------
# Prompt builders
# ---------------------------------------------------------------------------

def _per_sheet_reminder(sheet_name: str, special_instruction: str = "") -> str:
    sn = sheet_name.lower().strip()
    si = special_instruction.lower()
    notes: list[str] = []

    if "policy" in sn:
        # Determine whether the user has passed an explicit LOB selection.
        # The UI injects "SELECTED LOBs (Policy sheet): ..." into special_instruction.
        _all_lobs = [
            "general liability", "property", "crime",
            "inland marine", "commercial auto", "optional coverage",
        ]
        has_lob_selection = "selected lob" in si
        if has_lob_selection:
            # Build per-LOB directives from the selection hint already in the prompt
            gl_directive = "'Yes' if it appears in the SELECTED LOBs list, otherwise 'No'"
            lob_directive = (
                "Follow the SELECTED LOBs in the special instructions EXACTLY. "
                "Set each LOB column to 'Yes' or 'No' as specified. Never leave blank."
            )
        else:
            gl_directive = "ALWAYS 'Yes'"
            lob_directive = "default to 'Yes' unless the user says otherwise. Never leave these blank."

        notes += [
            "• Company and Program columns: ALWAYS 'NetRate'.",
            "• Rate and Rules columns: ALWAYS equal the Effective Date value.",
            f"• General Liability column: {gl_directive}. Excess Liability: ALWAYS 'No'. Umbrella: ALWAYS 'No'.",
            f"• Property / Crime / Inland Marine / Commercial Auto LOB columns: {lob_directive}",
            "• Corporation column: ALWAYS 'Corporation'.",
            "• Protection Class: MUST be an integer from 1 to 10.",
            "• Property Damage Deductible: pick from dropdown 250, 500, 1000, 2500, 5000, 10000, 25000, 50000, 100000.",
            "• Liability Limit: random integer between 10000 and 10000000. Liability Deductible: random integer between 250 and 100000.",
            "• Liability Symbol: random integer between 2 and 9.",
            "• Deductible Property Damage Only (col M): ALWAYS blank.",
            "• Type (col V): one of N/A, Apartment, Contractor, Mercantile, Motel/Hotel, Office, Institutional, Industrial/Processing.",
        ]
    elif "ims" in sn and "screen" in sn:
        notes += [
            "• Business Name and Name on Policy: same authentic company name for a given row.",
            "• Address (col E): STREET ONLY. NO city, state, or ZIP.",
            "• ZIP (col F): exactly 5 digits as STRING, zero-padded (e.g. '08618').",
            "• State (col J): full state name, not abbreviation.",
            "• Line: ALWAYS 'Package'. Company: ALWAYS 'Acme'. Billing Type: ALWAYS 'Agency Bill'.",
            "• Type (col B): ALWAYS 'Corporation'.",
            "• Producer by Location (col G): must be one of the exact dropdown values.",
        ]
    elif "location" in sn:
        notes += [
            "• Description: ALWAYS 'loc #1'.",
            "• Address (col C): FULL address with city/state/ZIP. Street part must match IMS row.",
            "• Insured Location: ALWAYS 'Yes'.",
        ]
    elif "property" in sn and "sub" not in sn and "inland" not in sn:
        notes += [
            "• Class Lookup (D) and Class Code (E) from SAME tuple.",
            "• Number of Stories (L): 1, 2, or 3.",
            "• Coinsurance (X): 80, 90, or 100.",
            "• Inflation Guard (AB): 2%/3%/4%/6%/8%/10%/N/A.",
            "• Cause of Loss (W): Basic / Special / Broad.",
            "• Deductible (Y): 250/500/1000/2500/5000/10000.",
            "• Watchman Service (AZ): 'None'. Alarm (BA): 'None'.",
            "• Tenant Improvements: '$0'.",
            "• V (Total Exposure) = T (Building Value).",
            "• Time Period (col BM): 72-Hour / 48-Hour / 24-Hour / 12-Hour / No Waiting.",
            "• Business Income Type (col BG): Mercantile / Manufacturing / Rental Properties.",
        ]
    elif "inland" in sn:
        notes += [
            "• Class Lookup, Class Code, Construction: COPY from PROPERTY sheet.",
            "• Every coverage toggle 'No' → dependent columns blank.",
        ]
    elif "crime" in sn:
        notes += [
            "• Coverage is Written (col F): Primary / Excess / Concurrent.",
            "• CONDITIONAL BLANKING: coverage=No → dependent columns empty string.",
            "• Coverage=Yes → all dependent child fields populated.",
        ]
    elif "general li" in sn:
        notes += [
            "• Sheet name may be misspelled 'LIBIALITY' — treat as General Liability.",
            "• Class of Business (M): EXACT string from dropdown with code in parens.",
            "• Class Code (N): numeric code from same M tuple.",
        ]
    elif "commercial auto" in sn or "comm auto" in sn:
        notes += [
            "• Registration State (AN): match Location's 2-letter state.",
            "• Vehicle Type (AQ): EXACT string from user's dropdown.",
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
        lines.append(f"STATES: Only generate data for: {', '.join(pi.allowed_states)}.")
    if pi.unique_addresses:
        lines.append("ADDRESSES: Every Data Set row must have a unique street address.")
    else:
        lines.append("ADDRESSES: Address reuse across DS rows is permitted.")
    if pi.submitted_date_mode == "match":
        lines.append("SUBMITTED DATE: Must equal the Effective Date for each row.")
    elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
        lines.append(f"SUBMITTED DATE: Must always be {_format_mmddyyyy(pi.submitted_date_static)}.")
    else:
        lines.append("SUBMITTED DATE: Must be today's date at time of generation.")
    if pi.pct_format == "symbol":
        lines.append("PERCENTAGES: Use % symbol strings (e.g. '2%').")
    else:
        lines.append("PERCENTAGES: Use decimal fractions (e.g. 0.02 for 2%).")
    if pi.active_lobs:
        _all = ["general liability", "property", "crime",
                "inland marine", "commercial auto", "optional coverage"]
        off = [l.title() for l in _all if l not in pi.active_lobs]
        on  = [l.title() for l in pi.active_lobs]
        lines.append(
            f"LOB COLUMNS (Policy sheet): Set {', '.join(on)} = Yes. "
            f"Set {', '.join(off)} = No."
        )
    return "\n".join(f"  • {line}" for line in lines)


# ---------------------------------------------------------------------------
# Handler
# ---------------------------------------------------------------------------

class ImsHandler:
    policy_type = "IMS"

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = sheet_name.lower().strip()
        if "ims" in sn and "screen" in sn:
            return "ims_screen"
        if "policy" in sn:
            return "policy"
        if "property" in sn and "sub" not in sn and "inland" not in sn:
            return "property"
        if "location" in sn:
            return "location"
        if "crime" in sn:
            return "crime"
        if "inland" in sn:
            return "inland_marine"
        if "general li" in sn:
            return "general_liability"
        if "commercial auto" in sn or "comm auto" in sn:
            return "commercial_auto"
        if "optional" in sn:
            return "optional_coverage"
        return "unknown"

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        pi = parse_special_instructions(special_instruction, row_count=original_row_count)
        constraints = _build_configurable_constraints(pi)
        reminder = _per_sheet_reminder(sheet_name, special_instruction)
        extras = f"\nSTRUCTURED CONSTRAINTS:\n{constraints}{reminder}"
        return original_row_count, extras

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
    ) -> list[dict[str, Any]] | None:
        return None

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        if not rows:
            return rows
        pi = parse_special_instructions(special_instruction, row_count=len(rows))
        rows = enforce_effective_expiration_date_range(rows, pi)
        rows = _enforce_sheet_business_rules(rows, sheet_name)
        rows = enforce_cross_sheet_consistency(rows, sheet_name, previous_sheets_data)
        return rows
