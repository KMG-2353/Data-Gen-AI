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
    # Extracted from phrases like "effective date 2020-2026" or "from 2021 to 2025"
    date_range_start: int | None = None
    date_range_end: int | None = None

    # True when the user asks for every year to appear at least once
    # e.g. "include every year", "each year should appear", "all years 2020-2026"
    distribute_all_years: bool = False

    # True when the user says dates must stay within the declared range
    # (expiration capped at end_year even if that breaks the 1-year gap)
    cap_expiration_to_range: bool = True

    # ---- data set numbering ------------------------------------------------
    # Prefix for DS column: default "DS_", overridable via "DS prefix: INS_"
    ds_prefix: str = "DS_"
    # Starting index for DS numbering
    ds_start: int = 1

    # ---- row distribution / state filtering --------------------------------
    # e.g. "only generate for CA, TX, NY"
    allowed_states: list[str] = field(default_factory=list)

    # ---- LOB flags ---------------------------------------------------------
    # e.g. "only generate Property and GL rows"
    active_lobs: list[str] = field(default_factory=list)

    # ---- address uniqueness ------------------------------------------------
    # True = every DS must have a unique address (default enforced per rule book)
    unique_addresses: bool = True

    # ---- submitted-date behaviour -----------------------------------------
    # "current"  → today's date at generation time (default)
    # "match"    → same as effective date
    # "static:YYYY-MM-DD" → a fixed date
    submitted_date_mode: str = "current"
    submitted_date_static: date | None = None

    # ---- percentage format -------------------------------------------------
    # "symbol"   → "2%" string (default, per IMS rule book)
    # "decimal"  → 0.02 float
    pct_format: str = "symbol"

    # ---- raw pass-through --------------------------------------------------
    # The full original text is always forwarded to the LLM so it can act on
    # anything we haven't structured yet.
    raw: str = ""


def parse_special_instructions(text: str) -> ParsedInstructions:
    """
    Parse a free-text special instruction string into a structured
    ParsedInstructions object.

    Design principles:
    - Flexible: recognises many natural phrasings for the same concept.
    - Non-destructive: unrecognised text is kept in `raw` and forwarded to LLM.
    - Additive: adding a new configurable dimension only requires adding a
      new attribute + a new detection block here; nothing else changes.
    """
    pi = ParsedInstructions(raw=text)
    if not text:
        return pi

    t = text.lower()

    # ---- date range --------------------------------------------------------
    # Matches: "2020-2026", "2020 to 2026", "2020–2026", "from 2020 to 2026"
    date_range_match = re.search(r"(20\d{2})\s*(?:-|to|–|—)\s*(20\d{2})", t)
    if date_range_match:
        s, e = int(date_range_match.group(1)), int(date_range_match.group(2))
        pi.date_range_start = min(s, e)
        pi.date_range_end = max(s, e)

    # Single-year mention without a range: "effective date 2023"
    if pi.date_range_start is None:
        single_year = re.search(r"\b(20\d{2})\b", t)
        if single_year and any(kw in t for kw in ["effective", "expir", "date"]):
            yr = int(single_year.group(1))
            pi.date_range_start = yr
            pi.date_range_end = yr

    # Distribute all years flag
    pi.distribute_all_years = any(phrase in t for phrase in [
        "every year", "each year", "all years", "year from", "years from",
        "include all years", "to be included", "one per year", "spread across years",
    ])

    # Cap expiration to range (default True; user can override)
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
        text,  # use original text to preserve capitalisation
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
# Date-enforcement (uses ParsedInstructions instead of raw text)
# ---------------------------------------------------------------------------

def _build_date_policy_summary(pi: ParsedInstructions) -> str:
    """Return a compact human-readable date policy injected into the LLM prompt."""
    if pi.date_range_start is not None:
        s, e = pi.date_range_start, pi.date_range_end
        if pi.distribute_all_years:
            return (
                f"Effective dates must span ALL years from {s} to {e} (inclusive), "
                f"distributed evenly across rows. Expiration date is exactly 1 year "
                f"after effective{'.' if not pi.cap_expiration_to_range else f', capped at {e}/12/31 if needed.'}"
            )
        return (
            f"Effective dates must fall within {s}–{e} (inclusive). "
            f"Expiration date is exactly 1 year after effective"
            f"{'' if not pi.cap_expiration_to_range else f', capped at {e}/12/31 if the +1 year exceeds the range'}."
        )
    return (
        "No explicit date range given. Use today's date for Effective Date; "
        "Expiration Date is exactly 1 year later."
    )


def _enforce_effective_expiration_date_range(
    rows: list[dict[str, Any]],
    pi: ParsedInstructions,
) -> list[dict[str, Any]]:
    """
    Post-process rows to enforce date policy derived from ParsedInstructions.
    This runs after the LLM generates rows so dates are always correct
    regardless of whether the model followed the instruction perfectly.
    """
    effective_candidates = ["effective date", "effetive date", "effetive", "effective"]
    expiration_candidates = ["expiration date", "expiry date", "expiry", "expiration", "exp date"]
    submitted_candidates = ["submitted"]

    today = date.today()
    has_range = pi.date_range_start is not None

    # Build the year cycle for round-robin distribution
    if has_range:
        start_yr, end_yr = pi.date_range_start, pi.date_range_end  # type: ignore[assignment]
        year_pool = list(range(start_yr, end_yr + 1)) if end_yr > start_yr else [start_yr]
    else:
        year_pool = []

    for idx, row in enumerate(rows):
        eff_key = _find_header_key(row, effective_candidates)
        exp_key = _find_header_key(row, expiration_candidates)
        sub_key = _find_header_key(row, submitted_candidates)

        # ---- effective date ------------------------------------------------
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

            # Cap expiration inside range when requested
            if has_range and pi.cap_expiration_to_range and exp_dt.year > end_yr:
                exp_dt = date(end_yr, eff_dt.month, eff_dt.day)

            if eff_key:
                row[eff_key] = _format_mmddyyyy(eff_dt)
            if exp_key:
                row[exp_key] = _format_mmddyyyy(exp_dt)

        # ---- submitted date ------------------------------------------------
        if sub_key:
            if pi.submitted_date_mode == "current":
                sub_dt = today
            elif pi.submitted_date_mode == "match" and eff_key:
                sub_dt = eff_dt if (eff_key or exp_key) else today  # type: ignore[possibly-undefined]
            elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
                sub_dt = pi.submitted_date_static
            else:
                sub_dt = today
            row[sub_key] = _format_mmddyyyy(sub_dt)

    return rows


# ---------------------------------------------------------------------------
# Prompt-building helpers
# ---------------------------------------------------------------------------

def _build_configurable_constraints(pi: ParsedInstructions) -> str:
    """
    Convert structured ParsedInstructions into a bullet-point constraint block
    that is injected into the LLM prompt. This keeps the prompt clean and makes
    each constraint explicit and traceable.
    """
    lines: list[str] = []

    # Date range
    lines.append(f"DATE POLICY: {_build_date_policy_summary(pi)}")

    # DS numbering
    lines.append(
        f"DATA SET IDs: Use prefix '{pi.ds_prefix}' starting at {pi.ds_start} "
        f"(e.g. {pi.ds_prefix}{pi.ds_start}, {pi.ds_prefix}{pi.ds_start + 1}, ...)."
    )

    # State filter
    if pi.allowed_states:
        lines.append(
            f"STATES: Only generate data for these US states: {', '.join(pi.allowed_states)}."
        )

    # Address uniqueness
    if pi.unique_addresses:
        lines.append(
            "ADDRESSES: Every Data Set row must have a unique street address. "
            "Do not reuse the same address across DS rows."
        )
    else:
        lines.append("ADDRESSES: Address reuse across DS rows is permitted.")

    # Submitted date
    if pi.submitted_date_mode == "match":
        lines.append("SUBMITTED DATE: Must equal the Effective Date for each row.")
    elif pi.submitted_date_mode == "static" and pi.submitted_date_static:
        lines.append(f"SUBMITTED DATE: Must always be {_format_mmddyyyy(pi.submitted_date_static)}.")
    else:
        lines.append("SUBMITTED DATE: Must be today's date at time of generation.")

    # Percentage format
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
# IMS Rule Book — Allowed value sets (from Notion: Data Gen Rule - IMS)
# ---------------------------------------------------------------------------

IMS_PRODUCERS = [
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

PROPERTY_CLASS_LOOKUP = [
    "Amusement Centers (0844)", "Art Studios (0921)", "Book Distributors (0567)",
    "Churches or Other Houses of Worship (0900)", "Drugstores (0562)",
    "Exhibition or Convention Buildings (0843)", "Gift Shops (0567)",
    "Locksmiths (0922)", "Luggage Goods Stores (0567)", "Marble Products - Mfg. (6009)",
    "Masonry - contractors storage (0567)", "Match Mfg. (5100)",
    "Optical Goods - Mfg. (6009)", "Pet Food Mfg. (2059)", "Pet Stores (0570)",
    "Photographers (0921)", "Rental Stores (0567)", "Restaurants - With cooking (0542)",
    "Saunas and Baths - public (0844)", "Saw Mills or Planing Mills (3809)",
    "Sugar Refining (2300)", "Sun Tanning Salons (0844)", "Supermarkets (0532)",
    "Theaters - Drive-in (0833)", "Tire - Mfg. (5759)", "Video Stores (0570)",
    "Waste and Reclaimed Materials (1400)", "Watch Mfg. (6900)",
    "Water Botting (2350)", "Zoos (0844)",
]

PROPERTY_PROTECTION_CLASS = ["1", "2", "5", "8", "8B", "8X"]
PROPERTY_CONSTRUCTION = [
    "Fire Resistive", "Frame", "Joisted Masonry",
    "Masonry Non-Combustible", "Modified Fire Resistive", "Non-Combustible",
]
PROPERTY_BLANKET_OPTION = [
    "Building & Contents", "Building Only", "Building Only & Contents Only",
    "Contents Only", "",
]
PROPERTY_NUM_STORIES = ["1", "2", "3"]
PROPERTY_CAUSE_OF_LOSS = ["Basic", "Special", "Broad"]
PROPERTY_COINSURANCE = ["80", "90", "100"]
PROPERTY_DEDUCTIBLE = ["250", "500", "1000", "2500", "5000", "10000"]
PROPERTY_WIND_HAIL = ["1%", "2%", "5%", "250", "500", "1000", "2500", "N/A", ""]
PROPERTY_THEFT_DEDUCTIBLE = ["250", "500", "1000", "2500", "5000"]
PROPERTY_INFLATION_GUARD = ["2%", "3%", "4%", "6%", "8%", "10%", "N/A"]
PROPERTY_VALUATION = ["ACV", "FRC", "RC"]
PROPERTY_RISK_TYPE = ["Class", "Special", "Specific"]
PROPERTY_YEAR_BUILT = ["1965", "1970", "1972", "1985", "1986", "1992", "1995", "2000", "2001", ""]
PROPERTY_BI_TYPE = ["Mercantile", "Manufacturing", "Rental Properties"]
PROPERTY_BI_OPTION = ["Coinsurance", "Monthly Indemnity", "Maximum Period Indemnity"]
PROPERTY_BI_BLANKET = ["N/A", "Business Income Only"]
PROPERTY_TIME_PERIOD = ["72-Hour Waiting", "48-Hour Waiting", "24-Hour Waiting", "12-Hour Waiting", "No Waiting"]
PROPERTY_ORDINARY_PAYROLL = ["No Limitation", "0 days", "90 days", "180 days"]

POLICY_TYPE = ["N/A", "Apartment", "Contractors", "Industrial/Processing",
               "Institutional", "Mercantile", "Motel/Hotel", "Office", "Service"]

CRIME_CLASS_LOOKUP = [
    "Churches or Other Houses of Worship (8131)", "Abrasive Wheel Mfg. (3270)",
    "Amusement Centers (7130)", "Airports (4880)", "Amusement Parks (7130)",
    "Bakeries (4452)", "Barber Shops (8121)", "Candle Mfg (3395)",
    "Caterers (7225)", "Antique Stores (4530)", "Cosmetics Mfg. (3250)",
    "Distributors Food or Drink (4230)", "Dredging (2330)",
    "Drug Distributors (4222)", "Electronics Stores (4430)",
    "Elevator Mfg. (3330)", "Fiberglass Mfg. (3270)", "Grain Milling (3110)",
    "Ice Cream Stores (4452)", "Metal Scrap Dealers (4210)",
    "Milk Processing (3110)", "Mining -surface (2150)", "Paper Mfg. (3225)",
    "Pet Stores (4530)", "Shaft Sinking (2150)", "Toy Stores (4510)",
    "Wheel Mfg. (3367)",
]
CRIME_COVERAGE_IS_WRITTEN = ["Primary", "Excess", "Coinsurance", "Concurrent"]
CRIME_DEDUCTIBLE = [
    "N/A", "0", "100", "250", "500", "1000", "2500", "5000", "10000",
    "25000", "50000", "75000", "100000", "250000", "500000", "1000000", "2000000",
]
CRIME_FAITHFUL_PERFORMANCE = ["N/A", "Government", "Union", "Union Sole Insured"]
CRIME_INCLUDE_LIMIT_CREDIT = ["N/A", "Include Credit Debit or charge cards", "Limit Covered Instruments"]
CRIME_VERIFICATION_OPTION = [
    "N/A", "Required for all transfer instructions",
    "Required for all transfer excess of an amount", "Not Required",
]
CRIME_ROBBERY_TYPE = ["Blanket", "Scheduled"]

GL_EACH_OCCURRENCE = [
    "25,000", "50,000", "100,000", "200,000", "300,000", "350,000",
    "500,000", "750,000", "1,000,000", "2,000,000", "3,000,000",
    "4,000,000", "5,000,000", "10,000,000",
]
GL_GENERAL_AGGREGATE = [
    "50,000", "100,000", "200,000", "300,000", "350,000", "500,000",
    "600,000", "700,000", "750,000", "900,000", "1,000,000", "1,500,000",
    "2,000,000", "2,250,000", "3,000,000", "4,000,000", "5,000,000",
    "10,000,000", "20,000,000",
]
GL_DEDUCTIBLE = [
    "N/A", "250", "500", "700", "1,000", "2,000", "3,000", "4,000",
    "5,000", "10,000", "15,000", "20,000", "25,000", "50,000", "75,000", "100,000",
]
GL_DEDUCTIBLE_TYPE = ["N/A", "BI", "PD", "BI/PD"]
GL_DEDUCTIBLE_BASIS = ["Per Occurrence", "Per Claim"]
GL_MEDICAL_PAYMENTS = ["Excluded", "2000", "5,000", "10,000"]
GL_POLICY_TYPE = ["Occurrence", "Claims Made"]
GL_CLASS_LOOKUP = [
    "Churches or Other Houses of Worship (41650)", "Abrasive Wheel Mfg. (50010)",
    "Amusement Centers (10015)", "Airports Commercial (40010)",
    "Amusement Parks (10020)", "Bakeries (10100)", "Barber Shops (10113)",
    "Candle Mfg (51741)", "Caterers (11039)", "Stores (11234)",
    "Cosmetics Mfg. (51970)", "Distributors Food or Drink (12361)",
    "Dredging (92055)", "Drug Distributors (12373)", "Electronics Stores (12393)",
    "Elevator Mfg. (52581)", "Fiberglass Mfg. (53121)", "Grain Milling (13621)",
    "Ice Cream Stores (14401)", "Importers (55410)", "Machine Shops (97220)",
    "Metal Scrap Dealers (15406)", "Milk Processing (57002)", "Mining (98002)",
    "Paper Mfg. (57726)", "Pet Stores (16403)", "Rubber Mfg. (58756)",
    "Shaft Sinking (98871)", "Toy Stores (18834)", "Wheel Mfg. (59941)",
]

AUTO_COMPOSITE_BASIS = ["No Composite", "Gross Receipts", "Mileage", "Per Unit", "ACV Unit"]
AUTO_CLASS_LOOKUP = [
    "Abattoirs (311611)", "Administrative courts (922110)", "Air Force (928110)",
    "Animal shelters (812910)", "Automobile rental (532111)",
    "Automotive tire dealers (441320)", "Aviation schools (611512)",
    "Baked ham stores (445240)", "Banks, commercial (522110)",
    "Bar associations (813920)", "Camcorder rental (532210)",
    "Car rental (532111)", "Check room services (812990)",
    "Dairy cattle farming (112120)", "Earthworm hatcheries (112990)",
    "Electrician (238210)", "Family welfare services (624190)",
    "Freight car cleaning services (488210)", "Fruit precooling (115114)",
    "Funeral parlors (812210)",
]
AUTO_VEHICLE_TYPE = [
    "Airport Bus/Limo", "Athlete/Entertainer Bus", "Car Service", "Charter Bus",
    "Church Bus", "Inter-City Bus", "Limousine (8 or less seats)",
    "Limousine (Over 8 seats)", "Paratransit", "Private Passenger",
    "Public Auto NOC", "School Bus (Other)", "School Bus (Political Sub)",
    "Semi-Trailer", "Service or Utility Trailer", "Sightseeing Bus",
    "Social Service Auto (Emp Oper)", "Social Service Auto (Other)",
    "Taxicab (All Other)", "Taxicab (Owner-Driver)", "Trailer", "Truck",
    "Truck-Tractor", "Urban Bus", "Van Pool (Employer Furnished)", "Van Pool (Other)",
]
AUTO_OTC_COVERAGE = [
    "No Coverage", "Comp-All Perils with Full Glass",
    "Comp-All Perils without Full Glass",
    "Comp-Theft, Mischief or Vandalism with Full Glass",
    "Comp-Theft, Mischief or Vandalism without Full Glass",
    "Specified Causes of Loss", "Fire Only", "Fire and Theft Only",
    "Fire, Theft and Windstorm Only", "Limited Specified Causes of Loss",
]
AUTO_OTC_DEDUCTIBLE = ["No Coverage", "50", "100", "200", "250", "500", "1000", "2000", "3000", "5000"]
AUTO_COLLISION_DEDUCTIBLE = ["No Coverage", "50", "100", "250", "500", "1,000", "2,000", "3,000", "5,000"]
AUTO_UNINSURED = [
    "No Coverage", "25/50", "50/100", "100/300", "250/500", "500/1000",
    "50,000", "75,000", "100,000", "125,000", "150,000", "200,000",
    "250,000", "300,000", "350,000", "400,000", "500,000", "750,000",
    "1,000,000", "1,500,000", "2,000,000", "10,000,000",
]
AUTO_MEDICAL_PAYMENTS = ["No Coverage", "500", "1000", "2000", "5000"]
AUTO_LOAN_LEASE_GAP = ["None", "Other Than Collision", "Collision", "Both"]
AUTO_RENTAL_REIMBURSEMENT = [
    "No Coverage", "Specified", "Comprehensive", "Collision",
    "Specified and Collision", "Comprehensive and Collision",
]
AUTO_OTC_COMMON = ["No Coverage", "Full", "50", "100", "1000", "2000", "3000", "5000"]
AUTO_COLLISION_COMMON = ["No Coverage", "100", "250", "500", "1000", "2000", "3000", "5000"]

IM_DEDUCTIBLE_COMMERCIAL_ARTICLES = ["250", "500", "1000", "2500", "5000", "10000"]
IM_DEDUCTIBLE_EDP = ["500", "1000", "1500", "2000", "2500", "3000", "4000", "5000"]
IM_COINSURANCE = ["50%", "60%", "70%", "80%", "90%", "100%"]
IM_RECEPTACLE_AR = [
    "Class A", "Class B", "Class C", "1/2 Hr Exp", "2\" wall or vault",
    "Fully enclosed", "None of the Above",
]
IM_DUPLICATE_RECORDS = ["90%", "51%-89%", "<51%", "N/A"]
IM_ALARM_TYPE = ["AA", "A", "BB", "CC", "C", "None"]
IM_ALARM_QUALITY = ["Basic", "Intermediate", "High"]
IM_LEVEL_OF_PROTECTION = ["Central Station", "Police Connected", "Local"]
IM_CAUSE_OF_LOSS = ["Special", "Basic"]
IM_RECEPTACLE_VP = [
    "Class A Fire/SMNA F1-D Four Hrs Exp", "Class B Fire/SMNA F1-D Two Hrs Exp",
    "Class C Fire/SMNA F1-D One Hr Exp", "UL/SMNA 1/2 Hr Exp",
    "Unlabeled metal safe 2\"", "12\" air space", "None of the Above",
]
IM_DEDUCTIBLE_CAMERA = ["250", "500", "1000", "2500", "5000", "10000", "25000", "50000", "75000", "100000"]
IM_OUTSIDE_DEDUCTIBLE = ["Full Coverage", "Deductible Coverage"]


# ---------------------------------------------------------------------------
# Sheet-specific post-processing functions
# ---------------------------------------------------------------------------

def _find_key_exact(row: dict[str, Any], col_letter_or_name: str, candidates: list[str]) -> str | None:
    """Find row key matching any candidate substring (case-insensitive)."""
    return _find_header_key(row, candidates)


def _validate_against_allowed(value: Any, allowed: list[str]) -> str:
    """Return value if it's in the allowed list, else return a random valid value."""
    str_val = str(value).strip() if value is not None else ""
    if str_val in allowed:
        return str_val
    # Try case-insensitive match
    for a in allowed:
        if a.lower() == str_val.lower():
            return a
    # Return random valid value
    return random.choice(allowed)


def _enforce_property_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce all PROPERTY sheet rules from the IMS Rule Book."""
    if not rows:
        return rows

    for idx, row in enumerate(rows):
        keys_lower = {k.lower(): k for k in row.keys()}

        def _set_by_substr(candidates: list[str], value: Any) -> None:
            key = _find_header_key(row, candidates)
            if key:
                row[key] = value

        def _get_by_substr(candidates: list[str]) -> Any:
            key = _find_header_key(row, candidates)
            return row.get(key, "") if key else ""

        def _validate_field(candidates: list[str], allowed: list[str]) -> None:
            key = _find_header_key(row, candidates)
            if key:
                row[key] = _validate_against_allowed(row.get(key, ""), allowed)

        # Rule 2: Protection Class
        _validate_field(["protection class"], PROPERTY_PROTECTION_CLASS)

        # Rule 5: Risk Type
        _validate_field(["risk type"], PROPERTY_RISK_TYPE)

        # Rule 8: Construction
        _validate_field(["construction"], PROPERTY_CONSTRUCTION)

        # Rule 9: Blanket Option (Col J)
        _validate_field(["blanket option"], PROPERTY_BLANKET_OPTION)

        # Rule 10: Year Built
        _validate_field(["year built"], PROPERTY_YEAR_BUILT)

        # Rule 11: Number of Stories (Col L) — must be 1, 2, or 3
        _validate_field(["number of stories"], PROPERTY_NUM_STORIES)

        # Rule 16: Tenant Improvements (Col U) = $0
        _set_by_substr(["tenant improvement"], "$0")

        # Rule 17: Total Exposure = Building Exposure
        bldg_exp_key = _find_header_key(row, ["building exposure"])
        total_exp_key = _find_header_key(row, ["total exposure"])
        if bldg_exp_key and total_exp_key:
            row[total_exp_key] = row[bldg_exp_key]

        # Rule 18: Cause of Loss
        _validate_field(["cause of loss"], PROPERTY_CAUSE_OF_LOSS)

        # Rule 19: Coinsurance
        for cand in [["coinsurance"]]:
            key = _find_header_key(row, cand)
            if key:
                val = str(row.get(key, "")).strip()
                if val not in PROPERTY_COINSURANCE:
                    row[key] = random.choice(PROPERTY_COINSURANCE)

        # Rule 20: Deductible
        _validate_field(["deductible"], PROPERTY_DEDUCTIBLE)

        # Rule 21: Wind Hail Deductible
        _validate_field(["wind hail"], PROPERTY_WIND_HAIL)

        # Rule 22: Theft Deductible
        _validate_field(["theft deductible"], PROPERTY_THEFT_DEDUCTIBLE)

        # Rule 23: Inflation Guard
        _validate_field(["inflation guard", "inflation gaurd"], PROPERTY_INFLATION_GUARD)

        # Rule 24: Valuation
        _validate_field(["valuation"], PROPERTY_VALUATION)

        # Rule 26: Building Ordinance (Col AF) = ALWAYS NO
        _set_by_substr(["building ordinance"], "No")
        # But don't overwrite Building Ordinance (B) and (C) — handle below

        # Rule 27: Building Ordinance (B) and (C) = 0
        for cand in [["building ordinance (b)", "building ordinance(b)"],
                     ["building ordinance (c)", "building ordinance(c)"]]:
            _set_by_substr(cand, "0")

        # Rule 28: Personal Property Sub-Description (Col AI) = EMPTY
        _set_by_substr(["personal property sub description", "personal property sub-description",
                        "personal property sub desc"], "")

        # Rule 30: Cols AK-AO = EMPTY (Contents, Property of Insured, Property of Others,
        # Improvements, Stock, Other sub-fields)
        for cand in [["property of insured"], ["property of others"],
                     ["improvements"], ["stock"], ["other"]]:
            # Only clear sub-fields under Personal Property, not main fields
            key = _find_header_key(row, cand)
            if key and "personal" in key.lower():
                row[key] = ""

        # Rule 34: Watchman Service = None, Alarm = None
        _set_by_substr(["watchman service"], "None")
        _set_by_substr(["alarm"], "None")

        # Rule 35: Alarm Grade = random A, B, C
        _validate_field(["alarm grade", "alarm guard"], ["A", "B", "C"])

        # Rule 36: Alarm Protect = random 1, 2, 3
        _validate_field(["alarm protect"], ["1", "2", "3"])

        # Rule 37: Business Income Blanket Option
        _validate_field(["business income blanket", "bi blanket"], PROPERTY_BI_BLANKET)

        # Rule 39: Income with Extra Expense (Col BF) = No
        _set_by_substr(["income with extra expense"], "No")

        # Rule 40: Business Income Type — enforce variety
        _validate_field(["business income type"], PROPERTY_BI_TYPE)
        # Force distribution: rotate through types
        bi_key = _find_header_key(row, ["business income type"])
        if bi_key:
            row[bi_key] = PROPERTY_BI_TYPE[idx % len(PROPERTY_BI_TYPE)]

        # Rule 41: Business Income Option
        _validate_field(["business income option"], PROPERTY_BI_OPTION)

        # Rule 42: Option (Col BI) = "None"
        _set_by_substr(["option"], "None")

        # Rule 44: Days in Deductible (Col BK) = N/A
        _set_by_substr(["days in deductible"], "N/A")

        # Rule 45: Civil Authority (Col BL) = N/A
        _set_by_substr(["civil authority"], "N/A")

        # Rule 46: Time Period
        _validate_field(["time period"], PROPERTY_TIME_PERIOD)

        # Rule 46b: Ordinary Payroll
        _validate_field(["ordinary payroll"], PROPERTY_ORDINARY_PAYROLL)

        # Rule 48: Dependent Properties Limit (Col BP) = 0 or blank
        _set_by_substr(["dependent properties limit", "dependent prop"], "0")

        # Rule 49: Dependent Properties Form (Col BQ) = N/A
        _set_by_substr(["dependent properties form"], "N/A")

        # Rule 50: Limit Percentage (Col BR) = 100%
        _set_by_substr(["limit percentage"], "100%")

        # Rule 51: Extra Expense Limit (Col BS) = $0
        _set_by_substr(["extra expense limit"], "$0")

        # Rule 52: Cols BT-BW = N/A (Liability, Limits on Loss Payment,
        # Utility Service Coverage Provided, Public or Other)
        for cand in [["liability"], ["limits on loss"], ["utility service"], ["public or other"]]:
            key = _find_header_key(row, cand)
            if key and any(x in key.lower() for x in ["bt", "bu", "bv", "bw", "liability", "utility", "loss payment"]):
                row[key] = "N/A"

        # Rule 53: Cols BX-BZ = $0 (Building sub-limit, Personal Property sub-limit, Earthquake sub-limit)
        for cand in [["building sub-limit", "building sublimit"],
                     ["personal property sub-limit", "personal property sublimit"],
                     ["earthquake sub-limit", "earthquake sublimit"]]:
            _set_by_substr(cand, "$0")

        # Rule 54: Cols CA-CE = No (Water Supply, Communication Supply, Power Supply, Power Lines, Communication Lines)
        for cand in [["water supply"], ["communication supply"], ["power supply"],
                     ["power lines"], ["communication lines"]]:
            _set_by_substr(cand, "No")

        # Rule 55: Pollutant Cleanup and Removal (Col CF) = $0
        _set_by_substr(["pollutant cleanup and removal", "pollutant clean up and removal"], "$0")

        # Rule 56: Pollutant Cleanup Deductible (Col CG) = N/A
        _set_by_substr(["pollutant cleanup deductible", "pollutant clean up deductible"], "N/A")

        # Rule 57: Cols CH-CN = $0 (Debris Removal Over 10K through Vacancy Permit Exposure)
        for cand in [["debris removal", "debris over"], ["spoil coverage", "spoilage"],
                     ["vacancy permit exposure"]]:
            _set_by_substr(cand, "$0")

        # Rule 60: Deductible (Earthquake) = 5%
        # handled by earthquake deductible field
        _set_by_substr(["earthquake"], "N/A")

        # Rule 61: Masonry Veneer Limit = N/A
        _set_by_substr(["masonry veneer"], "N/A")

        # Rule 62: Number of Stories (second section, Col CT) = 0
        # This is a SECOND "Number of Stories" column — look for (Col CT) or second instance
        # Handled by column suffix logic if present

        # Rule 63: Roof Tank on Building (Col CU) = No
        _set_by_substr(["roof tank"], "No")

        # Rule 64: Sublimit Amount (Col CV) = $0
        _set_by_substr(["sublimit amount"], "$0")

        # Rule 65: Sprinkler Leakage and Sublimit Percentage = N/A
        _set_by_substr(["sprinkler leakage"], "N/A")
        _set_by_substr(["sublimit percentage"], "N/A")

        # Rule 66: Building Construction, Building Description = EMPTY
        _set_by_substr(["building construction"], "")
        _set_by_substr(["building description"], "")

        # Rule 32: AQ-AW (duplicate Cause of Loss set) = same as first set
        # These are already handled by cross-sheet prompt; reinforce here
        for dup_cand, src_cand in [
            (["cause of loss (col", "cause of loss.1"], ["cause of loss"]),
            (["coinsurance (col", "coinsurance.1"], ["coinsurance"]),
            (["deductible (col", "deductible.1"], ["deductible"]),
            (["theft deductible (col", "theft deductible.1"], ["theft deductible"]),
            (["inflation guard (col", "inflation guard.1", "inflation gaurd (col"], ["inflation guard", "inflation gaurd"]),
            (["valuation (col", "valuation.1"], ["valuation"]),
        ]:
            src_key = _find_header_key(row, src_cand)
            dup_key = _find_header_key(row, dup_cand)
            if src_key and dup_key and src_key != dup_key:
                row[dup_key] = row[src_key]

    return rows


def _enforce_policy_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce all POLICY sheet rules from the IMS Rule Book."""
    if not rows:
        return rows

    for row in rows:
        # Rule 1: Company and Program = NetRate
        _set = lambda candidates, val: (
            row.__setitem__(_find_header_key(row, candidates), val)
            if _find_header_key(row, candidates) else None
        )

        # Rule 3: General Liability = Yes
        key = _find_header_key(row, ["general liability"])
        if key:
            row[key] = "Yes"

        # Rule 4: Excess Liability = No
        key = _find_header_key(row, ["excess liability"])
        if key:
            row[key] = "No"

        # Rule 5: Umbrella = No
        key = _find_header_key(row, ["umbrella"])
        if key:
            row[key] = "No"

        # Rule 9: Cols P, Q, U = No
        for cand in [["hired auto physical damage"], ["non-owned auto physical"]]:
            key = _find_header_key(row, cand)
            if key:
                row[key] = "No"

        # Rule 11: Type (Col V) — validate dropdown
        key = _find_header_key(row, ["type"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), POLICY_TYPE)

        # Rule 12: Legal Entity = Corporation
        key = _find_header_key(row, ["legal entity"])
        if key:
            row[key] = "Corporation"

        # Rule 17: Cols M and N = ALWAYS BLANK
        # M = Deductible Property Damage Only, N = Any Vehicles Transportation
        for cand in [["deductible property damage", "property damage only"],
                     ["any vehicles used", "transportation network", "on-demand capacity"]]:
            key = _find_header_key(row, cand)
            if key:
                row[key] = ""

        # Rule 16: If Commercial Auto = No, Liability Limit and Deductible = blank
        auto_key = _find_header_key(row, ["commercial auto"])
        if auto_key and str(row.get(auto_key, "")).strip().lower() == "no":
            for cand in [["liability limit"], ["liability deductible"]]:
                key = _find_header_key(row, cand)
                if key:
                    row[key] = ""

    return rows


def _enforce_crime_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce CRIME sheet rules."""
    if not rows:
        return rows

    for row in rows:
        # Rule 5: Coverage Is Written = Primary/Excess/Coinsurance/Concurrent
        key = _find_header_key(row, ["coverage is written", "coverage written"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), CRIME_COVERAGE_IS_WRITTEN)

        # Rule 9: Faithful Performance of Duty
        key = _find_header_key(row, ["faithful performance"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), CRIME_FAITHFUL_PERFORMANCE)

        # Rule 15: Include/Limit Credit Cards
        key = _find_header_key(row, ["include/limit credit", "include credit", "limit credit"])
        if key and row.get(key, ""):
            row[key] = _validate_against_allowed(row.get(key, ""), CRIME_INCLUDE_LIMIT_CREDIT)

        # Rule 22: Verification Option
        key = _find_header_key(row, ["verification option"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), CRIME_VERIFICATION_OPTION)

        # Rule 27: Robbery Type
        key = _find_header_key(row, ["type"])
        if key and "robbery" in str(_find_header_key(row, ["robbery"])).lower():
            row[key] = _validate_against_allowed(row.get(key, ""), CRIME_ROBBERY_TYPE)

        # Validate all Deductible fields against allowed list
        for k, v in list(row.items()):
            if "deductible" in k.lower():
                str_val = str(v).strip().replace("$", "").replace(",", "")
                if str_val and str_val not in CRIME_DEDUCTIBLE:
                    row[k] = random.choice([d for d in CRIME_DEDUCTIBLE if d not in ["N/A", "0"]])

        # Conditional blank logic: If Employee Theft and Forgery = No,
        # many child fields must be blank
        etf_key = _find_header_key(row, ["employee theft"])
        if etf_key and str(row.get(etf_key, "")).strip().lower() == "no":
            for cand in [["ratable employees"], ["additional premises"],
                         ["limit of insurance"], ["deductible"],
                         ["include/limit credit", "include credit"],
                         ["faithful performance"]]:
                key = _find_header_key(row, cand)
                if key and key != etf_key:
                    row[key] = ""

    return rows


def _enforce_gl_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce General Liability sheet rules."""
    if not rows:
        return rows

    for row in rows:
        # Rule 1: Policy Type
        key = _find_header_key(row, ["policy type"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_POLICY_TYPE)

        # Rule 2: Each Occurrence
        key = _find_header_key(row, ["each occurrence"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_EACH_OCCURRENCE)

        # Rule 3: General Aggregate
        key = _find_header_key(row, ["general aggregate"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_GENERAL_AGGREGATE)

        # Rule 4: Products Aggregate
        key = _find_header_key(row, ["products aggregate"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_GENERAL_AGGREGATE)

        # Rule 5: Deductible
        key = _find_header_key(row, ["deductible"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_DEDUCTIBLE)

        # Rule 6: Deductible Type
        key = _find_header_key(row, ["deductible type"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_DEDUCTIBLE_TYPE)

        # Rule 7: Deductible Basis
        key = _find_header_key(row, ["deductible basis"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_DEDUCTIBLE_BASIS)

        # Rule 8: Medical Payments
        key = _find_header_key(row, ["medical payment"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), GL_MEDICAL_PAYMENTS)

        # Rule 12: Class — must include code with name
        key = _find_header_key(row, ["class"])
        if key and "code" not in key.lower():
            val = str(row.get(key, ""))
            if val and "(" not in val:
                row[key] = _validate_against_allowed(val, GL_CLASS_LOOKUP)

    return rows


def _enforce_auto_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce Commercial Auto sheet rules."""
    if not rows:
        return rows

    for row in rows:
        # Rule 2: Composite Basis
        key = _find_header_key(row, ["composite basis"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_COMPOSITE_BASIS)

        # Rule 11: Class Lookup
        key = _find_header_key(row, ["class lookup"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_CLASS_LOOKUP)

        # Rule 13: Description = match Class Lookup
        desc_key = _find_header_key(row, ["description"])
        cl_key = _find_header_key(row, ["class lookup"])
        if desc_key and cl_key:
            cl_val = str(row.get(cl_key, ""))
            # Extract name without code
            name_part = re.sub(r"\s*\(\d+\)\s*$", "", cl_val).strip()
            row[desc_key] = name_part if name_part else row.get(desc_key, "")

        # Rule 31: OTC Coverage (Col AA)
        key = _find_header_key(row, ["other than collision coverage", "otc coverage"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_OTC_COVERAGE)

        # Rule 32: OTC Deductible (Col AB)
        key = _find_header_key(row, ["other than collision deductible", "otc deductible"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_OTC_DEDUCTIBLE)

        # Rule 33: Collision Deductible (Col AC)
        key = _find_header_key(row, ["collision deductible"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_COLLISION_DEDUCTIBLE)

        # Rule 34: Uninsured (Col AD)
        key = _find_header_key(row, ["uninsured"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_UNINSURED)

        # Rule 47: Vehicle Type (Col AQ)
        key = _find_header_key(row, ["vehicle type"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_VEHICLE_TYPE)

        # Rule 48: OTC Coverage (vehicle-level, Col AR)
        key = _find_header_key(row, ["otc coverage"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_OTC_COVERAGE)

        # Rule 52: Auto Loan/Lease Gap
        key = _find_header_key(row, ["auto loan", "lease gap"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_LOAN_LEASE_GAP)

        # Rule 54: Medical Payments
        key = _find_header_key(row, ["medical payment"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_MEDICAL_PAYMENTS)

        # Rule 56: Uninsured (vehicle-level, Col AZ)
        # Already handled by "uninsured" check above — both Col AD and AZ

        # Rule 61: Rental Reimbursement
        key = _find_header_key(row, ["rental reimbursement"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_RENTAL_REIMBURSEMENT)

        # Common coverages OTC (Col S)
        key = _find_header_key(row, ["other than collision"])
        if key and "deductible" not in key.lower() and "coverage" not in key.lower():
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_OTC_COMMON)

        # Common coverages Collision (Col T)
        key = _find_header_key(row, ["collision"])
        if key and "deductible" not in key.lower():
            row[key] = _validate_against_allowed(row.get(key, ""), AUTO_COLLISION_COMMON)

        # Tapes, Records and Discs = Yes/No checkbox
        key = _find_header_key(row, ["tapes", "record", "discs"])
        if key:
            val = str(row.get(key, "")).strip().lower()
            if val not in ["yes", "no"]:
                row[key] = random.choice(["Yes", "No"])

        # Rate Override = Yes/No
        key = _find_header_key(row, ["rate override"])
        if key:
            val = str(row.get(key, "")).strip().lower()
            if val not in ["yes", "no"]:
                row[key] = random.choice(["Yes", "No"])

    return rows


def _enforce_inland_marine_rules(
    rows: list[dict[str, Any]],
    pi: ParsedInstructions,
    property_data: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Enforce Inland Marine sheet rules, including cross-sheet consistency with Property."""
    if not rows:
        return rows

    for idx, row in enumerate(rows):
        # Rules 2-4: Class Lookup, Class Code, Construction must match PROPERTY sheet
        if property_data and idx < len(property_data):
            prop_row = property_data[idx]
            for cand in [["class lookup"], ["class code"], ["construction"]]:
                src_key = _find_header_key(prop_row, cand)
                dst_key = _find_header_key(row, cand)
                if src_key and dst_key:
                    row[dst_key] = prop_row[src_key]

        # Validate conditional fields: checkbox-dependent blanking
        for parent_cand, child_cands in [
            (["accounts receivable"], [
                ["nonreporting limit"], ["away from premise"], ["receptacle"],
                ["classification of risk"], ["reporting limit"], ["branch premise"],
                ["duplicate records"], ["coinsurance"],
            ]),
            (["commercial articles"], [
                ["deductible"], ["organ"], ["commercial limit"],
                ["motion pictures"], ["individual-professional"],
                ["other than individual"],
            ]),
            (["physicians and surgeons"], [
                ["deductible"], ["coinsurance"], ["additional coverage"],
                ["artificially generated energy deductible"],
                ["limited property coverage"], ["equipment limit"],
                ["artificially generated energy limit"],
            ]),
            (["valuable papers"], [
                ["deductible"], ["scheduled limit"], ["away from premise"],
                ["receptacle"], ["unscheduled limit"],
            ]),
            (["electronic data processing"], [
                ["deductible"], ["data and media"], ["special revisions"],
                ["base rate"], ["electronic equipment"], ["extra expense limit"], ["limit"],
            ]),
            (["equipment dealers"], [
                ["deductible"], ["coinsurance"], ["outside building"],
                ["additional covered property"], ["monthly payment"],
                ["inside building"], ["elsewhere limit"],
            ]),
            (["signs coverage"], [
                ["outside deductible"], ["inside deductible"],
                ["outside coverage"], ["inside coverage"],
            ]),
        ]:
            parent_key = _find_header_key(row, parent_cand)
            if parent_key and str(row.get(parent_key, "")).strip().lower() == "no":
                for child_cand in child_cands:
                    child_key = _find_header_key(row, child_cand)
                    if child_key and child_key != parent_key:
                        row[child_key] = ""

        # Validate enum fields
        key = _find_header_key(row, ["premise alarm type", "alarm type"])
        if key and row.get(key, ""):
            row[key] = _validate_against_allowed(row.get(key, ""), IM_ALARM_TYPE)

        key = _find_header_key(row, ["alarm quality"])
        if key and row.get(key, ""):
            row[key] = _validate_against_allowed(row.get(key, ""), IM_ALARM_QUALITY)

        key = _find_header_key(row, ["level of protection"])
        if key and row.get(key, ""):
            row[key] = _validate_against_allowed(row.get(key, ""), IM_LEVEL_OF_PROTECTION)

        key = _find_header_key(row, ["cause of loss"])
        if key:
            row[key] = _validate_against_allowed(row.get(key, ""), IM_CAUSE_OF_LOSS)

    return rows


def _enforce_ims_screen_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce IMS Screen sheet rules."""
    if not rows:
        return rows

    for row in rows:
        # Rule K: Company = always Acme
        key = _find_header_key(row, ["company"])
        if key:
            row[key] = "Acme"

        # Rule I: Line = always Package
        key = _find_header_key(row, ["line"])
        if key:
            row[key] = "Package"

        # Rule L: Billing Type = Agency Bill
        key = _find_header_key(row, ["billing type", "billing"])
        if key:
            row[key] = "Agency Bill"

        # Rule B: Legal Entity = Corporation
        key = _find_header_key(row, ["legal entity"])
        if key:
            row[key] = "Corporation"

        # Rule G: Producer by Location must be from approved list
        key = _find_header_key(row, ["producer by location", "producer"])
        if key:
            val = str(row.get(key, "")).strip()
            if val and val not in IMS_PRODUCERS:
                row[key] = random.choice(IMS_PRODUCERS)

        # Rule M: Underwriter format (last name, first name)
        key = _find_header_key(row, ["underwriter"])
        if key:
            val = str(row.get(key, "")).strip()
            allowed_underwriters = ["Khan, Nadeem", "Khanra, Kushal", "Singh, Ravendra"]
            if val not in allowed_underwriters:
                row[key] = random.choice(allowed_underwriters)

    return rows


def _enforce_optional_coverage_rules(rows: list[dict[str, Any]], pi: ParsedInstructions) -> list[dict[str, Any]]:
    """Enforce Optional Coverage sheet rules."""
    if not rows:
        return rows

    for idx, row in enumerate(rows):
        # Rule 1: Line of Business = Liability
        key = _find_header_key(row, ["line of business"])
        if key:
            row[key] = "Liability"

        # Rule 2: Coverage = Stop Gap
        key = _find_header_key(row, ["coverage"])
        if key:
            row[key] = "Stop Gap"

        # Rule 5: Rate = 0.95
        key = _find_header_key(row, ["rate"])
        if key:
            row[key] = "0.95"

        # DS must be populated
        ds_key = _find_header_key(row, ["data set", "ds"])
        if ds_key and not str(row.get(ds_key, "")).strip():
            row[ds_key] = f"{pi.ds_prefix}{str(pi.ds_start + idx).zfill(2)}"

    return rows


def _enforce_cross_sheet_loss_payee(
    rows: list[dict[str, Any]],
    sheet_name: str,
    policy_data: list[dict[str, Any]] | None,
) -> list[dict[str, Any]]:
    """Ensure Loss Payee/Mortgagee matches between sheets and Policy."""
    if not policy_data or not rows:
        return rows

    for idx, row in enumerate(rows):
        if idx >= len(policy_data):
            break
        policy_row = policy_data[idx]

        # Property Rule 67 / Auto Rule 64: Loss Payee must match Policy Name
        lp_key = _find_header_key(row, ["loss payee", "mortgagee"])
        name_key = _find_header_key(policy_row, ["name"])
        if lp_key and name_key:
            policy_name = str(policy_row.get(name_key, "")).strip()
            if policy_name:
                row[lp_key] = policy_name

    return rows


# ---------------------------------------------------------------------------
# Master post-processor — dispatches to sheet-specific enforcers
# ---------------------------------------------------------------------------

def post_process_sheet(
    rows: list[dict[str, Any]],
    sheet_name: str,
    pi: ParsedInstructions,
    previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
) -> list[dict[str, Any]]:
    """
    Run all sheet-specific post-processing rules. Called after LLM generation
    and after date enforcement, to deterministically correct any rule violations.
    """
    sn = sheet_name.lower().strip()

    if "ims" in sn and "screen" in sn:
        rows = _enforce_ims_screen_rules(rows, pi)

    if "policy" in sn:
        rows = _enforce_policy_rules(rows, pi)

    if "property" in sn:
        rows = _enforce_property_rules(rows, pi)

    if "crime" in sn:
        rows = _enforce_crime_rules(rows, pi)

    if "general liability" in sn or ("gl" in sn and "screen" not in sn):
        rows = _enforce_gl_rules(rows, pi)

    if "commercial auto" in sn or "comm" in sn and "auto" in sn:
        rows = _enforce_auto_rules(rows, pi)

    if "inland marine" in sn:
        property_data = None
        if previous_sheets_data:
            for k, v in previous_sheets_data.items():
                if "property" in k.lower():
                    property_data = v
                    break
        rows = _enforce_inland_marine_rules(rows, pi, property_data)

    if "optional" in sn and "coverage" in sn:
        rows = _enforce_optional_coverage_rules(rows, pi)

    # Cross-sheet: Loss Payee/Mortgagee consistency
    if previous_sheets_data:
        policy_data = None
        for k, v in previous_sheets_data.items():
            if "policy" in k.lower():
                policy_data = v
                break
        if policy_data:
            rows = _enforce_cross_sheet_loss_payee(rows, sheet_name, policy_data)

    return rows


# ---------------------------------------------------------------------------
# LOB filtering — remove rows where Policy LOB flag = No
# ---------------------------------------------------------------------------

LOB_SHEET_MAP = {
    "property": ["property"],
    "crime": ["crime"],
    "inland marine": ["inland marine", "inland"],
    "commercial auto": ["commercial auto", "comm auto", "auto"],
}


def extract_lob_flags(policy_data: list[dict[str, Any]]) -> dict[int, dict[str, bool]]:
    """
    From POLICY sheet data, extract LOB flags per DS row.
    Returns: {row_index: {"property": True/False, "crime": True/False, ...}}
    """
    flags: dict[int, dict[str, bool]] = {}
    for idx, row in enumerate(policy_data):
        row_flags: dict[str, bool] = {}
        for lob_name in ["property", "crime", "inland marine", "commercial auto"]:
            key = _find_header_key(row, [lob_name])
            if key:
                val = str(row.get(key, "")).strip().lower()
                row_flags[lob_name] = val == "yes"
            else:
                row_flags[lob_name] = True  # default to Yes if column not found
        flags[idx] = row_flags
    return flags


def filter_rows_by_lob(
    rows: list[dict[str, Any]],
    sheet_name: str,
    lob_flags: dict[int, dict[str, bool]],
) -> list[dict[str, Any]]:
    """Remove rows from a sub-sheet where the corresponding LOB flag is No in Policy."""
    sn = sheet_name.lower().strip()
    target_lob = None
    for lob_name, keywords in LOB_SHEET_MAP.items():
        if any(kw in sn for kw in keywords):
            target_lob = lob_name
            break

    if target_lob is None:
        return rows  # Not a LOB sub-sheet

    filtered = []
    for idx, row in enumerate(rows):
        if idx in lob_flags:
            if lob_flags[idx].get(target_lob, True):
                filtered.append(row)
            # else: skip this row — LOB=No for this DS
        else:
            filtered.append(row)  # Keep if no flag info

    return filtered


# ---------------------------------------------------------------------------
# Sheet-specific prompt enhancements
# ---------------------------------------------------------------------------

def _build_sheet_specific_prompt(sheet_name: str) -> str:
    """Return additional IMS-specific prompt rules for each sheet type."""
    sn = sheet_name.lower().strip()
    lines: list[str] = []

    if "ims" in sn and "screen" in sn:
        lines.append("IMS SCREEN RULES:")
        lines.append("  - Company (Col K) must ALWAYS be 'Acme'")
        lines.append("  - Line (Col I) must ALWAYS be 'Package'")
        lines.append("  - Billing Type (Col L) must ALWAYS be 'Agency Bill'")
        lines.append("  - Legal Entity (Col B) must ALWAYS be 'Corporation'")
        lines.append("  - Underwriter (Col M): random from 'Khan, Nadeem', 'Khanra, Kushal', 'Singh, Ravendra'")
        lines.append(f"  - Producer by Location (Col G): random from: {', '.join(IMS_PRODUCERS[:10])}...")
        lines.append("  - Address (Col E): street-level ONLY — NO city, state, or ZIP")
        lines.append("  - Submitted (Col H): current date at generation time")

    elif "policy" in sn:
        lines.append("POLICY SHEET RULES:")
        lines.append("  - Company and Program are ALWAYS 'NetRate'")
        lines.append("  - General Liability (Col G) = ALWAYS Yes")
        lines.append("  - Excess Liability (Col H) = ALWAYS No")
        lines.append("  - Umbrella (Col I) = ALWAYS No")
        lines.append("  - Cols M and N must be COMPLETELY BLANK (empty string)")
        lines.append(f"  - Type (Col V): random from {POLICY_TYPE}")
        lines.append("  - Legal Entity (Col W) = ALWAYS 'Corporation'")
        lines.append("  - Property (Col R), Crime (Col S), Inland Marine (Col T) = random Yes/No")
        lines.append("  - Commercial Auto (Col J) = random Yes/No")
        lines.append("  - If Commercial Auto = No, Liability Limit and Deductible must be BLANK")
        lines.append("  - Liability Limit: random between 10,000 and 10,000,000")
        lines.append("  - Liability Deductible: random between 250 and 100,000")

    elif "property" in sn:
        lines.append("PROPERTY SHEET RULES (CRITICAL — MANY FIXED VALUES):")
        lines.append(f"  - Protection Class (Col C): random from {PROPERTY_PROTECTION_CLASS}")
        lines.append(f"  - Class Lookup (Col D): random from {[c[:30] for c in PROPERTY_CLASS_LOOKUP[:5]]}...")
        lines.append("  - Class Code (Col E): MUST match the code in Class Lookup (Col D)")
        lines.append(f"  - Risk Type (Col F): random from {PROPERTY_RISK_TYPE}")
        lines.append(f"  - Construction (Col I): random from {PROPERTY_CONSTRUCTION}")
        lines.append(f"  - Blanket Option (Col J): random from {PROPERTY_BLANKET_OPTION}")
        lines.append(f"  - Number of Stories (Col L): 1, 2, or 3 ONLY")
        lines.append("  - Building Exposure (Col T): random $ between $728,395 and $1,000,000")
        lines.append("  - Tenant Improvements (Col U): ALWAYS $0")
        lines.append("  - Total Exposure (Col V) = EXACTLY same as Building Exposure (Col T)")
        lines.append(f"  - Cause of Loss (Col W): {PROPERTY_CAUSE_OF_LOSS}")
        lines.append(f"  - Wind Hail Deductible (Col Z): {PROPERTY_WIND_HAIL} — use % symbol not decimal")
        lines.append(f"  - Inflation Guard (Col AB): {PROPERTY_INFLATION_GUARD} — use % symbol")
        lines.append("  - Building Ordinance (Col AF): ALWAYS 'No'")
        lines.append("  - Building Ordinance B (Col AG): ALWAYS 0")
        lines.append("  - Building Ordinance C (Col AH): ALWAYS 0")
        lines.append("  - Personal Property Sub-Description (Col AI): ALWAYS EMPTY")
        lines.append("  - Cols AK-AO (Personal Property sub-fields): ALL EMPTY")
        lines.append("  - Watchman Service (Col AZ): ALWAYS 'None'")
        lines.append("  - Alarm (Col BA): ALWAYS 'None'")
        lines.append("  - Income with Extra Expense (Col BF): ALWAYS 'No'")
        lines.append(f"  - Business Income Type (Col BG): random from {PROPERTY_BI_TYPE} with EQUAL distribution")
        lines.append("  - Option (Col BI): ALWAYS 'None'")
        lines.append("  - Days in Deductible (Col BK): ALWAYS 'N/A'")
        lines.append("  - Civil Authority (Col BL): ALWAYS 'N/A'")
        lines.append("  - Dependent Properties Limit (Col BP): ALWAYS 0")
        lines.append("  - Dependent Properties Form (Col BQ): ALWAYS 'N/A'")
        lines.append("  - Limit Percentage (Col BR): ALWAYS '100%'")
        lines.append("  - Extra Expense Limit (Col BS): ALWAYS '$0'")
        lines.append("  - Cols BT-BW: ALWAYS 'N/A'")
        lines.append("  - Cols BX-BZ (sub-limits): ALWAYS '$0'")
        lines.append("  - Cols CA-CE (Water/Power/Communication): ALWAYS 'No'")
        lines.append("  - Pollutant Cleanup (Col CF): ALWAYS '$0'")
        lines.append("  - Cols CH-CN: ALWAYS '$0'")
        lines.append("  - Roof Tank (Col CU): ALWAYS 'No'")
        lines.append("  - Building Construction/Description (CY/CZ): ALWAYS EMPTY")

    elif "crime" in sn:
        lines.append("CRIME SHEET RULES:")
        lines.append(f"  - Class Lookup: random from {[c[:30] for c in CRIME_CLASS_LOOKUP[:5]]}...")
        lines.append("  - Description (Col D) auto-matches Class Lookup (Col C)")
        lines.append(f"  - Coverage Is Written (Col F): MUST be from {CRIME_COVERAGE_IS_WRITTEN}")
        lines.append("  - NEVER use 'Yes', 'No', or 'Coindemnity' for Coverage Is Written")
        lines.append(f"  - All Deductible fields: from {CRIME_DEDUCTIBLE[:8]}...")
        lines.append("  - When Employee Theft=No: ALL child fields (H-R) must be BLANK")
        lines.append("  - When Robbery=No: child fields must be BLANK")

    elif "general liability" in sn or ("gl" in sn and "screen" not in sn):
        lines.append("GENERAL LIABILITY RULES:")
        lines.append(f"  - Policy Type (Col B): {GL_POLICY_TYPE}")
        lines.append(f"  - Each Occurrence (Col C): from {GL_EACH_OCCURRENCE[:5]}...")
        lines.append(f"  - Deductible (Col F): from {GL_DEDUCTIBLE[:8]}...")
        lines.append(f"  - Class (Col M): MUST include name AND code, e.g. 'Bakeries (10100)'")
        lines.append("  - All monetary values: use comma-formatted strings (1,000,000 not 1000000)")

    elif "commercial auto" in sn or ("comm" in sn and "auto" in sn):
        lines.append("COMMERCIAL AUTO RULES:")
        lines.append(f"  - Composite Basis (Col C): {AUTO_COMPOSITE_BASIS}")
        lines.append(f"  - Class Lookup (Col J): from {[c[:30] for c in AUTO_CLASS_LOOKUP[:5]]}...")
        lines.append("  - Description (Col L): MUST match Class Lookup value")
        lines.append(f"  - Vehicle Type (Col AQ): from {AUTO_VEHICLE_TYPE[:5]}...")
        lines.append(f"  - OTC Coverage: from {AUTO_OTC_COVERAGE[:4]}...")
        lines.append(f"  - OTC Deductible: from {AUTO_OTC_DEDUCTIBLE}")
        lines.append(f"  - Uninsured: from {AUTO_UNINSURED[:8]}...")
        lines.append("  - Rate Override (Col E): Yes/No only")
        lines.append("  - Tapes, Records and Discs: Yes/No only")

    elif "inland marine" in sn:
        lines.append("INLAND MARINE RULES:")
        lines.append("  - Class Lookup (Col C): COPY EXACTLY from PROPERTY sheet for same DS")
        lines.append("  - Class Code (Col D): COPY EXACTLY from PROPERTY sheet for same DS")
        lines.append("  - Construction (Col E): COPY EXACTLY from PROPERTY sheet for same DS")
        lines.append("  - When a parent checkbox = No, ALL child fields must be BLANK")
        lines.append(f"  - Alarm Type: from {IM_ALARM_TYPE}")
        lines.append(f"  - Alarm Quality: from {IM_ALARM_QUALITY}")
        lines.append(f"  - Level of Protection: from {IM_LEVEL_OF_PROTECTION}")
        lines.append(f"  - Cause of Loss: from {IM_CAUSE_OF_LOSS} (NOT 'Broad')")

    elif "optional" in sn and "coverage" in sn:
        lines.append("OPTIONAL COVERAGE RULES:")
        lines.append("  - Line of Business (Col B): ALWAYS 'Liability'")
        lines.append("  - Coverage (Col C): ALWAYS 'Stop Gap'")
        lines.append("  - Rate (Col F): ALWAYS '0.95'")
        lines.append("  - Data Set and Exposure fields MUST be populated (not empty)")

    if "location" in sn:
        lines.append("LOCATION SHEET RULES:")
        lines.append("  - Description: ALWAYS 'Loc #1'")
        lines.append("  - Address (Col C): full format [Street] [City], [State] [ZIP]")
        lines.append("  - Insured Location (Col D): ALWAYS 'Yes'")

    return "\n".join(lines)


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

    Special instructions are parsed into a structured ParsedInstructions object
    so that every configurable dimension (date range, DS numbering, state filter,
    address uniqueness, etc.) is extracted once and applied consistently —
    both in the LLM prompt and in a deterministic post-processing pass.
    """
    client = configure_openai()
    model_name = os.getenv("OPENAI_MODEL", "")

    # Parse instructions ONCE — used everywhere below
    pi = parse_special_instructions(special_instruction)

    # ---- cross-sheet context -----------------------------------------------
    previous_context = ""
    if previous_sheets_data:
        previous_context = (
            "\nCRITICAL — CROSS-SHEET DATA CONSISTENCY:\n"
            "The following sheets have already been generated. You MUST reuse the SAME values "
            "for matching fields across sheets. Row N in this sheet corresponds to Row N in all "
            "other sheets (same insured / policy / location).\n\n"
            "Fields that MUST be identical across sheets (when the column exists):\n"
            "  - Names (Insured Name, Contact Name, Agent Name, etc.)\n"
            "  - Addresses (Street, City, State, ZIP)\n"
            "  - Phone numbers, Email addresses\n"
            "  - Policy numbers, Account numbers\n"
            "  - Effective dates, Expiration dates, Submitted dates\n"
            "  - SSN, FEIN, Tax ID\n"
            "  - Data Set IDs (DS_01, DS_02, etc.) — must match row-for-row\n"
            "  - Class Lookup and Class Code — copy from PROPERTY sheet for Inland Marine\n\n"
            "Previously generated data:\n"
        )
        for prev_sheet, prev_data in previous_sheets_data.items():
            previous_context += f"\n--- Sheet: '{prev_sheet}' ---\n"
            previous_context += json.dumps(prev_data, indent=2)
            previous_context += "\n"

    # ---- assemble prompt ---------------------------------------------------
    user_instruction_block = pi.raw.strip() or "No additional instructions provided."
    configurable_constraints = _build_configurable_constraints(pi)
    sheet_rules = _build_sheet_specific_prompt(sheet_name)

    prompt = f"""You are a test data generator for US insurance applications (IMS / NetRate system).

Sheet: "{sheet_name}"

Generate exactly {row_count} rows of realistic test data for these columns:
{json.dumps(headers)}

DUPLICATE COLUMN HANDLING:
Columns with a "(Col X)" suffix are separate columns that happen to share a name in the original
spreadsheet. Generate DIFFERENT values for each. Use ALL column names (with suffixes) as JSON keys.

{previous_context}

INSTRUCTION PRIORITY (strictly followed in order):
1. USER SPECIAL INSTRUCTIONS — follow exactly:
{user_instruction_block}

2. SHEET-SPECIFIC IMS RULES (from the Notion rule book — MUST be followed):
{sheet_rules}

3. STRUCTURED CONSTRAINTS (derived automatically from the instructions above):
{configurable_constraints}

4. CROSS-SHEET CONSISTENCY — use values from previously generated sheets for matching rows.

5. DEFAULT RULES — apply only when not overridden by anything above:
  • Generate realistic US insurance test data.
  • Phone: (XXX) XXX-XXXX | SSN: XXX-XX-XXXX | ZIP: 5-digit or ZIP+4
  • State: 2-letter abbreviation (CA, TX, NY …)
  • Dates: MM/DD/YYYY
  • Policy numbers: realistic US insurance format
  • Dollar values: use comma-formatted strings with $ sign (e.g. $1,000,000 not $1000000)
  • Data Set IDs: {pi.ds_prefix}{pi.ds_start}, {pi.ds_prefix}{pi.ds_start + 1} … (zero-padded to 2 digits if < 100 rows)
  • Effective and Expiration dates must be IDENTICAL across all sheets for the same row.
  • When a field says "ALWAYS X" — use that exact value, no exceptions.
  • When a field says "EMPTY" — leave it as empty string "", not null, not $0, not N/A.
  • Dropdown/enum fields: ONLY use values from the specified list. NEVER invent new values.

Return ONLY a JSON object: {{"data": [{{col: val, ...}}, ...]}} — no markdown, no explanation.
"""

    # ---- call LLM with retry on rate-limit ---------------------------------
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
                            "of row objects. User special instructions override all defaults."
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

            # Post-process: deterministically enforce date rules derived from PI
            normalized = _enforce_effective_expiration_date_range(normalized, pi)

            # Post-process: sheet-specific IMS rule enforcement
            normalized = post_process_sheet(
                normalized, sheet_name, pi, previous_sheets_data
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