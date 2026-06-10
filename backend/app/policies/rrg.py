"""RRG Excel Rater policy handler.

Enforces all hard validation rules from the RRG rulebook:
- GL_Rating: Claims Made/Occurrence Selection valid values, non-zero schedule rating (Rule 15/17)
- Vehicle Type: double-hyphen format (Rule 26)
- Policy Basis cross-sheet consistency GL → PL/EBL (Rule 23/35)
- Conditional row counts for PL, Abuse, and Auto sheets
- Vehicles/Auto_Rating linked to correct Auto=Yes insureds (cross-sheet)
- EBL employee counts stored as integers
- Multiple locations/vehicles/CoB prompt guidance (Rules 36-38)
"""
from __future__ import annotations

import random
from datetime import datetime
from typing import Any


def _parse_date(val: str):
    """Parse MM/DD/YYYY or MMDDYYYY into a date object, or return None."""
    s = str(val or "").strip()
    for fmt in ("%m/%d/%Y", "%m%d%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _format_zip5(val) -> str | None:
    """Normalize a ZIP value to a strict 5-digit string (Rules 3 / 10 / 28).

    Strips non-digits, then either left-pads to 5 (preserving leading zeros
    for CT/NJ ZIPs like 06103, 07102) or trims to 5 if longer. Returns the
    cleaned string, or None when the input had no digits at all so callers
    can decide to fall back on the insured's ZIP.
    """
    digits = "".join(ch for ch in str(val or "") if ch.isdigit())
    if not digits:
        return None
    if len(digits) < 5:
        return digits.zfill(5)
    return digits[:5]


def _format_us_phone(val) -> str:
    """Format a contact/phone value as a USA 10-digit number "XXX XXX XXXX".

    Strips every non-digit character, then groups the digits 3-3-4 separated by
    single spaces (e.g. 2125550147 → "212 555 0147") per Rule 1 / Rule 9
    [DS_044 / DS_045]. If the cleaned value is not exactly 10 digits a valid
    10-digit number is fabricated (area code not starting with 0 or 1) so the
    output always satisfies the format constraint.
    """
    digits = "".join(ch for ch in str(val or "") if ch.isdigit())
    if len(digits) != 10:
        digits = str(random.randint(2, 9)) + "".join(
            str(random.randint(0, 9)) for _ in range(9)
        )
    return f"{digits[:3]} {digits[3:6]} {digits[6:]}"

VALID_STATES = ["NY", "AL", "IL", "TX", "FL", "PA", "CT", "NJ"]

VALID_ORG_TYPES = [
    "Corporation", "Individual", "Joint Venture", "Limited Partnership",
    "LLC", "Other", "Partnership", "Trust",
]

VALID_CLASS_CODES = [
    "10150", "10205", "10257", "10331", "10332", "11039", "11128", "12361",
    "13506", "13507", "16881", "16901", "18438", "22204", "40005", "40006",
    "40117", "41421", "41422", "41667", "41668", "41670", "41677", "41678",
    "41715", "41716", "44277", "44427", "44432", "44433", "44434", "44438",
    "44439", "44440", "44500", "46427", "46607", "46622", "46671", "47367",
    "47469", "47474", "47475", "47476", "47477", "47478", "48557", "48558",
    "48600", "48925", "49870", "49890", "49891", "60010", "60011", "61000",
    "61212", "61216", "61217", "61218", "61225", "61226", "61227", "63010",
    "63011", "63013", "63216", "63218", "66309", "67017", "67509", "67510",
    "67511", "67513", "68606", "68607", "68703", "68706", "68707", "96410",
    "96816", "97047", "98598", "98751",
]

# Rule 15: valid GL Claims Made/Occurrence Selection values
VALID_GL_OCCURRENCE = ["Occurrence", "1", "2", "3", "4", "5", "More than 5"]
VALID_GL_CLAIMS_MADE_RETRO = ["1", "2", "3", "4", "5", "More than 5"]

VALID_MED_PAY_LIMITS = ["5000", "4000", "3000", "2000", "1000", "Excluded"]
VALID_NEW_RENEWAL = ["New Business", "Renewal"]
VALID_PL_POLICY_BASIS = ["Occurrence", "Claims Made"]

VALID_COB = [
    "Facility", "Agency", "Doctoral/Specialized Degree", "Graduate/Certified Degree",
    "Counselor/Social Worker", "Teacher/Nutritionist", "Other Degrees/Certifications",
    "Adoptions/Foster Care",
]

VALID_HAZARD_TIER = ["Moderate", "High"]
VALID_ABUSE_HAZARD = ["Lower Hazard", "Higher Hazard"]

# Rule 26: vehicle types use double-hyphen (--)
VALID_VEHICLE_TYPES = [
    "School/Church Bus -- Seats 1-8",
    "School/Church Bus -- Seats 9-20",
    "School/Church Bus -- Seats 21-60",
    "School/Church Bus -- Seats 61+",
    "Private Passenger",
    "Light Truck -- Service",
    "Light Truck -- Commercial",
    "Medium Truck -- Service",
    "Medium Truck -- Commercial",
    "Trailer",
]

VALID_RADIUS = ["Local", "Intermediate", "Long Haul"]
VALID_VEHICLE_USE = ["Service", "Retail", "Passenger", "Contractor"]

# Reference/lookup sheets — no LLM generation needed
_REFERENCE_SHEETS = {
    "gl_classtable", "pl_table", "auto_table", "ebl_table", "system_defaults",
}

# Valid cities per allowed state — injected into prompts to prevent city/state mismatches
# (e.g., "Boston, NY" is invalid; Boston is in MA which is not an allowed state)
VALID_CITIES_BY_STATE: dict[str, list[str]] = {
    "NY": ["New York City", "Buffalo", "Albany", "Syracuse", "Rochester", "Yonkers"],
    "AL": ["Birmingham", "Montgomery", "Huntsville", "Mobile", "Tuscaloosa"],
    "IL": ["Chicago", "Springfield", "Rockford", "Naperville", "Peoria", "Aurora"],
    "TX": ["Houston", "Dallas", "Austin", "San Antonio", "Fort Worth", "El Paso"],
    "FL": ["Tampa", "Miami", "Orlando", "Jacksonville", "Fort Lauderdale", "St. Petersburg"],
    "PA": ["Philadelphia", "Pittsburgh", "Harrisburg", "Allentown", "Erie"],
    "CT": ["Hartford", "New Haven", "Stamford", "Bridgeport", "Waterbury"],
    "NJ": ["Newark", "Jersey City", "Trenton", "Camden", "Paterson", "Elizabeth"],
}

# Real street addresses per city — injected into prompts so the LLM uses
VALID_ADDRESSES_BY_CITY: dict[str, list[tuple[str, str]]] = {
    # (street, ZIP)
    "New York City": [("125 Broad St", "10004"), ("350 Fifth Ave", "10118"), ("30 Rockefeller Plaza", "10112"), ("1 Penn Plaza", "10119")],
    "Buffalo":       [("95 Perry St", "14203"), ("One Seneca Tower", "14203"), ("257 Franklin St", "14202"), ("12 Fountain Plaza", "14202")],
    "Albany":        [("1 Commerce Plaza", "12260"), ("60 State St", "12207"), ("677 Broadway", "12207")],
    "Syracuse":      [("100 S Clinton St", "13202"), ("300 S State St", "13202"), ("120 Madison St", "13202")],
    "Rochester":     [("100 Chestnut St", "14604"), ("1 E Main St", "14614"), ("45 East Ave", "14604")],
    "Yonkers":       [("87 Nepperhan Ave", "10701"), ("40 S Broadway", "10701"), ("20 S Broadway", "10701")],
    "Birmingham":    [("2100 3rd Ave N", "35203"), ("420 20th St N", "35203"), ("505 20th St N", "35203"), ("1901 6th Ave N", "35203")],
    "Montgomery":    [("1 Dexter Ave", "36104"), ("445 Dexter Ave", "36104"), ("100 N Union St", "36104")],
    "Huntsville":    [("200 Sparkman Dr NW", "35805"), ("320 Pelham Ave SW", "35801"), ("700 Monroe St", "35801")],
    "Mobile":        [("150 Government St", "36602"), ("205 Government St", "36602"), ("251 St Joseph St", "36602")],
    "Tuscaloosa":    [("2200 University Blvd", "35401"), ("500 Greensboro Ave", "35401"), ("1800 McFarland Blvd E", "35404")],
    "Chicago":       [("233 S Wacker Dr", "60606"), ("111 E Wacker Dr", "60601"), ("200 E Randolph St", "60601"), ("55 E Monroe St", "60603")],
    "Springfield":   [("300 E Monroe St", "62701"), ("1 Old State Capitol Plaza", "62701"), ("524 S 2nd St", "62701")],
    "Rockford":      [("401 S Main St", "61101"), ("321 W State St", "61101"), ("100 N 1st St", "61107")],
    "Naperville":    [("400 S Eagle St", "60540"), ("55 S Main St", "60540"), ("200 E 5th Ave", "60563")],
    "Peoria":        [("100 NE Monroe St", "61602"), ("401 Main St", "61602"), ("456 Fulton St", "61602")],
    "Aurora":        [("44 E Downer Pl", "60505"), ("20 E Downer Pl", "60505"), ("1 E Benton St", "60505")],
    "Houston":       [("1000 Main St", "77002"), ("600 Travis St", "77002"), ("910 Louisiana St", "77002"), ("700 Milam St", "77002")],
    "Dallas":        [("1401 Elm St", "75202"), ("2200 Ross Ave", "75201"), ("500 N Akard St", "75201"), ("1717 Main St", "75201")],
    "Austin":        [("300 E 6th St", "78701"), ("100 Congress Ave", "78701"), ("823 Congress Ave", "78701")],
    "San Antonio":   [("300 E Commerce St", "78205"), ("100 W Houston St", "78205"), ("200 E Market St", "78205")],
    "Fort Worth":    [("500 W 7th St", "76102"), ("301 Commerce St", "76102"), ("801 Cherry St", "76102")],
    "El Paso":       [("500 E San Antonio Ave", "79901"), ("300 N Campbell St", "79901"), ("125 W Mills Ave", "79901")],
    "Tampa":         [("100 N Tampa St", "33602"), ("200 N Tampa St", "33602"), ("400 N Tampa St", "33602")],
    "Miami":         [("100 SE 2nd St", "33131"), ("200 S Biscayne Blvd", "33131"), ("1 SE 3rd Ave", "33131")],
    "Orlando":       [("400 S Orange Ave", "32801"), ("201 S Orange Ave", "32801"), ("100 E Pine St", "32801")],
    "Jacksonville":  [("117 W Duval St", "32202"), ("501 W Church St", "32202"), ("225 E Coastline Dr", "32202")],
    "Fort Lauderdale": [("100 SE 3rd Ave", "33301"), ("200 E Broward Blvd", "33301"), ("1 E Broward Blvd", "33301")],
    "St. Petersburg": [("175 5th St N", "33701"), ("100 2nd Ave N", "33701"), ("400 1st Ave S", "33701")],
    "Philadelphia":  [("1500 Market St", "19102"), ("1700 Market St", "19103"), ("1818 Market St", "19103"), ("200 S Broad St", "19102")],
    "Pittsburgh":    [("600 Grant St", "15219"), ("525 William Penn Pl", "15219"), ("1 PPG Pl", "15222")],
    "Harrisburg":    [("212 Locust St", "17101"), ("333 Market St", "17101"), ("200 N 3rd St", "17101")],
    "Allentown":     [("435 Hamilton St", "18101"), ("702 Hamilton St", "18101"), ("100 N 7th St", "18101")],
    "Erie":          [("100 State St", "16507"), ("208 E Bayfront Pkwy", "16507"), ("626 State St", "16501")],
    "Hartford":      [("185 Asylum St", "06103"), ("100 Pearl St", "06103"), ("280 Trumbull St", "06103")],
    "New Haven":     [("195 Church St", "06510"), ("900 Chapel St", "06510"), ("265 Church St", "06510")],
    "Stamford":      [("1 Landmark Sq", "06901"), ("300 Main St", "06901"), ("680 Washington Blvd", "06901")],
    "Bridgeport":    [("999 Broad St", "06604"), ("300 Main St", "06604"), ("1000 Lafayette Blvd", "06604")],
    "Waterbury":     [("235 Grand St", "06702"), ("55 W Main St", "06702"), ("160 Bank St", "06702")],
    "Newark":        [("1 Gateway Center", "07102"), ("744 Broad St", "07102"), ("550 Broad St", "07102"), ("80 Park Plaza", "07102")],
    "Jersey City":   [("30 Hudson St", "07302"), ("101 Hudson St", "07302"), ("2 Exchange Pl", "07302")],
    "Trenton":       [("1 John Fitch Way", "08611"), ("225 W State St", "08608"), ("33 W State St", "08608")],
    "Camden":        [("200 Federal St", "08103"), ("1 Port Center", "08103"), ("1 Riverside Dr", "08103")],
    "Paterson":      [("155 Market St", "07505"), ("100 Hamilton Plaza", "07505"), ("125 Ellison St", "07505")],
    "Elizabeth":     [("50 Winfield Scott Plaza", "07201"), ("100 First St", "07201"), ("289 N Broad St", "07208")],
}

def _build_address_reference() -> str:
    """Build a compact address reference string for prompt injection."""
    lines = []
    for state, cities in VALID_CITIES_BY_STATE.items():
        city_lines = []
        for city in cities:
            addrs = VALID_ADDRESSES_BY_CITY.get(city, [])
            if addrs:
                examples = "; ".join(f"{a[0]}, {a[1]}" for a in addrs[:2])
                city_lines.append(f"    {city}: {examples}")
        lines.append(f"  {state}:\n" + "\n".join(city_lines))
    return "\n".join(lines)

_ADDRESS_REFERENCE = _build_address_reference()

# One-line string listing State: City, City... for prompt injection
_CITIES_REFERENCE = "  " + "\n  ".join(
    f"{state}: {', '.join(cities)}" for state, cities in VALID_CITIES_BY_STATE.items()
)


# Full allowed ranges per the RRG rulebook. The "checkpoints" are the
# representative values a test set must cover so boundary/maximum scenarios are
# exercised at high volume (DS_047 / DS_048 / DS_049). Small values dominate via
# the fill pattern; checkpoints guarantee the spread reaches the documented max.
LOCATION_COUNT_CHECKPOINTS = [1, 2, 3, 5, 8, 12, 16, 20]  # Rule 37: 1–20
VEHICLE_COUNT_CHECKPOINTS = [1, 2, 3, 5, 8, 12, 16, 20]   # Rule 36: 1–20 (Auto=Yes)
COB_COUNT_CHECKPOINTS = [1, 2, 3, 4, 5, 6, 7, 8]          # Rule 38: 1–8


def _coverage_counts(
    n: int, checkpoints: list[int], fill_pattern: list[int]
) -> list[int]:
    """Per-insured counts that guarantee full-range coverage at high volume.

    Most insureds get the small, repeating ``fill_pattern`` values (keeping the
    generated row volume sane). Once there are at least as many insureds as
    checkpoints, the checkpoint values are overlaid at evenly spread positions so
    the dataset provably contains the boundary and maximum counts (e.g. a 20-
    location insured, an 8-CoB insured). Below that threshold the original light
    cycling behavior is preserved, so small runs are unchanged.
    """
    if n <= 0:
        return []
    counts = [fill_pattern[i % len(fill_pattern)] for i in range(n)]
    cps = sorted(set(checkpoints))
    if n >= len(cps):
        for j, cp in enumerate(cps):
            pos = j * (n - 1) // (len(cps) - 1) if len(cps) > 1 else 0
            counts[pos] = cp
    return counts


def _vehicle_counts_for_insureds(auto_insureds: list) -> list[int]:
    """Per-insured vehicle counts spanning the full 1–20 range (DS_033 / DS_048)."""
    return _coverage_counts(len(auto_insureds), VEHICLE_COUNT_CHECKPOINTS, [2, 1, 3])


def _location_counts_for_insureds(policy_data: list) -> list[int]:
    """Per-insured location counts spanning the full 1–20 range (DS_035 / DS_047)."""
    return _coverage_counts(len(policy_data), LOCATION_COUNT_CHECKPOINTS, [2, 3, 1])


def _cob_counts_for_insureds(pl_insureds: list) -> list[int]:
    """Per-insured Class of Business counts spanning the full 1–8 range (DS_038 / DS_049).

    This ensures some insureds have many CoB rows (up to the rulebook max of 8,
    Rule 38) while others have a single entry, creating realistic variability and
    full boundary coverage in PL Rating data.
    """
    return _coverage_counts(len(pl_insureds), COB_COUNT_CHECKPOINTS, [2, 1, 2])


def _unit_suffix(v: int) -> str:
    """Spreadsheet-style suffix for a vehicle unit index (0→A, 25→Z, 26→AA).

    Supports vehicle counts well beyond the 5-letter range the prompt previously
    assumed, so an insured with up to 20 vehicles (Rule 36 / DS_048) gets unique
    unit labels without an index error.
    """
    letters = ""
    v += 1
    while v > 0:
        v, rem = divmod(v - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _build_insured_row_map(counts: list[int]) -> list[int]:
    """Map flat row index → insured index given per-insured counts."""
    mapping = []
    for insured_idx, count in enumerate(counts):
        mapping.extend([insured_idx] * count)
    return mapping


def _find_col(row: dict, *keywords: str) -> str | None:
    """Return first key whose lowercase form contains ALL keyword substrings."""
    for key in row:
        kl = key.lower()
        if all(k.lower() in kl for k in keywords):
            return key
    return None


def _normalize_vehicle_type(value: str) -> str:
    """Normalize em/en dashes to double-hyphen; snap to exact allowed value."""
    if not isinstance(value, str):
        return value
    normalized = value.replace("–", "--").replace("—", "--").replace(" - ", " -- ")
    for valid in VALID_VEHICLE_TYPES:
        if normalized.strip().lower() == valid.lower():
            return valid
    def _core(s: str) -> str:
        return s.lower().replace("--", "").replace("-", "").replace("  ", " ").strip()
    for valid in VALID_VEHICLE_TYPES:
        if _core(normalized) == _core(valid):
            return valid
    return normalized


def _count_lob_yes_rows(policy_data: list[dict], lob_fragment: str) -> int:
    """Count policy rows where the LOB (Yes/No) column equals Yes."""
    count = 0
    for row in policy_data:
        for key, val in row.items():
            kl = key.lower()
            if lob_fragment.lower() in kl and "(yes/no)" in kl:
                if str(val).strip().lower() == "yes":
                    count += 1
                break
    return count


def _get_lob_yes_insureds(policy_data: list[dict], lob_fragment: str) -> list[dict]:
    """Return policy rows where the given LOB (Yes/No) column equals Yes."""
    result = []
    for row in policy_data:
        for key, val in row.items():
            kl = key.lower()
            if lob_fragment.lower() in kl and "(yes/no)" in kl:
                if str(val).strip().lower() == "yes":
                    result.append(row)
                break
    return result


def _insured_summary(insureds: list[dict]) -> str:
    """Build a concise list of insured names + states for injection into prompts."""
    lines = []
    for i, row in enumerate(insureds, 1):
        name  = row.get("Named Insured", "Unknown")
        state = row.get("Rating State", row.get("State", "?"))
        city  = row.get("City", "?")
        zipcode = row.get("ZIP Code", "?")
        lines.append(f"  {i}. {name} — State: {state}, City: {city}, ZIP: {zipcode}")
    return "\n".join(lines)


class RrgHandler:
    policy_type = "RRG"

    # ------------------------------------------------------------------
    # Sheet type detection
    # ------------------------------------------------------------------

    def detect_sheet_type(self, sheet_name: str) -> str:
        sn = sheet_name.lower().strip()
        if "policy information" in sn:
            # Return "policy" so main.py sets policy_data and conditional
            # sheets (PL, Abuse, Vehicles, Auto) get correct row counts
            return "policy"
        if "sched of location" in sn:
            return "locations"
        if "gl_rating" in sn or sn == "gl rating":
            return "gl_rating"
        if "pl_rating" in sn or sn == "pl rating":
            return "pl_rating"
        if "ebl_rating" in sn or sn == "ebl rating":
            return "ebl_rating"
        if "abuse_rating" in sn or sn == "abuse rating":
            return "abuse_rating"
        if "sched of vehicle" in sn:
            return "vehicles"
        if "auto_rating" in sn or sn == "auto rating":
            return "auto_rating"
        # "Test Scenerio Details" (note the client's misspelling of "Scenario")
        if "scenario" in sn or "scenerio" in sn:
            return "scenario_details"
        sn_no_space = sn.replace(" ", "_").replace("-", "_")
        if any(ref in sn_no_space for ref in _REFERENCE_SHEETS):
            return "reference"
        return "unknown"

    # ------------------------------------------------------------------
    # Sheet context (row count + rules injected into LLM prompt)
    # ------------------------------------------------------------------

    def build_sheet_context(
        self,
        sheet_name: str,
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        original_row_count: int,
        special_instruction: str = "",
    ) -> tuple[int, str]:
        sheet_type = self.detect_sheet_type(sheet_name)

        if sheet_type == "reference":
            return 0, ""

        # Test Scenerio Details is a deterministic summary built in pre_generate
        # from the already-generated sheets — never sent to the LLM.
        if sheet_type == "scenario_details":
            return 0, ""

        codes_sample = ", ".join(VALID_CLASS_CODES[:30]) + ", ..."

        # ---- Policy Information ----------------------------------------
        if sheet_type == "policy":
            n = original_row_count

            # --- LOB Yes/No distribution guidance (DS_042 / DS_033) ---
            # Ensure a realistic mix of Yes and No for each optional LOB.
            # GL is always Yes. The others must have BOTH Yes AND No values.
            def _lob_guidance(lob: str, yes_min: int, yes_max: int) -> str:
                no_min = max(1, n - yes_max)
                return (
                    f"- {lob} (Yes/No): assign 'Yes' to {yes_min}–{yes_max} insureds "
                    f"and 'No' to the rest ({no_min}–{n - yes_min}). "
                    f"NEVER set all rows to 'Yes' — there MUST be at least 1 'No'. [DS_042]"
                )

            if n >= 5:
                pl_guidance  = _lob_guidance("PL",    n - 2, n - 1)  # most Yes, 1-2 No
                ebl_guidance = _lob_guidance("EBL",   n - 2, n - 1)
                auto_guidance = (
                    f"- Auto (Yes/No): MUST assign 'Yes' to 2–{min(3, n - 1)} different insureds "
                    f"and 'No' to the rest. NEVER assign Auto=Yes to all or only 1. [DS_033]"
                )
                abuse_guidance = _lob_guidance("Abuse", 2, min(3, n - 1))
            elif n >= 3:
                pl_guidance  = _lob_guidance("PL",    n - 1, n - 1)  # 1 No
                ebl_guidance = _lob_guidance("EBL",   n - 1, n - 1)
                auto_guidance = _lob_guidance("Auto",  1, 2)
                abuse_guidance = _lob_guidance("Abuse", 1, 2)
            else:
                # 2 rows or fewer — just ensure at least 1 Yes
                pl_guidance   = "- PL (Yes/No): at least 1 Yes and try to include 1 No if possible"
                ebl_guidance  = "- EBL (Yes/No): at least 1 Yes and try to include 1 No if possible"
                auto_guidance = "- Auto (Yes/No): at least 1 Yes"
                abuse_guidance = "- Abuse (Yes/No): at least 1 Yes"

            lob_block = "\n".join([pl_guidance, ebl_guidance, auto_guidance, abuse_guidance])

            rules = f"""
RRG POLICY INFORMATION RULES (HARD CONSTRAINTS):
- Test ID: MUST follow the format DS-01, DS-02, DS-03, ... (sequential, zero-padded).
  NEVER use prefixes like PI-, GL-, SL-, or any other format. Always "DS-XX".
- Contact Number: USA 10-digit phone in "XXX XXX XXXX" format — three groups of 3-3-4 separated by single spaces (e.g. 212 555 0147). No dashes, parentheses, or other special chars [Rule 1 / DS_044]
- State: MUST be one of {VALID_STATES} [Rule 2]
- ZIP Code: 5-digit numeric only [Rule 3]
- Org Type / Entity: MUST be one of {VALID_ORG_TYPES} [Rule 4]
- State of Operation: MUST be one of {VALID_STATES} [Rule 5]
- Rating State: MUST be one of {VALID_STATES} [Rule 6]
- New / Renewal: MUST be "New Business" or "Renewal" [Rule 7]
- GL (Yes/No): ALWAYS "Yes" — GL is mandatory for all insureds [Rule 8]
- Producer Phone: USA 10-digit phone in "XXX XXX XXXX" format — three groups of 3-3-4 separated by single spaces (e.g. 646 555 0198) [Rule 9 / DS_045]
- Effective Date: MM/DD/YYYY format (e.g. 05/28/2026) — MUST VARY across rows, not all the same
- Expiration Date: MM/DD/YYYY format, exactly 1 year after effective date — MUST match effective date's year+1
- Country: Always "USA"
- Address: City MUST be a real city inside the chosen State — never use a city from another state.
  Street address MUST be a real street that exists in the chosen CITY (not another city in the same state).
  CRITICAL: Do NOT mix streets between cities. "1401 Lavaca St" is Austin, TX — NEVER put it in Dallas, TX.
  Use ONLY addresses from this reference (street, ZIP per city):
{_ADDRESS_REFERENCE}
LOB DISTRIBUTION — each optional LOB MUST have a realistic mix of Yes and No:
{lob_block}
"""
            return original_row_count, rules

        # ---- Schedule of Locations -------------------------------------
        if sheet_type == "locations":
            # Build per-insured location guidance when policy_data is available.
            # Each insured gets 2 locations to satisfy Rule 37 (multiple locations supported).
            insured_block = ""
            loc_row_count = original_row_count
            if policy_data:
                counts = _location_counts_for_insureds(policy_data)
                loc_row_count = sum(counts)
                lines = []
                row_num = 1
                for i, (insured_row, count) in enumerate(zip(policy_data, counts), 1):
                    name    = insured_row.get("Named Insured", "Unknown")
                    addr    = insured_row.get("Address Line 1", "")
                    city    = insured_row.get("City", "")
                    state   = insured_row.get("State", insured_row.get("Rating State", ""))
                    zipcode = insured_row.get("ZIP Code", "")
                    loc_lines = [
                        f"    Row {row_num} (Primary) — MUST use: {addr}, {city}, {state} {zipcode}"
                    ]
                    # For secondary locations, suggest a different city in the same state
                    alt_cities = [c for c in VALID_CITIES_BY_STATE.get(state, []) if c != city]
                    for extra in range(1, count):
                        alt_city = alt_cities[(extra - 1) % len(alt_cities)] if alt_cities else city
                        alt_addrs = VALID_ADDRESSES_BY_CITY.get(alt_city, [])
                        if alt_addrs:
                            alt_addr, alt_zip = alt_addrs[0]
                            loc_lines.append(
                                f"    Row {row_num + extra} (Secondary {extra}) — use: {alt_addr}, {alt_city}, {state} {alt_zip}"
                            )
                        else:
                            loc_lines.append(
                                f"    Row {row_num + extra} (Secondary {extra}) — different REAL address in {alt_city}, {state}"
                            )
                    lines.append(f"  Insured {i}: {name} — {count} location(s)\n" + "\n".join(loc_lines))
                    row_num += count
                insured_block = (
                    f"\nINSURED LIST — variable location counts ({loc_row_count} rows total):\n"
                    + "\n".join(lines)
                    + "\n\nPrimary location MUST exactly match the Policy Information address."
                    "\nSecondary/tertiary locations MUST be real alternate addresses in the same state."
                )

            rules = f"""
RRG SCHEDULE OF LOCATIONS RULES (HARD CONSTRAINTS):
- Test ID: MUST use the same DS-01, DS-02, ... IDs from Policy Information. NEVER invent new IDs like SL-001.
- "#": per-insured location number that RESTARTS at 1 for each insured's first location (1, 2, 3 within each Test ID) — NOT a global running count across all insureds [DS_043]
- State: MUST be one of {VALID_STATES} [Rule 11]
- ZIP: 5-digit numeric only [Rule 10]
- Class Code: MUST be from approved list: {codes_sample} [Rule 13]
- City, State, ZIP must be geographically consistent [Rule 12]
- Exposure Amount: realistic dollar value (e.g. $250000)
- Location counts VARY per insured across the full 1–20 range — follow the exact per-insured counts below [Rule 37]
- Location Name: "Primary Office" / "Main Clinic" for row 1; "Branch Office" / "Secondary Location" for extras
- City MUST be a real city inside the chosen State.
- Street address MUST be a real street that exists in the chosen CITY — NEVER use a street from another city.
  CRITICAL: "1401 Lavaca St" is in Austin, TX — do NOT use it for Dallas or Houston.
  "N Lamar Blvd" is in Austin, TX — do NOT use it for Houston.
  Use ONLY addresses from this reference (street, ZIP per city):
{_ADDRESS_REFERENCE}
{insured_block}
"""
            return loc_row_count, rules

        # ---- GL Rating -------------------------------------------------
        if sheet_type == "gl_rating":
            rules = f"""
RRG GL RATING RULES (HARD CONSTRAINTS):
- Test ID: MUST use the same DS-01, DS-02, ... IDs from Policy Information. NEVER invent new IDs like GL-001.
- Claims Made / Occurrence Selection: MUST be EXACTLY one of {VALID_GL_OCCURRENCE} [Rule 15]
  * Use "Occurrence" for occurrence-based policies
  * Use "1","2","3","4","5","More than 5" for claims-made (retroactive years back)
  * NEVER use "Claims Made" as a value — it is not in the allowed list
- Medical Payments Limit Selected: MUST be one of {VALID_MED_PAY_LIMITS} [Rule 16]
- Schedule Rating Credit/Debit (%): [Rule 17]
  * Credits MUST be negative (e.g. -5%, -10%, -15%) — NEVER 0% or positive for credits
  * Debits MUST be positive (e.g. 5%, 10%, 12%) — NEVER 0% or negative for debits
  * Do NOT use 0% — every row must be either a credit (negative) or debit (positive)
"""
            return original_row_count, rules

        # ---- PL Rating -------------------------------------------------
        if sheet_type == "pl_rating":
            # Conditional: only generate rows for PL=Yes insureds
            # DS_038: some insureds get multiple Class of Business rows (Rule 38)
            if policy_data:
                pl_insureds = _get_lob_yes_insureds(policy_data, "pl")
                if not pl_insureds:
                    return 0, ""
                cob_counts = _cob_counts_for_insureds(pl_insureds)
                row_count = sum(cob_counts)
                lines = []
                row_num = 1
                cob_list = list(VALID_COB)  # for example guidance
                for i, (insured_row, count) in enumerate(zip(pl_insureds, cob_counts)):
                    name  = insured_row.get("Named Insured", "Unknown")
                    state = insured_row.get("Rating State", insured_row.get("State", "?"))
                    tid   = insured_row.get("Test ID", f"DS-{i+1:02d}")
                    row_lines = []
                    for c in range(count):
                        cob_hint = cob_list[(i + c) % len(cob_list)]
                        row_lines.append(
                            f"    Row {row_num + c}: Test ID = {tid}, CoB = \"{cob_hint}\" "
                            f"(different CoB per row)"
                        )
                    lines.append(
                        f"  Insured {i+1}: {name} ({state}) — {count} CoB row(s)\n"
                        + "\n".join(row_lines)
                    )
                    row_num += count
                insured_block = (
                    f"\nGenerate variable Class of Business rows per PL=Yes insured "
                    f"({row_count} rows total). [Rule 38 / DS_038]\n"
                    "Each row for the SAME insured MUST use a DIFFERENT Class of Business value.\n"
                    "All rows for the same insured share the same Test ID, Policy Basis, and Retroactive Date.\n"
                    + "\n".join(lines)
                )
            else:
                row_count = original_row_count
                insured_block = ""

            rules = f"""
RRG PL RATING RULES (HARD CONSTRAINTS):
- ONLY generate rows for insureds where PL (Yes/No) = "Yes" in Policy Information [Rule 8]
- Policy Basis: MUST be one of {VALID_PL_POLICY_BASIS} [Rule 18]
  * If GL "Claims Made / Occurrence Selection" = "Occurrence" → Policy Basis = "Occurrence"
  * If GL selection is numeric (1-5, More than 5) → Policy Basis = "Claims Made" [Rule 23/35]
- Retroactive Date: MM/DD/YYYY format; MUST be <= effective date [Rule 19]
- Class of Business: MUST be one of {VALID_COB} [Rule 19b]
- Count: numeric integer only [Rule 20]
- Hazard Tier: MUST be "Moderate" or "High" [Rule 21]
- An insured MUST have MULTIPLE Class of Business rows — each in a separate row with a different CoB value [Rule 38]
  * Class of Business counts VARY per insured across the full 1–8 range — follow the exact per-insured counts below
  * Each row for the same insured uses the SAME Test ID but a DIFFERENT Class of Business
{insured_block}
"""
            return row_count, rules

        # ---- EBL Rating ------------------------------------------------
        if sheet_type == "ebl_rating":
            # Conditional: only generate rows for EBL=Yes insureds
            if policy_data:
                ebl_insureds = _get_lob_yes_insureds(policy_data, "ebl")
                if not ebl_insureds:
                    return 0, ""
                row_count = len(ebl_insureds)
                insured_block = (
                    "\nGenerate one row per EBL=Yes insured (in this order):\n"
                    + _insured_summary(ebl_insureds)
                )
            else:
                row_count = original_row_count
                insured_block = ""

            rules = f"""
RRG EBL RATING RULES (HARD CONSTRAINTS):
- ONLY generate rows for insureds where EBL (Yes/No) = "Yes" in Policy Information
- Total Employee Count: integer numeric only — NO strings, NO quotes [Rule 22]
- Policy Basis: MUST be one of {VALID_PL_POLICY_BASIS} [Rule 18]
  * Must match GL selection: "Occurrence" if GL = "Occurrence", else "Claims Made" [Rule 23/35]
- Retroactive Date: MM/DD/YYYY format; MUST be <= effective date [Rule 24]
{insured_block}
"""
            return row_count, rules

        # ---- Abuse Rating ----------------------------------------------
        if sheet_type == "abuse_rating":
            if policy_data:
                abuse_insureds = _get_lob_yes_insureds(policy_data, "abuse")
                if not abuse_insureds:
                    return 0, ""
                row_count = len(abuse_insureds)
                insured_block = (
                    "\nGenerate one row per Abuse=Yes insured (in this order):\n"
                    + _insured_summary(abuse_insureds)
                )
            else:
                row_count = original_row_count
                insured_block = ""

            rules = f"""
RRG ABUSE RATING RULES (HARD CONSTRAINTS):
- Abuse Hazard Tier: MUST be "Lower Hazard" or "Higher Hazard" [Rule 25]
- ONLY generate rows for insureds where Abuse (Yes/No) = "Yes" in Policy Information
{insured_block}
"""
            return row_count, rules

        # ---- Schedule of Vehicles --------------------------------------
        if sheet_type == "vehicles":
            if policy_data:
                auto_insureds = _get_lob_yes_insureds(policy_data, "auto")
                if not auto_insureds:
                    return 0, ""
                # Variable vehicle counts per insured for scenario variability (DS_033)
                veh_counts = _vehicle_counts_for_insureds(auto_insureds)
                row_count  = sum(veh_counts)
                lines = []
                row_num = 1
                for i, (ins, count) in enumerate(zip(auto_insureds, veh_counts), 1):
                    name    = ins.get("Named Insured", "Unknown")
                    state   = ins.get("Rating State", ins.get("State", "?"))
                    zipcode = ins.get("ZIP Code", "?")
                    veh_lines = []
                    for v in range(count):
                        unit = f"{i:02d}{_unit_suffix(v)}"
                        veh_lines.append(f"    Row {row_num + v}: Unit #{unit} — vehicle {v+1} of {count}")
                    lines.append(
                        f"  Insured {i}: {name} — {count} vehicle(s), State: {state}, ZIP: {zipcode}\n"
                        + "\n".join(veh_lines)
                    )
                    row_num += count
                insured_block = (
                    f"\nGenerate variable vehicle counts per Auto=Yes insured "
                    f"({row_count} rows total). Garaging State and ZIP MUST match insured's state/ZIP:\n"
                    + "\n".join(lines)
                )
            else:
                row_count = original_row_count
                insured_block = ""

            vehicle_list = "\n  ".join(f"* {v}" for v in VALID_VEHICLE_TYPES)
            rules = f"""
RRG SCHEDULE OF VEHICLES RULES (HARD CONSTRAINTS):
- ONLY generate rows for insureds where Auto (Yes/No) = "Yes" in Policy Information [Rule 36]
- EVERY Auto=Yes insured MUST have at least 1 vehicle and at most 20 — NEVER leave an Auto=Yes insured with zero vehicle rows [Rule 36 / DS_046]
- Follow the per-insured vehicle counts listed below; use a different Vehicle Type for each vehicle of the same insured [Rule 36]
- Vehicle Type: MUST be EXACTLY one of (use double-hyphen --): [Rule 26]
  {vehicle_list}
- Garaging State: MUST match the insured's Rating State [Rule 27]
- ZIP Code: 5-digit numeric, MUST match insured's ZIP Code [Rule 28]
- Radius Category: MUST be one of {VALID_RADIUS} [Rule 29]
- Vehicle Use Type: MUST be one of {VALID_VEHICLE_USE} [Rule 30]
- PIP Incl?: Y or N only [Rule 31]
- Med Pay Incl?: Y or N only [Rule 32]
- UM Incl?: Y or N only [Rule 33]
{insured_block}
"""
            return row_count, rules

        # ---- Auto Rating -----------------------------------------------
        if sheet_type == "auto_rating":
            if policy_data:
                auto_insureds = _get_lob_yes_insureds(policy_data, "auto")
                if not auto_insureds:
                    return 0, ""
                row_count = len(auto_insureds)
                insured_block = (
                    "\nGenerate one row per Auto=Yes insured (in this order).\n"
                    "Primary Garaging/Rating State MUST match insured's Rating State:\n"
                    + _insured_summary(auto_insureds)
                )
            else:
                row_count = original_row_count
                insured_block = ""

            rules = f"""
RRG AUTO RATING RULES (HARD CONSTRAINTS):
- ONLY generate rows for insureds where Auto (Yes/No) = "Yes" in Policy Information [Rule 34]
- Primary Garaging / Rating State: MUST match the insured's Rating State
- Excess Coverage Selected? (Y/N): MUST be "Yes" or "No" only [Rule 34]
{insured_block}
"""
            return row_count, rules

        return original_row_count, ""

    # ------------------------------------------------------------------
    # Deterministic pre-generation (reference tables → empty)
    # ------------------------------------------------------------------

    def pre_generate(
        self,
        sheet_name: str,
        unique_headers: list[str],
        policy_data: list[dict[str, Any]] | None,
        driver_data: list[dict[str, Any]] | None,
        vehicle_data: list[dict[str, Any]] | None,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]] | None:
        sheet_type = self.detect_sheet_type(sheet_name)
        if sheet_type == "reference":
            return []
        if sheet_type == "scenario_details":
            # One summary row per data set (DS), derived from the upstream
            # Policy Information, Sched of Locations, PL_Rating, and Sched of
            # Vehicles sheets that have already been generated.
            return self._build_scenario_details(unique_headers, previous_sheets_data or {})
        return None

    # ------------------------------------------------------------------
    # Post-processing: hard constraint enforcement
    # ------------------------------------------------------------------

    def post_process(
        self,
        rows: list[dict[str, Any]],
        sheet_name: str,
        special_instruction: str,
        previous_sheets_data: dict[str, list[dict[str, Any]]] | None = None,
    ) -> list[dict[str, Any]]:
        sheet_type = self.detect_sheet_type(sheet_name)

        if sheet_type == "policy":
            rows = self._fix_policy_info(rows)
        elif sheet_type == "locations":
            rows = self._fix_locations(rows, previous_sheets_data)
        elif sheet_type == "gl_rating":
            rows = self._fix_gl_rating(rows, previous_sheets_data)
        elif sheet_type == "pl_rating":
            rows = self._fix_pl_rating(rows, previous_sheets_data)
        elif sheet_type == "ebl_rating":
            rows = self._fix_ebl_rating(rows, previous_sheets_data)
        elif sheet_type == "abuse_rating":
            rows = self._fix_abuse_rating(rows, previous_sheets_data)
        elif sheet_type == "vehicles":
            rows = self._fix_vehicles(rows, previous_sheets_data)
        elif sheet_type == "auto_rating":
            rows = self._fix_auto_rating(rows, previous_sheets_data)

        return rows

    # ------------------------------------------------------------------
    # Per-sheet fixers
    # ------------------------------------------------------------------

    def _fix_policy_info(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        # Force Test ID to DS-01, DS-02, ... format (rulebook requirement)
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if tid_key is not None:
                row[tid_key] = f"DS-{idx + 1:02d}"

        for row in rows:
            for key in list(row.keys()):
                kl = key.lower()
                val = str(row[key]) if row[key] is not None else ""
                if kl in ("state", "state of operation", "rating state"):
                    if val not in VALID_STATES:
                        row[key] = random.choice(VALID_STATES)
                elif "new" in kl and "renewal" in kl:
                    if val not in VALID_NEW_RENEWAL:
                        row[key] = random.choice(VALID_NEW_RENEWAL)
                elif "org type" in kl or ("entity" in kl and "org" in kl):
                    if val not in VALID_ORG_TYPES:
                        row[key] = random.choice(VALID_ORG_TYPES)
                elif kl == "gl (yes/no)":
                    row[key] = "Yes"
                elif kl == "contact number":
                    # Rule 1: format as USA 10-digit "XXX XXX XXXX" [DS_044]
                    row[key] = _format_us_phone(val)
                elif "producer" in kl and ("phone" in kl or "contact number" in kl):
                    # Rule 9: format as USA 10-digit "XXX XXX XXXX" [DS_045]
                    row[key] = _format_us_phone(val)
                elif kl == "zip code" or kl == "zip":
                    # Rule 3: 5-digit numeric ZIP (preserve leading zeros)
                    cleaned = _format_zip5(val)
                    if cleaned:
                        row[key] = cleaned
                elif "effective date" in kl or "expiration date" in kl:
                    # Convert MMDDYYYY (8-digit no-separator) → MM/DD/YYYY [DS_026]
                    if val and len(val) == 8 and val.isdigit():
                        row[key] = f"{val[:2]}/{val[2:4]}/{val[4:]}"
                elif "commission" in kl:
                    # Ensure Commission % always has a trailing % sign
                    if val and not val.endswith("%"):
                        row[key] = val + "%"

        # DS_042: Ensure PL and EBL (Yes/No) are not ALL "Yes" — force at
        # least 1 row to "No" for each optional LOB when row count >= 3.
        if len(rows) >= 3:
            for lob_fragment in ("pl", "ebl"):
                lob_key = None
                yes_indices = []
                for i, row in enumerate(rows):
                    for k in row:
                        if lob_fragment in k.lower() and "(yes/no)" in k.lower():
                            lob_key = k
                            if str(row[k]).strip().lower() == "yes":
                                yes_indices.append(i)
                            break
                if lob_key and len(yes_indices) == len(rows):
                    # All are "Yes" — flip 1 random row to "No"
                    flip_idx = random.choice(yes_indices[1:])  # keep at least row 0 as Yes
                    rows[flip_idx][lob_key] = "No"

        return rows

    def _fix_locations(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None = None,
    ) -> list[dict[str, Any]]:
        """Validate Class Code, State, and restamp Test IDs for location rows."""
        # Restamp Test IDs using location→insured mapping from Policy Information
        pi_rows = (previous_sheets_data or {}).get("Policy Information", [])
        if pi_rows:
            loc_counts = _location_counts_for_insureds(pi_rows)
            row_insured_map = _build_insured_row_map(loc_counts)
            for idx, row in enumerate(rows):
                tid_key = next((k for k in row if k.lower() == "test id"), None)
                if tid_key is not None and idx < len(row_insured_map):
                    insured_idx = row_insured_map[idx]
                    insured = pi_rows[min(insured_idx, len(pi_rows) - 1)]
                    row[tid_key] = insured.get("Test ID", row.get(tid_key, ""))

        # DS_043: renumber the "#" column per insured so each location maps
        # clearly back to its insured (1, 2, 3 within each Test ID) instead of
        # a confusing global 1..N running counter.
        seq_by_tid: dict[str, int] = {}
        for row in rows:
            num_key = next((k for k in row if k.strip() == "#"), None)
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if num_key and tid_key:
                tid = str(row.get(tid_key, ""))
                seq_by_tid[tid] = seq_by_tid.get(tid, 0) + 1
                row[num_key] = seq_by_tid[tid]

        for row in rows:
            # Rule 13: Class Code must be from approved list
            cc_key = _find_col(row, "class code")
            if cc_key:
                val = str(row[cc_key]).strip()
                if val not in VALID_CLASS_CODES:
                    row[cc_key] = random.choice(VALID_CLASS_CODES)

            # Rule 11: State must be a valid allowed state
            state_key = _find_col(row, "state")
            if state_key:
                if row[state_key] not in VALID_STATES:
                    row[state_key] = random.choice(VALID_STATES)

            # Rule 10: 5-digit ZIP (preserve leading zeros for CT/NJ)
            zip_key = _find_col(row, "zip")
            if zip_key:
                cleaned = _format_zip5(row.get(zip_key))
                if cleaned:
                    row[zip_key] = cleaned

        return rows

    def _fix_gl_rating(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None = None,
    ) -> list[dict[str, Any]]:
        """Fix Occurrence Selection, Schedule Rating, and Test IDs (Rules 15/17)."""
        # Restamp Test IDs from Policy Information (1:1 — all insureds have GL=Yes)
        pi_rows = (previous_sheets_data or {}).get("Policy Information", [])
        for idx, row in enumerate(rows):
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if tid_key is not None and idx < len(pi_rows):
                row[tid_key] = pi_rows[idx].get("Test ID", row.get(tid_key, ""))

        for row in rows:
            # Rule 15: valid Claims Made/Occurrence values
            key = _find_col(row, "claims made", "occurrence")
            if key:
                val = str(row[key]).strip() if row[key] is not None else ""
                if val.lower() == "claims made":
                    row[key] = random.choice(VALID_GL_CLAIMS_MADE_RETRO)
                elif val not in VALID_GL_OCCURRENCE:
                    row[key] = random.choice(VALID_GL_OCCURRENCE)

            # Rule 16: valid Medical Payments Limit
            med_key = _find_col(row, "medical payment")
            if med_key:
                med_val = str(row[med_key]).strip() if row[med_key] is not None else ""
                if med_val not in VALID_MED_PAY_LIMITS:
                    row[med_key] = random.choice(VALID_MED_PAY_LIMITS)

            # Rule 17: schedule rating must be non-zero (credit=negative, debit=positive)
            sched_key = _find_col(row, "schedule rating")
            if sched_key:
                sched_val = str(row[sched_key]).strip().replace("%", "")
                try:
                    numeric = float(sched_val)
                    if numeric == 0:
                        # Replace 0% with a random credit or debit
                        row[sched_key] = random.choice(["-5%", "-10%", "-15%", "5%", "8%", "10%", "12%"])
                except ValueError:
                    pass

        return rows

    def _fix_pl_rating(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None,
    ) -> list[dict[str, Any]]:
        """Enforce Policy Basis consistency, multi-CoB rows, and validate fields.

        DS_038: PL rows may contain multiple Class of Business entries for a
        single insured.  We use the same _cob_counts_for_insureds() cycling
        pattern [2, 1, 2] that the prompt builder uses, so each flat row index
        maps to the correct insured via _build_insured_row_map().
        """
        gl_rows = (previous_sheets_data or {}).get("GL_Rating", [])

        pi_rows = (previous_sheets_data or {}).get("Policy Information", [])
        pl_insureds = _get_lob_yes_insureds(pi_rows, "pl") if pi_rows else []

        # Build row→insured mapping using the same CoB count pattern as build_sheet_context
        if pl_insureds:
            cob_counts = _cob_counts_for_insureds(pl_insureds)
            expected_row_count = sum(cob_counts)
            row_insured_map = _build_insured_row_map(cob_counts)
            # Trim excess rows if LLM generated more than expected
            if len(rows) > expected_row_count:
                rows = rows[:expected_row_count]
        else:
            row_insured_map = list(range(len(rows)))  # fallback: 1:1

        # DS_039: Map each PL=Yes insured back to its *global* position in
        # Policy Information so we can index into GL_Rating correctly.
        # (PL insured idx 1 might be PI row 2 if PI row 1 has PL=No.)
        pl_to_global_idx: list[int] = []
        for pl_ins in pl_insureds:
            pl_tid = pl_ins.get("Test ID", "")
            for gi, pi_row in enumerate(pi_rows):
                if pi_row.get("Test ID", "") == pl_tid:
                    pl_to_global_idx.append(gi)
                    break
            else:
                pl_to_global_idx.append(len(pl_to_global_idx))

        # Track which CoB values have been used per insured to avoid duplicates
        used_cobs_per_insured: dict[int, set[str]] = {}

        for idx, row in enumerate(rows):
            # Determine which insured this row belongs to
            if idx < len(row_insured_map):
                insured_idx = row_insured_map[idx]
            else:
                insured_idx = idx  # fallback

            insured = (
                pl_insureds[min(insured_idx, len(pl_insureds) - 1)]
                if pl_insureds else {}
            )
            # Global PI index for GL lookup
            global_idx = (
                pl_to_global_idx[insured_idx]
                if insured_idx < len(pl_to_global_idx)
                else insured_idx
            )

            # DS_030: stamp Test ID from the actual PL=Yes insured
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if tid_key and insured:
                row[tid_key] = insured.get("Test ID", row.get(tid_key, ""))

            # DS_039: Policy Basis from GL — use global_idx to match the correct
            # GL row (the insured's position in full Policy Information list,
            # not its position within PL=Yes-only insureds)
            pb_key = _find_col(row, "policy", "basis")
            if pb_key:
                desired = self._derive_policy_basis(global_idx, gl_rows)
                if desired:
                    row[pb_key] = desired
                elif row[pb_key] not in VALID_PL_POLICY_BASIS:
                    row[pb_key] = random.choice(VALID_PL_POLICY_BASIS)

            # Hazard Tier
            ht_key = _find_col(row, "hazard tier")
            if ht_key and row[ht_key] not in VALID_HAZARD_TIER:
                row[ht_key] = random.choice(VALID_HAZARD_TIER)

            # Class of Business — ensure valid and unique within same insured (DS_038)
            cob_key = _find_col(row, "class of business")
            if cob_key:
                val = str(row[cob_key]).strip() if row[cob_key] else ""
                used = used_cobs_per_insured.setdefault(insured_idx, set())
                if val not in VALID_COB or val in used:
                    # Pick a CoB not yet used by this insured
                    available = [c for c in VALID_COB if c not in used]
                    if not available:
                        available = list(VALID_COB)
                    row[cob_key] = random.choice(available)
                used.add(str(row[cob_key]))

            # Rule 19: retroactive date must be <= effective date
            retro_key = _find_col(row, "retroactive date")
            if retro_key and insured:
                eff_dt   = _parse_date(insured.get("Effective Date", ""))
                retro_dt = _parse_date(str(row.get(retro_key, "")))
                if eff_dt and retro_dt and retro_dt > eff_dt:
                    row[retro_key] = eff_dt.strftime("%m/%d/%Y")

        return rows

    def _fix_ebl_rating(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None,
    ) -> list[dict[str, Any]]:
        """Enforce Policy Basis consistency, fix employee count types, sync retro dates with PL.

        DS_037: Policy Basis MUST be derived from GL_Rating using the insured's
        position (not the EBL row index, which is always 1:1 here but we use
        insured_idx for safety).
        """
        gl_rows  = (previous_sheets_data or {}).get("GL_Rating", [])
        pl_rows  = (previous_sheets_data or {}).get("PL_Rating", [])
        ebl_pi_rows = (previous_sheets_data or {}).get("Policy Information", [])
        ebl_insureds = _get_lob_yes_insureds(ebl_pi_rows, "ebl") if ebl_pi_rows else []

        # Build a lookup: Test ID → first PL row (for retro-date sync)
        pl_by_tid: dict[str, dict] = {}
        for pl_row in pl_rows:
            tid = str(pl_row.get("Test ID", pl_row.get(
                next((k for k in pl_row if k.lower() == "test id"), ""), ""
            ))).strip()
            if tid and tid not in pl_by_tid:
                pl_by_tid[tid] = pl_row  # first CoB row per insured

        # Also map each EBL=Yes insured back to its *global* position in Policy
        # Information so we can index into GL_Rating correctly (DS_037).
        ebl_to_global_idx: list[int] = []
        for ebl_ins in ebl_insureds:
            ebl_tid = ebl_ins.get("Test ID", "")
            for gi, pi_row in enumerate(ebl_pi_rows):
                if pi_row.get("Test ID", "") == ebl_tid:
                    ebl_to_global_idx.append(gi)
                    break
            else:
                ebl_to_global_idx.append(len(ebl_to_global_idx))

        for idx, row in enumerate(rows):
            insured = ebl_insureds[idx] if idx < len(ebl_insureds) else {}
            global_idx = ebl_to_global_idx[idx] if idx < len(ebl_to_global_idx) else idx

            # DS_031: stamp Test ID from the actual EBL=Yes insured at this position
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if tid_key and insured:
                row[tid_key] = insured.get("Test ID", row.get(tid_key, ""))

            # DS_037: Policy Basis from GL — use global_idx to index the correct
            # GL row (the insured's position in Policy Information, not the EBL
            # row index, which may differ when not all insureds have EBL=Yes)
            pb_key = _find_col(row, "policy", "basis")
            if pb_key:
                desired = self._derive_policy_basis(global_idx, gl_rows)
                if desired:
                    row[pb_key] = desired
                elif row[pb_key] not in VALID_PL_POLICY_BASIS:
                    row[pb_key] = random.choice(VALID_PL_POLICY_BASIS)

            # Rule 22: Total Employee Count must be a numeric integer, not a string
            emp_key = _find_col(row, "employee count")
            if emp_key:
                val = row[emp_key]
                if val is not None:
                    try:
                        row[emp_key] = int(str(val).strip())
                    except (ValueError, TypeError):
                        row[emp_key] = random.randint(1, 100)

            retro_key = _find_col(row, "retroactive date")
            if retro_key:
                # Sync with PL retroactive date — match by Test ID (PL now has
                # multi-CoB rows so index-based matching is invalid)
                insured_tid = insured.get("Test ID", "")
                if insured_tid and insured_tid in pl_by_tid:
                    pl_retro_key = _find_col(pl_by_tid[insured_tid], "retroactive date")
                    if pl_retro_key:
                        pl_retro_val = pl_by_tid[insured_tid].get(pl_retro_key, "")
                        if pl_retro_val:
                            row[retro_key] = pl_retro_val

                # Rule 24: retroactive date must be <= effective date (independent guard)
                if insured:
                    eff_dt   = _parse_date(insured.get("Effective Date", ""))
                    retro_dt = _parse_date(str(row.get(retro_key, "")))
                    if eff_dt and retro_dt and retro_dt > eff_dt:
                        row[retro_key] = eff_dt.strftime("%m/%d/%Y")

        return rows

    def _fix_abuse_rating(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None = None,
    ) -> list[dict[str, Any]]:
        """Ensure Abuse Hazard Tier values are valid and Test IDs match Abuse=Yes insureds."""
        pi_rows = (previous_sheets_data or {}).get("Policy Information", [])
        abuse_insureds = _get_lob_yes_insureds(pi_rows, "abuse") if pi_rows else []

        for idx, row in enumerate(rows):
            # Restamp Test ID to the correct Abuse=Yes insured at this position
            tid_key = next((k for k in row if k.lower() == "test id"), None)
            if tid_key and abuse_insureds and idx < len(abuse_insureds):
                correct_id = abuse_insureds[idx].get("Test ID", "")
                if correct_id:
                    row[tid_key] = correct_id

            # Rule 25: valid Abuse Hazard Tier values
            key = _find_col(row, "abuse hazard tier")
            if key and row[key] not in VALID_ABUSE_HAZARD:
                row[key] = random.choice(VALID_ABUSE_HAZARD)

        return rows

    def _fix_vehicles(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None = None,
    ) -> list[dict[str, Any]]:
        """Normalize vehicle types and align garaging state/ZIP to Auto=Yes insureds."""
        # Align each vehicle row to the corresponding Auto=Yes insured
        auto_insureds: list[dict] = []
        pi_rows: list[dict] = []
        if previous_sheets_data:
            pi_rows = previous_sheets_data.get("Policy Information", [])
            auto_insureds = _get_lob_yes_insureds(pi_rows, "auto")

        # Build row→insured mapping using the same variable-count function as build_sheet_context
        if auto_insureds:
            veh_counts   = _vehicle_counts_for_insureds(auto_insureds)
            row_insured_map = _build_insured_row_map(veh_counts)
            # DS_050: never carry more vehicle rows than the Auto=Yes insureds
            # account for. Extra LLM-hallucinated rows would otherwise surface a
            # populated Vehicle Type for an insured that owns no vehicles.
            expected_row_count = sum(veh_counts)
            if len(rows) > expected_row_count:
                rows = rows[:expected_row_count]
        elif pi_rows:
            # Policy Information is known and NO insured has Auto=Yes → there must
            # be no vehicle rows at all (and so no stray Vehicle Type values).
            # [Rule 36 / DS_050]
            return []
        else:
            # No policy context available (legacy fallback) — can't validate
            # against insureds, so leave the LLM rows in place.
            row_insured_map = []

        for idx, row in enumerate(rows):
            # Vehicle Type: normalize dashes (Rule 26)
            vt_key = _find_col(row, "vehicle type")
            if vt_key:
                row[vt_key] = _normalize_vehicle_type(str(row[vt_key]) if row[vt_key] else "")

            # Map this vehicle row to its Auto=Yes insured using pre-computed mapping
            gs_key  = _find_col(row, "garaging state")
            zip_key = _find_col(row, "zip code")
            if auto_insureds and idx < len(row_insured_map):
                insured_idx = row_insured_map[idx]
                insured = auto_insureds[min(insured_idx, len(auto_insureds) - 1)]
                if gs_key:
                    row[gs_key] = insured.get("Rating State", row.get(gs_key, ""))
                if zip_key:
                    # Rule 28: 5-digit ZIP (preserve leading zeros for CT/NJ)
                    raw_zip = insured.get("ZIP Code", row.get(zip_key, ""))
                    row[zip_key] = _format_zip5(raw_zip) or raw_zip
                # Stamp Test ID to match the correct Auto=Yes insured
                tid_key = next((k for k in row if k.lower() == "test id"), None)
                if tid_key and insured.get("Test ID"):
                    row[tid_key] = insured["Test ID"]
            elif gs_key and row.get(gs_key, "") not in VALID_STATES:
                row[gs_key] = random.choice(VALID_STATES)

            # Radius Category (Rule 29)
            rc_key = _find_col(row, "radius")
            if rc_key and row[rc_key] not in VALID_RADIUS:
                row[rc_key] = random.choice(VALID_RADIUS)

            # Vehicle Use Type (Rule 30)
            vu_key = _find_col(row, "vehicle use")
            if vu_key and row[vu_key] not in VALID_VEHICLE_USE:
                row[vu_key] = random.choice(VALID_VEHICLE_USE)

            # PIP / Med Pay / UM flags (Rules 31-33)
            for col_fragment in (("pip",), ("med pay",), ("um incl",)):
                flag_key = _find_col(row, *col_fragment)
                if flag_key and str(row[flag_key]).strip().upper() not in ("Y", "N"):
                    row[flag_key] = random.choice(["Y", "N"])

        # DS_046 / Rule 36: guarantee EVERY Auto=Yes insured has at least 1
        # vehicle row. If the LLM skipped an insured, synthesize a valid vehicle
        # row from a template so no Auto=Yes insured is left with zero vehicles.
        if auto_insureds and rows:
            def _row_tid(r: dict) -> str:
                tk = next((k for k in r if k.lower() == "test id"), None)
                return str(r.get(tk, "")) if tk else ""

            present_tids = {_row_tid(r) for r in rows}
            template = rows[0]
            for i, insured in enumerate(auto_insureds, 1):
                tid = str(insured.get("Test ID", ""))
                if not tid or tid in present_tids:
                    continue
                new_row = dict(template)
                tid_key = next((k for k in new_row if k.lower() == "test id"), None)
                if tid_key:
                    new_row[tid_key] = tid
                if gs_key := _find_col(new_row, "garaging state"):
                    new_row[gs_key] = insured.get("Rating State", insured.get("State", ""))
                if zip_key := _find_col(new_row, "zip code"):
                    raw_zip = insured.get("ZIP Code", "")
                    new_row[zip_key] = _format_zip5(raw_zip) or raw_zip
                if unit_key := _find_col(new_row, "unit"):
                    new_row[unit_key] = f"{i:02d}A"
                rows.append(new_row)
                present_tids.add(tid)

        # Renumber the "#" column sequentially so it stays coherent after any
        # synthesized rows above.
        for n, row in enumerate(rows, 1):
            num_key = next((k for k in row if k.strip() == "#"), None)
            if num_key:
                row[num_key] = n

        return rows

    def _fix_auto_rating(
        self,
        rows: list[dict[str, Any]],
        previous_sheets_data: dict | None = None,
    ) -> list[dict[str, Any]]:
        """Align state to Auto=Yes insureds and validate Excess Coverage (Rule 34)."""
        auto_insureds: list[dict] = []
        if previous_sheets_data:
            pi_rows = previous_sheets_data.get("Policy Information", [])
            auto_insureds = _get_lob_yes_insureds(pi_rows, "auto")

        for idx, row in enumerate(rows):
            # Primary Garaging State must match Auto=Yes insured's state
            state_key = _find_col(row, "garaging", "state") or _find_col(row, "rating state")
            if state_key is None:
                # Try primary garaging / rating state header
                for k in row:
                    if "garaging" in k.lower() or ("primary" in k.lower() and "state" in k.lower()):
                        state_key = k
                        break

            if state_key and idx < len(auto_insureds):
                row[state_key] = auto_insureds[idx].get(
                    "Rating State", row.get(state_key, "")
                )

            # Excess Coverage Selected: Yes/No only (Rule 34)
            exc_key = _find_col(row, "excess coverage")
            if exc_key:
                val = str(row[exc_key]).strip()
                if val.upper() in ("Y", "YES", "TRUE", "1"):
                    row[exc_key] = "Yes"
                elif val.upper() in ("N", "NO", "FALSE", "0"):
                    row[exc_key] = "No"
                elif val not in ("Yes", "No"):
                    row[exc_key] = random.choice(["Yes", "No"])

        return rows

    # ------------------------------------------------------------------
    # Test Scenerio Details (deterministic summary, one row per DS)
    # ------------------------------------------------------------------

    def _build_scenario_details(
        self,
        unique_headers: list[str],
        previous_sheets_data: dict[str, list[dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        """Build one summary row per insured (DS) for the Test Scenerio Details sheet.

        Columns (per the client scenario spec):
        - Scenario ID            → the insured's Test ID (DS-01, DS-02, ...)
        - State                  → Policy Information "State"
        - Org Type/Entity        → Policy Information "Org Type / Entity"
        - Location Count         → # of Sched of Locations rows for the insured
        - Class Code Count       → # of DISTINCT class codes across those locations
        - Class of Business Count→ # of DISTINCT Class of Business rows in PL_Rating
        - Vehicle Count          → # of Sched of Vehicles rows for the insured
        """
        pi_rows  = previous_sheets_data.get("Policy Information", []) or []
        loc_rows = previous_sheets_data.get("Sched of Locations", []) or []
        pl_rows  = previous_sheets_data.get("PL_Rating", []) or []
        veh_rows = previous_sheets_data.get("Sched of Vehicles", []) or []

        if not pi_rows:
            return []

        def _hdr(*frags: str) -> str | None:
            """First scenario header containing ALL fragments (case-insensitive)."""
            for h in unique_headers:
                hl = h.lower()
                if all(f in hl for f in frags):
                    return h
            return None

        sid_key = _hdr("scenario")
        state_key = _hdr("state")
        org_key = _hdr("org")
        loc_count_key = _hdr("location", "count")
        cc_count_key = _hdr("class code")
        cob_count_key = _hdr("class of business")
        veh_count_key = _hdr("vehicle", "count")

        def _tid(row: dict) -> str:
            tk = next((k for k in row if k.lower() == "test id"), None)
            return str(row.get(tk, "")).strip() if tk else ""

        # Aggregate per Test ID across the upstream sheets.
        loc_count: dict[str, int] = {}
        class_codes: dict[str, set] = {}
        for row in loc_rows:
            tid = _tid(row)
            if not tid:
                continue
            loc_count[tid] = loc_count.get(tid, 0) + 1
            cc_key = _find_col(row, "class code")
            if cc_key:
                cc_val = str(row.get(cc_key, "")).strip()
                if cc_val:
                    class_codes.setdefault(tid, set()).add(cc_val)

        cob_count: dict[str, set] = {}
        for row in pl_rows:
            tid = _tid(row)
            if not tid:
                continue
            cob_key = _find_col(row, "class of business")
            cob_val = str(row.get(cob_key, "")).strip() if cob_key else ""
            cob_count.setdefault(tid, set())
            if cob_val:
                cob_count[tid].add(cob_val)

        veh_count: dict[str, int] = {}
        for row in veh_rows:
            tid = _tid(row)
            if tid:
                veh_count[tid] = veh_count.get(tid, 0) + 1

        # Resolve Policy Information state / org-type keys once.
        pi_state_key = next(
            (k for k in pi_rows[0] if k.strip().lower() == "state"), None
        )
        pi_org_key = next((k for k in pi_rows[0] if "org type" in k.lower()), None)

        scenario_rows: list[dict[str, Any]] = []
        for pi in pi_rows:
            tid = _tid(pi)
            row: dict[str, Any] = {}
            if sid_key:
                row[sid_key] = tid
            if state_key:
                row[state_key] = pi.get(pi_state_key, "") if pi_state_key else ""
            if org_key:
                row[org_key] = pi.get(pi_org_key, "") if pi_org_key else ""
            if loc_count_key:
                row[loc_count_key] = loc_count.get(tid, 0)
            if cc_count_key:
                row[cc_count_key] = len(class_codes.get(tid, set()))
            if cob_count_key:
                row[cob_count_key] = len(cob_count.get(tid, set()))
            if veh_count_key:
                row[veh_count_key] = veh_count.get(tid, 0)
            scenario_rows.append(row)

        return scenario_rows

    # ------------------------------------------------------------------
    # Helper
    # ------------------------------------------------------------------

    def _derive_policy_basis(self, row_idx: int, gl_rows: list[dict]) -> str | None:
        """Return Policy Basis for PL/EBL derived from the matching GL row."""
        if row_idx >= len(gl_rows):
            return None
        gl_row = gl_rows[row_idx]
        gl_key = _find_col(gl_row, "claims made", "occurrence")
        if not gl_key:
            return None
        gl_val = str(gl_row.get(gl_key, "")).strip()
        if gl_val == "Occurrence":
            return "Occurrence"
        if gl_val in VALID_GL_CLAIMS_MADE_RETRO:
            return "Claims Made"
        return None
