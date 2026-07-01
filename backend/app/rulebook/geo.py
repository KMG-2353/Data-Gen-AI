"""Canonical US City/State/ZIP reference (single source of truth).

Class-B defect (2026-07-01, logged on every SPG LOB — IM, DW, HO, Cargo, WH,
APD): the generated address had a **valid** ZIP that belonged to a *different*
state than the one in the row, so City/State/ZIP did not correspond. The root
cause is structural: :func:`app.rulebook.variety.spread_address_states` fans the
State column across the full pool for variety, but the City and ZIP were left as
the LLM produced them (for the original state) — so after the spread they no
longer match.

City/State/ZIP correspondence is a **deterministic** field (CLAUDE.md §2: ZIP and
state are on the deterministic side of the boundary): the engine must *produce* a
consistent triple, not hope the LLM keeps them aligned. This module holds one real
(city, ZIP) representative set per state so the variety pass can assign a whole
consistent block — State **and** its sibling City/ZIP — together.

Each ZIP is a real 5-digit ZIP located in the paired city/state. ZIPs are stored
as strings to preserve leading zeros (e.g. Massachusetts ``"02108"``). Keys cover
every state in :data:`app.rulebook.primitives.US_STATES`.
"""
from __future__ import annotations

# state -> tuple of (city, zip5) real representatives. Multiple entries per state
# keep address variety while guaranteeing City/State/ZIP correspond.
STATE_GEO: dict[str, tuple[tuple[str, str], ...]] = {
    "AL": (("Birmingham", "35203"), ("Montgomery", "36104"), ("Mobile", "36602")),
    "AK": (("Anchorage", "99501"), ("Fairbanks", "99701"), ("Juneau", "99801")),
    "AZ": (("Phoenix", "85004"), ("Tucson", "85701"), ("Mesa", "85201")),
    "AR": (("Little Rock", "72201"), ("Fayetteville", "72701"), ("Fort Smith", "72901")),
    "CA": (("Los Angeles", "90012"), ("San Diego", "92101"), ("Sacramento", "95814")),
    "CO": (("Denver", "80202"), ("Colorado Springs", "80903"), ("Boulder", "80302")),
    "CT": (("Hartford", "06103"), ("New Haven", "06510"), ("Stamford", "06901")),
    "DE": (("Wilmington", "19801"), ("Dover", "19901"), ("Newark", "19711")),
    "FL": (("Miami", "33130"), ("Orlando", "32801"), ("Tampa", "33602")),
    "GA": (("Atlanta", "30303"), ("Savannah", "31401"), ("Augusta", "30901")),
    "HI": (("Honolulu", "96813"), ("Hilo", "96720"), ("Kailua", "96734")),
    "ID": (("Boise", "83702"), ("Idaho Falls", "83402"), ("Nampa", "83651")),
    "IL": (("Chicago", "60602"), ("Springfield", "62701"), ("Peoria", "61602")),
    "IN": (("Indianapolis", "46204"), ("Fort Wayne", "46802"), ("Evansville", "47708")),
    "IA": (("Des Moines", "50309"), ("Cedar Rapids", "52401"), ("Davenport", "52801")),
    "KS": (("Wichita", "67202"), ("Topeka", "66603"), ("Kansas City", "66101")),
    "KY": (("Louisville", "40202"), ("Lexington", "40507"), ("Frankfort", "40601")),
    "LA": (("New Orleans", "70112"), ("Baton Rouge", "70801"), ("Shreveport", "71101")),
    "ME": (("Portland", "04101"), ("Augusta", "04330"), ("Bangor", "04401")),
    "MD": (("Baltimore", "21202"), ("Annapolis", "21401"), ("Rockville", "20850")),
    "MA": (("Boston", "02108"), ("Worcester", "01608"), ("Springfield", "01103")),
    "MI": (("Detroit", "48226"), ("Grand Rapids", "49503"), ("Lansing", "48933")),
    "MN": (("Minneapolis", "55401"), ("Saint Paul", "55102"), ("Rochester", "55901")),
    "MS": (("Jackson", "39201"), ("Gulfport", "39501"), ("Biloxi", "39530")),
    "MO": (("Kansas City", "64106"), ("St. Louis", "63101"), ("Springfield", "65806")),
    "MT": (("Billings", "59101"), ("Missoula", "59802"), ("Helena", "59601")),
    "NE": (("Omaha", "68102"), ("Lincoln", "68508"), ("Grand Island", "68801")),
    "NV": (("Las Vegas", "89101"), ("Reno", "89501"), ("Carson City", "89701")),
    "NH": (("Manchester", "03101"), ("Concord", "03301"), ("Nashua", "03060")),
    "NJ": (("Newark", "07102"), ("Jersey City", "07302"), ("Trenton", "08608")),
    "NM": (("Albuquerque", "87102"), ("Santa Fe", "87501"), ("Las Cruces", "88001")),
    "NY": (("New York", "10007"), ("Buffalo", "14202"), ("Albany", "12207")),
    "NC": (("Charlotte", "28202"), ("Raleigh", "27601"), ("Greensboro", "27401")),
    "ND": (("Fargo", "58102"), ("Bismarck", "58501"), ("Grand Forks", "58201")),
    "OH": (("Columbus", "43215"), ("Cleveland", "44113"), ("Cincinnati", "45202")),
    "OK": (("Oklahoma City", "73102"), ("Tulsa", "74103"), ("Norman", "73069")),
    "OR": (("Portland", "97204"), ("Salem", "97301"), ("Eugene", "97401")),
    "PA": (("Philadelphia", "19107"), ("Pittsburgh", "15222"), ("Harrisburg", "17101")),
    "RI": (("Providence", "02903"), ("Warwick", "02886"), ("Cranston", "02920")),
    "SC": (("Columbia", "29201"), ("Charleston", "29401"), ("Greenville", "29601")),
    "SD": (("Sioux Falls", "57104"), ("Rapid City", "57701"), ("Pierre", "57501")),
    "TN": (("Nashville", "37203"), ("Memphis", "38103"), ("Knoxville", "37902")),
    "TX": (("Houston", "77002"), ("Dallas", "75201"), ("Austin", "78701")),
    "UT": (("Salt Lake City", "84101"), ("Provo", "84601"), ("Ogden", "84401")),
    "VT": (("Burlington", "05401"), ("Montpelier", "05602"), ("Rutland", "05701")),
    "VA": (("Richmond", "23219"), ("Norfolk", "23510"), ("Alexandria", "22314")),
    "WA": (("Seattle", "98101"), ("Spokane", "99201"), ("Tacoma", "98402")),
    "WV": (("Charleston", "25301"), ("Huntington", "25701"), ("Morgantown", "26505")),
    "WI": (("Milwaukee", "53202"), ("Madison", "53703"), ("Green Bay", "54301")),
    "WY": (("Cheyenne", "82001"), ("Casper", "82601"), ("Laramie", "82070")),
}


def geo_for_state(state: str) -> tuple[tuple[str, str], ...]:
    """Return the (city, zip) representatives for a state, or () if unknown."""
    return STATE_GEO.get((state or "").strip().upper(), ())
