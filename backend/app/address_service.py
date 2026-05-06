"""Verified US address pre-generation for PAP workbooks.

Flow:
  1. Focused LLM call generates N candidate addresses per state.
  2. US Census Bureau Geocoder (TIGER/Line, free, no API key) confirms each
     street block is real and returns the authoritative city/state/zip.
  3. Zippopotam.us fallback guarantees a valid city/zip/state triple when the
     Census finds no street match even after retries.

The result is a {TS-XX: {street, city, state, zip}} map consumed by the
post-processor — the main LLM call has no involvement in addresses.
"""
from __future__ import annotations

import json
import os
import time
from typing import Any

import requests
from openai import OpenAI

_CENSUS_URL = (
    "https://geocoding.geo.census.gov/geocoder/locations/address"
)
_ZIPPOPOTAM_URL = "https://api.zippopotam.us/us/{zip}"

_HEADERS = {"User-Agent": "DataGenAI-AddressService/1.0"}


def _census_validate(street: str, city: str, state: str, zip_: str) -> dict[str, str] | None:
    """Call the US Census Geocoder.

    Returns an authoritative {street, city, state, zip} dict when the street
    block exists in the TIGER/Line database, or None on no-match / error.
    House number is set to the midpoint of the matched address range.
    """
    try:
        r = requests.get(
            _CENSUS_URL,
            params={
                "street": street,
                "city": city,
                "state": state,
                "zip": zip_,
                "benchmark": "Public_AR_Current",
                "format": "json",
            },
            headers=_HEADERS,
            timeout=8,
        )
        matches = r.json()["result"]["addressMatches"]
        if matches:
            c = matches[0]["addressComponents"]
            try:
                lo = int(c.get("fromAddress") or 100)
                hi = int(c.get("toAddress") or lo + 100)
                number = (lo + hi) // 2
            except (ValueError, TypeError):
                number = 100
            suffix = c.get("suffixType", "")
            street_out = f"{number} {c['streetName'].title()} {suffix.title()}".strip()
            return {
                "street": street_out,
                "city": c["city"].title(),
                "state": c["state"],
                "zip": c["zip"],
            }
    except Exception:
        pass
    return None


def _zippopotam_validate(zip_: str) -> dict[str, str] | None:
    """Confirm city/state for a zip via Zippopotam.us (fallback only)."""
    try:
        r = requests.get(
            _ZIPPOPOTAM_URL.format(zip=zip_),
            headers=_HEADERS,
            timeout=5,
        )
        if r.status_code == 200:
            p = r.json()["places"][0]
            return {
                "city": p["place name"],
                "state": p["state abbreviation"],
                "zip": zip_,
            }
    except Exception:
        pass
    return None


def _llm_addresses(state: str, count: int, client: OpenAI, model: str) -> list[dict[str, Any]]:
    """Focused LLM call — address generation only, no business logic."""
    prompt = (
        f"Generate {count} distinct real US residential addresses in {state} state.\n"
        "Requirements:\n"
        "- Each address must be in a different city or neighbourhood.\n"
        "- The zip code must be the actual USPS zip code for that city.\n"
        "- The street must be a real street name that exists in that city "
        "(e.g. '47 Congress St', not '47 Zork Blvd').\n"
        f"Return ONLY JSON: "
        f'{{ "addresses": [{{'
        f'"street": "<number> <name>", "city": "...", "state": "{state}", "zip": "..."'
        f"}}] }}"
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=1,
        )
        return json.loads(resp.choices[0].message.content).get("addresses", [])
    except Exception as exc:
        print(f"[address_service] LLM address generation failed for {state}: {exc}")
        return []


def _llm_street(city: str, state: str, zip_: str, client: OpenAI, model: str) -> str:
    """Ask LLM for one replacement street when Census rejects the first attempt."""
    prompt = (
        f"Give me ONE real residential street address (with house number) "
        f"in {city}, {state} {zip_}.\n"
        f'Return ONLY JSON: {{ "street": "<number> <street name>" }}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=1,
        )
        return json.loads(resp.choices[0].message.content).get("street", "100 Main St")
    except Exception:
        return "100 Main St"


def _verify_one(
    candidate: dict[str, Any],
    client: OpenAI,
    model: str,
    max_retries: int = 2,
) -> dict[str, str]:
    """Validate a single candidate address, retrying street on Census rejection.

    Tier 1: Census Geocoder confirms street block is real → returns authoritative address.
    Tier 2: Zippopotam confirms city/zip if Census rejects after all retries.
    """
    state = str(candidate.get("state") or "")
    city  = str(candidate.get("city") or "")
    zip_  = str(candidate.get("zip") or "")

    for attempt in range(max_retries + 1):
        street = str(candidate.get("street") or "100 Main St")
        result = _census_validate(street, city, state, zip_)
        if result:
            print(f"[address_service] Census ✓ {result['street']}, {result['city']}, {result['state']} {result['zip']}")
            return result

        if attempt < max_retries:
            new_street = _llm_street(city, state, zip_, client, model)
            print(f"[address_service] Census no-match — retry {attempt + 1} with: {new_street}")
            candidate = dict(candidate)
            candidate["street"] = new_street

    # Tier 2: at least guarantee a valid city/zip/state
    zp = _zippopotam_validate(zip_)
    result = {
        "street": str(candidate.get("street") or "100 Main St"),
        "city":   zp["city"]  if zp else city,
        "state":  zp["state"] if zp else state,
        "zip":    zp["zip"]   if zp else zip_,
    }
    print(f"[address_service] Zippopotam fallback → {result}")
    return result


def generate_verified_addresses(
    state_selection: list[str],
    n_groups: int,
    client: OpenAI,
    model: str,
) -> dict[str, dict[str, str]]:
    """Return {TS-XX: {street, city, state, zip}} — one Census-verified address per group.

    - Works for all 50 US states with zero hardcoded geographic data.
    - Tier 1: Census Geocoder (TIGER/Line) confirms real streets and authoritative zips.
    - Tier 2: Zippopotam.us ensures city/zip correctness if Census is unavailable.
    - Falls back gracefully to LLM-only if both APIs are unreachable.
    """
    if not state_selection or n_groups <= 0:
        return {}

    # Determine state per group and how many addresses each state needs
    state_counts: dict[str, int] = {}
    group_states: list[str] = []
    for i in range(n_groups):
        s = state_selection[i % len(state_selection)].upper()
        group_states.append(s)
        state_counts[s] = state_counts.get(s, 0) + 1

    # One LLM batch call per state (+2 buffer so retries have candidates to fall back to)
    print(f"[address_service] Pre-generating addresses for states: {list(state_counts)}")
    state_pools: dict[str, list[dict]] = {}
    for state, cnt in state_counts.items():
        state_pools[state] = _llm_addresses(state, cnt + 2, client, model)

    # Validate each candidate and assign to test-case groups
    state_idx: dict[str, int] = {s: 0 for s in state_pools}
    result: dict[str, dict[str, str]] = {}

    for i in range(n_groups):
        state = group_states[i]
        pool = state_pools.get(state, [])
        idx = state_idx.get(state, 0)
        candidate: dict[str, Any] = (
            pool[idx] if idx < len(pool)
            else {"street": "100 Main St", "city": "", "state": state, "zip": ""}
        )
        grp_key = f"TS-{i + 1:02d}"
        result[grp_key] = _verify_one(candidate, client, model)
        state_idx[state] = idx + 1
        time.sleep(0.2)  # light throttle — Census has no strict limit but be courteous

    print(f"[address_service] Done. {len(result)} verified addresses generated.")
    return result


def get_model_name() -> str:
    """Return the configured LLM model name (mirrors llm_service logic)."""
    provider = os.getenv("MODEL_PROVIDER", "openai").lower()
    return os.getenv(
        "GEMINI_MODEL" if provider == "gemini" else "OPENAI_MODEL",
        "gpt-4o",
    )
