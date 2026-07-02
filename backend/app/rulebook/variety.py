"""Universal dropdown-variety enforcement (deterministic, data-driven).

The recurring QA defect across every SPG rater (IM, DW, HO, Cargo, APD) is not an
LLM-reasoning failure — it is a *determinism* failure: the model collapses
dropdown fields onto one value (every address state = the binding state, every
phone left unformatted), so QA never sees the combinations the dropdown allows.

Per the determinism boundary (CLAUDE.md §2), the fix is to *produce* the value,
not hope the LLM varies it. This module holds the two universal, template-agnostic
passes that close the biggest defect classes for the whole SPG family at once:

  * ``format_phone_fax`` — snap every phone/fax/contact-number cell to the
    canonical U.S. format (the ``format_us_phone`` primitive).
  * ``spread_address_states`` — fan physical-address state columns across the
    full :data:`~app.rulebook.primitives.US_STATES` pool so the dataset exercises
    many states, *never* collapsing onto the restricted Binding/Rating state.

Both are idempotent and operate on the generated row dicts + header names alone —
no per-template column list — so they stay adaptive to new raters that follow the
``*State*`` / ``Binding State`` naming convention. Template-specific dropdowns
(Radius, Type of Entity, …) stay in the per-template handlers/profiles.

Import direction stays handlers/main -> rulebook: this module imports only
primitives.
"""
from __future__ import annotations

import hashlib
import re
from typing import Any, Sequence

from .geo import geo_for_state
from .primitives import (
    US_STATES,
    column_seed,
    find_col,
    format_us_phone,
    is_binding_state_field,
    spread_pick,
    tid_value,
)


def _is_blank(val: Any) -> bool:
    return val is None or str(val).strip() == ""


def _number_col(row: dict[str, Any]) -> str | None:
    """The literal sequential ``#`` column, if present (not ``MC#`` / ``Item #``)."""
    return next((k for k in row if k.strip() == "#"), None)


def _tid_key(row: dict[str, Any]) -> str | None:
    """The row's ``Test ID`` column name (the cross-sheet join key), if present."""
    return next((k for k in row if k.lower().strip() == "test id"), None)


# How far above ``min_per_tid`` a per-insured count may vary when ``vary_key`` is
# supplied (CARGO-008 / APD-014): keeps sheets reasonably sized while breaking the
# fixed "always N" / "2-2-1" pattern QA flagged.
_VARY_BAND = 6


def _varied_target(vary_key: str, tid: str, min_per_tid: int, max_per_tid: int) -> int:
    """Deterministic per-(schedule, insured) row target in a band above the min.

    Stable for a given (vary_key, tid) so runs are reproducible; different schedules
    and different insureds get different counts, so no fixed pattern emerges.
    """
    band = min(max_per_tid, min_per_tid + _VARY_BAND) - min_per_tid
    if band <= 0:
        return min_per_tid
    h = int(hashlib.md5(f"{vary_key}|{tid}".encode()).hexdigest(), 16)
    return min_per_tid + (h % (band + 1))


def ensure_child_row_multiplicity(
    rows: list[dict[str, Any]],
    *,
    min_per_tid: int,
    max_per_tid: int,
    unique_frags: Sequence[str] = (),
    skip_predicate=None,
    vary_key: str | None = None,
    all_test_ids: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Deterministically guarantee each Test ID carries a multi-row child schedule.

    The recurring QA defect across every child schedule (equipment, locations,
    drivers, vehicles, trailers, commodities, loss payees, loss history) is that
    the LLM emits a *single* row per insured even though the ruleset allows many.
    Child-row **counts are a deterministic field** (CLAUDE.md §2: scenario counts
    live on the deterministic side of the boundary), so the engine *produces* the
    multiplicity instead of hoping the prompt is honoured.

    For every Test ID that already has at least one row, the group is:

    * **expanded** up to a per-group target (``min_per_tid`` by default) by cloning
      its existing rows round-robin — re-sequencing the literal ``#`` column and
      appending a ``-N`` suffix to any ``unique_frags`` identifier (Serial Number,
      VIN, Loan Number, …) so cloned rows stay distinct; and
    * **capped** at ``max_per_tid`` (the ruleset upper bound).

    When ``vary_key`` is supplied the target is *varied per insured* in a band above
    ``min_per_tid`` (deterministically, keyed by ``vary_key`` + Test ID) so counts
    do not follow a fixed pattern across policies (CARGO-008 / APD-014). Without it,
    the target is exactly ``min_per_tid`` (stable minimum-guarantee behaviour).

    ``skip_predicate(first_row)`` suppresses expansion for a group (used for
    "no losses" loss-history rows, which must stay a single blank row). Groups are
    emitted in first-seen Test ID order; rows with no Test ID are passed through.

    ``all_test_ids`` is the authoritative Test-ID roster (from Policy Info). On a
    large request the LLM anchors to the first few Test IDs and stops, so most
    insureds receive *zero* child rows (the HO "6 loss payees for 50 test cases"
    defect). Because child-row multiplicity is a **deterministic** field
    (CLAUDE.md §2), the roster — not the LLM output — decides which insureds get a
    schedule: any roster Test ID with no produced rows is seeded from a donor row
    (round-robin over the produced rows, stamped with the missing Test ID) so the
    expand/cap loop below fills it like any other group. Omit it to keep the prior
    "expand only what the LLM emitted" behaviour.
    """
    if not rows:
        return rows
    order: list[str] = []
    groups: dict[str, list[dict[str, Any]]] = {}
    for r in rows:
        t = tid_value(r)
        groups.setdefault(t, [])
        if t not in order:
            order.append(t)
        groups[t].append(r)

    tid_key_all = _tid_key(rows[0])

    if all_test_ids:
        # The Policy roster OWNS the child Test IDs (deterministic cross-sheet
        # key). When asked for a scaled row count the LLM invents extra Test IDs
        # (TS-021…TS-060 for a 20-insured request); those orphan rows must NOT
        # appear in output — every child row must join a real policy — but their
        # content is reused as donor variety so nothing is wasted. Output is
        # therefore restricted to EXACTLY the roster, in roster order.
        roster = list(all_test_ids)
        roster_set = set(roster)
        orphan_rows = [r for r in rows if tid_value(r) not in roster_set]
        roster_groups = {t: list(groups.get(t, [])) for t in roster}
        # Seed empty roster groups from a donor (prefer an orphan's real content).
        donors = orphan_rows or list(rows)
        di = 0
        for t in roster:
            if roster_groups[t]:
                continue
            seed = dict(donors[di % len(donors)]) if donors else {}
            di += 1
            if tid_key_all:
                seed[tid_key_all] = t
            # Keep unique identifiers (Loan Number, Serial/VIN, …) distinct across
            # insureds — the donor's value belongs to another Test ID.
            for uf in unique_frags:
                col = find_col(seed, uf)
                if col and str(seed.get(col) or "").strip():
                    seed[col] = f"{seed[col]}-{t}"
            roster_groups[t] = [seed]
        groups = roster_groups
        order = roster

    # Donor pool for expansion variety: every produced row except single-blank
    # "skip" rows (e.g. a no-loss loss-history row). When a group needs MORE rows
    # than the LLM emitted for it, the added rows are drawn from this pool — real
    # rows authored for OTHER insureds, restamped to this group's Test ID — so an
    # expanded schedule is genuinely varied instead of byte-identical clones of a
    # single row (APD-019/021/023/025 "generating duplicate data"). Orphan rows
    # stay in the pool as donor variety even though they are not output on their
    # own Test ID.
    donor_pool = [r for r in rows if not (skip_predicate and skip_predicate(r))]

    out: list[dict[str, Any]] = []
    for gi, t in enumerate(order):
        grp = list(groups[t][:max_per_tid])
        seeds = list(grp)
        if seeds and skip_predicate and skip_predicate(seeds[0]):
            # A "skip" insured (no losses / no loss payees) must carry exactly ONE
            # row. The LLM sometimes emits several identical blank rows for it —
            # collapse them so the schedule shows a single no-detail row, not
            # duplicates (DW/HO no-loss blank-row duplication).
            grp = grp[:1]
        elif seeds:
            target = (_varied_target(vary_key, t, min_per_tid, max_per_tid)
                      if vary_key is not None else min_per_tid)
            # Prefer varied donors from the full pool; fall back to the group's own
            # rows if the pool is empty. Offset the start per group so different
            # insureds pull different donors (avoids everyone cloning row 0).
            sources = donor_pool or seeds
            # Track the content already in the group so an added clone never
            # duplicates an existing row (schedules with no unique_frags / no "#"
            # column — e.g. Commodities — would otherwise repeat identical rows).
            seen = {tuple(sorted(r.items(), key=lambda kv: kv[0])) for r in grp}
            k = gi
            while len(grp) < target:
                tried = 0
                while tried < len(sources):
                    clone = dict(sources[k % len(sources)])
                    if tid_key_all:
                        clone[tid_key_all] = t
                    suffix = len(grp) + 1
                    for uf in unique_frags:
                        col = find_col(clone, uf)
                        if col and str(clone.get(col) or "").strip():
                            clone[col] = f"{clone[col]}-{suffix}"
                    sig = tuple(sorted(clone.items(), key=lambda kv: kv[0]))
                    k += 1
                    tried += 1
                    if sig not in seen:
                        break
                # If every donor duplicates an existing row, accept the last one
                # (target multiplicity wins over a marginal duplicate).
                seen.add(sig)
                grp.append(clone)
        ncol = _number_col(grp[0]) if grp else None
        if ncol:
            for i, r in enumerate(grp, start=1):
                r[ncol] = i
        out.extend(grp)
    return out


def format_phone_fax(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Snap every non-blank phone / fax / contact-number cell to U.S. format.

    Covers the IM/Cargo/APD phone+fax defects (DF-IM-010/011/014, CARGO-001/002,
    APD-001/002). Idempotent: already-formatted ``(XXX) XXX-XXXX`` is preserved.
    Blank cells are left blank (a blank phone is a separate, non-format concern).
    """
    for row in rows:
        for key in list(row.keys()):
            kl = key.lower()
            if ("phone" in kl or "fax" in kl or "contact number" in kl):
                if not _is_blank(row.get(key)):
                    row[key] = format_us_phone(row.get(key))
    return rows


def _binding_value(row: dict[str, Any]) -> str | None:
    """The row's Binding/Rating state value, so address states can avoid it."""
    for key in row:
        if is_binding_state_field(key):
            v = str(row.get(key) or "").strip().upper()
            return v or None
    return None


# Address-block column classification. A physical-address block is the set of
# City / State / ZIP columns that share a name prefix ("Address – City/State/Zip",
# "Mailing City/State/Zip", "Agency Address State", …). The block is keyed by that
# prefix so its sibling columns can be assigned a *consistent* (state, city, zip)
# triple together — never a valid ZIP under a mismatched state.
_KIND_STRIP = {"zip", "zipcode", "postal", "code", "city", "state", "st"}


_STATE_WORD = re.compile(r"\bstate\b")
_CITY_WORD = re.compile(r"\bcity\b")


def _addr_kind(col: str) -> str | None:
    """Classify an address column as ``state`` / ``city`` / ``zip`` (or None).

    The Binding/Rating dropdown is never an address column (it stays untouched).
    ``state`` / ``city`` match on **whole words** so value columns that merely
    contain the letters — "Stated Amount or ACV", "Real Estate", "Interstate",
    "Velocity" — are never mistaken for an address column and clobbered with a
    state code / city name.
    """
    if is_binding_state_field(col):
        return None
    kl = col.strip().lower()
    if "zip" in kl or "postal" in kl:
        return "zip"
    if _CITY_WORD.search(kl):
        return "city"
    if _STATE_WORD.search(kl) or kl == "st":
        return "state"
    return None


def _block_key(col: str) -> str:
    """The block prefix for an address column (kind words + digits stripped)."""
    toks = re.split(r"[^a-z0-9]+", col.strip().lower())
    keep = [t for t in toks if t and not t.isdigit() and t not in _KIND_STRIP]
    return " ".join(keep)


def _address_blocks(keys: Sequence[str]) -> dict[str, dict[str, str]]:
    """Group header keys into address blocks: prefix -> {kind: column}."""
    blocks: dict[str, dict[str, str]] = {}
    for k in keys:
        kind = _addr_kind(k)
        if not kind:
            continue
        blocks.setdefault(_block_key(k), {}).setdefault(kind, k)
    return blocks


def spread_address_states(
    rows: list[dict[str, Any]],
    *,
    state_selection: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Fan physical-address blocks across the US-state pool, keeping City/State/ZIP
    consistent.

    Address columns are grouped into blocks by their shared name prefix. For every
    block that carries a State column (but is **not** the restricted Binding/Rating
    dropdown), each *populated* cell is reassigned a state from the pool by
    :func:`spread_pick`, phased per block so sibling address blocks in one row show
    different states (real combinations). When the block also has City and/or ZIP
    columns, those are set from :data:`~app.rulebook.geo.STATE_GEO` so the whole
    triple corresponds — the ZIP always belongs to the assigned state and city.

    The binding/rating column itself is never touched, and blank cells are
    preserved (so handler dependency-blanking such as "Mailing State blank when not
    different" survives — a blank State skips its whole block for that row).

    Frontend overrides win: when ``state_selection`` is supplied (UI state filter),
    the pool is exactly those states, so address states stay inside the user's
    selection instead of the full list.

    Closes the state cluster (DF-IM-012/015/016/018, DEF-011/012/013/014/016,
    HO-015, CARGO-003, APD-003) and the City/State/ZIP correspondence class
    (DEF-024 / HO-023 / CARGO-007 / WH-004 / APD-015 and the IM generic address
    defect).
    """
    pool: tuple[str, ...] | list[str]
    if state_selection:
        pool = [str(s).strip().upper() for s in state_selection if str(s).strip()]
        if not pool:
            pool = list(US_STATES)
    else:
        pool = US_STATES

    blocks = _address_blocks(rows[0].keys() if rows else [])
    state_blocks = {bk: b for bk, b in blocks.items() if "state" in b}
    if not state_blocks:
        return rows

    for bk, cols in state_blocks.items():
        state_col = cols["state"]
        city_col = cols.get("city")
        zip_col = cols.get("zip")
        seed = column_seed(state_col)
        for i, row in enumerate(rows):
            if _is_blank(row.get(state_col)):
                continue  # preserve intentional blanks (dependency rules)
            avoid = _binding_value(row) if len(pool) > 1 else None
            state = spread_pick(i, pool, seed=seed, avoid=avoid)
            row[state_col] = state
            if not (city_col or zip_col):
                continue
            geo = geo_for_state(state)
            if not geo:
                continue
            city, zipcode = spread_pick(i, geo, seed=seed + 1)
            if city_col:
                row[city_col] = city
            if zip_col:
                row[zip_col] = zipcode
    return rows


def enforce_address_consistency(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Snap each address block's City/ZIP to correspond to its **own State**,
    without changing the state.

    This is the consistency half of :func:`spread_address_states` with the state
    *spread* removed: the state the row already carries is authoritative, and only
    City/ZIP are repaired to belong to it (from :data:`~app.rulebook.geo.STATE_GEO`).
    It closes the City/State/ZIP correspondence class (HO-023 / WH-004 and kin) for
    the LOBs that do **not** run the SPG state-spread pass — GENERIC and any uploaded
    template — so a valid-but-mismatched ZIP can never survive on any LOB. PAP is
    excluded by the caller (it carries real Census-verified addresses).

    Idempotent: a cell that already forms a valid geo triple for its state is left
    untouched, so re-running on already-consistent SPG output is a no-op, and the
    intentional blanks handlers leave (dependency rules) are preserved.
    """
    if not rows:
        return rows
    for cols in _address_blocks(rows[0].keys()).values():
        state_col = cols.get("state")
        city_col = cols.get("city")
        zip_col = cols.get("zip")
        if not state_col or not (city_col or zip_col):
            continue
        seed = column_seed(state_col)
        for i, row in enumerate(rows):
            state = str(row.get(state_col) or "").strip().upper()
            if not state:
                continue  # blank state -> nothing to correspond to (preserve blank)
            geo = geo_for_state(state)
            if not geo:
                continue
            cur_city = str(row.get(city_col) or "").strip() if city_col else ""
            cur_zip = str(row.get(zip_col) or "").strip() if zip_col else ""
            already_ok = (
                (not city_col or any(c == cur_city for c, _ in geo))
                and (not zip_col or any(z == cur_zip for _, z in geo))
                and (not (city_col and zip_col) or (cur_city, cur_zip) in {p for p in geo})
            )
            if already_ok:
                continue
            city, zipcode = spread_pick(i, geo, seed=seed + 1)
            if city_col:
                row[city_col] = city
            if zip_col:
                row[zip_col] = zipcode
    return rows


def enforce_variety_fields(
    rows: list[dict[str, Any]],
    *,
    state_selection: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Run the universal SPG passes (phone/fax format + address-state spread).

    Called once per generated sheet (excluding deterministic summary sheets) for
    the SPG LOB family. Order is irrelevant between the two passes; both are
    idempotent so re-running on already-clean data is a no-op.
    """
    if not rows:
        return rows
    format_phone_fax(rows)
    spread_address_states(rows, state_selection=state_selection)
    return rows
