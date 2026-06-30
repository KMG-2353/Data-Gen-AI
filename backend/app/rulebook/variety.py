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

from typing import Any, Sequence

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


def ensure_child_row_multiplicity(
    rows: list[dict[str, Any]],
    *,
    min_per_tid: int,
    max_per_tid: int,
    unique_frags: Sequence[str] = (),
    skip_predicate=None,
) -> list[dict[str, Any]]:
    """Deterministically guarantee each Test ID carries a multi-row child schedule.

    The recurring QA defect across every child schedule (equipment, locations,
    drivers, vehicles, trailers, commodities, loss payees, loss history) is that
    the LLM emits a *single* row per insured even though the ruleset allows many.
    Child-row **counts are a deterministic field** (CLAUDE.md §2: scenario counts
    live on the deterministic side of the boundary), so the engine *produces* the
    multiplicity instead of hoping the prompt is honoured.

    For every Test ID that already has at least one row, the group is:

    * **expanded** up to ``min_per_tid`` by cloning its existing rows round-robin —
      re-sequencing the literal ``#`` column and appending a ``-N`` suffix to any
      ``unique_frags`` identifier (Serial Number, VIN, Loan Number, …) so cloned
      rows stay distinct; and
    * **capped** at ``max_per_tid`` (the ruleset upper bound).

    ``skip_predicate(first_row)`` suppresses expansion for a group (used for
    "no losses" loss-history rows, which must stay a single blank row). Groups are
    emitted in first-seen Test ID order; rows with no Test ID are passed through.
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

    out: list[dict[str, Any]] = []
    for t in order:
        grp = list(groups[t][:max_per_tid])
        seeds = list(grp)
        if seeds and not (skip_predicate and skip_predicate(seeds[0])):
            k = 0
            while len(grp) < min_per_tid:
                clone = dict(seeds[k % len(seeds)])
                suffix = len(grp) + 1
                for uf in unique_frags:
                    col = find_col(clone, uf)
                    if col and str(clone.get(col) or "").strip():
                        clone[col] = f"{clone[col]}-{suffix}"
                grp.append(clone)
                k += 1
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


def spread_address_states(
    rows: list[dict[str, Any]],
    *,
    state_selection: Sequence[str] | None = None,
) -> list[dict[str, Any]]:
    """Fan physical-address state columns across the full US-state pool.

    For every column that is a state field but **not** the restricted
    Binding/Rating dropdown, each *populated* cell is reassigned a state from the
    pool by :func:`spread_pick`, phased per column so sibling address fields in
    one row show different states (real combinations). The binding/rating column
    itself is never touched, and blank cells are preserved (so handler
    dependency-blanking such as "Mailing State blank when not different" survives).

    Frontend overrides win: when ``state_selection`` is supplied (UI state
    filter), the pool is exactly those states, so address states stay inside the
    user's selection instead of the full list.

    Closes the state cluster: DF-IM-012/015/016/018, DEF-011/012/013/014/016,
    HO-015, CARGO-003, APD-003.
    """
    pool: tuple[str, ...] | list[str]
    if state_selection:
        pool = [str(s).strip().upper() for s in state_selection if str(s).strip()]
        if not pool:
            pool = list(US_STATES)
    else:
        pool = US_STATES

    # Address-state columns = any "*state*" column (or a bare "St" abbreviation,
    # as on the WH Locations sheet) that is not the Binding/Rating dropdown.
    def _is_addr_state_col(k: str) -> bool:
        kl = k.strip().lower()
        is_state = ("state" in kl) or kl == "st"
        return is_state and not is_binding_state_field(k)

    state_cols = [k for k in (rows[0].keys() if rows else []) if _is_addr_state_col(k)]
    if not state_cols:
        return rows

    for col in state_cols:
        seed = column_seed(col)
        for i, row in enumerate(rows):
            val = row.get(col)
            if _is_blank(val):
                continue  # preserve intentional blanks (dependency rules)
            avoid = _binding_value(row) if len(pool) > 1 else None
            row[col] = spread_pick(i, pool, seed=seed, avoid=avoid)
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
