"""Scenario parser: special-instructions text -> per-insured count specs.

A "scenario" is a user-declared per-insured child-row multiplicity, e.g.
"1 insured: 20 vehicles, 8 class codes" → one insured with 20 vehicle rows and
8 distinct class-of-business rows. Each ``N insured(s)`` phrase introduces N
insureds sharing the counts that follow it; multiple phrases produce multiple
ordered specs.

Counts above a per-rule maximum are capped (not errored) and the adjustment is
recorded for surfacing (R11). Text with no ``insured`` phrasing yields no specs,
signalling "no scenario" so the existing checkpoint behavior runs unchanged.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# Per-rule maxima, mirroring the RRG rulebook (Rule 36 vehicles 1–20,
# Rule 37 locations 1–20, Rule 38 class of business 1–8).
COUNT_MAXIMA: dict[str, int] = {
    "vehicles": 20,
    "locations": 20,
    "class_of_business": 8,
}

# unit phrase -> canonical count key. Order matters: longer/more-specific
# phrases first so "class of business" is not shadowed by a "class" prefix.
_UNIT_PATTERNS: list[tuple[str, str]] = [
    (r"class\s+of\s+business(?:es)?", "class_of_business"),
    (r"class\s+codes?", "class_of_business"),
    (r"\bcob\b", "class_of_business"),
    (r"vehicles?", "vehicles"),
    (r"locations?", "locations"),
]

_COUNT_RE = re.compile(
    r"(\d+)\s*(" + "|".join(p for p, _ in _UNIT_PATTERNS) + r")",
    re.IGNORECASE,
)

_INSURED_RE = re.compile(r"(\d+)\s*insureds?\b", re.IGNORECASE)


def _canonical_unit(raw: str) -> str:
    raw_l = raw.lower()
    for pattern, key in _UNIT_PATTERNS:
        if re.fullmatch(pattern, raw_l, re.IGNORECASE):
            return key
    return raw_l


@dataclass
class InsuredSpec:
    """Per-insured child-row counts (only the keys the user specified)."""

    counts: dict[str, int] = field(default_factory=dict)

    def get(self, key: str, default=None):
        return self.counts.get(key, default)

    def __bool__(self) -> bool:
        return bool(self.counts)


@dataclass
class ScenarioParseResult:
    """Ordered insured specs plus any cap adjustments surfaced to the user."""

    specs: list[InsuredSpec] = field(default_factory=list)
    adjustments: list[str] = field(default_factory=list)

    def __bool__(self) -> bool:
        return bool(self.specs)

    def __len__(self) -> int:
        return len(self.specs)


def _parse_counts(body: str, adjustments: list[str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for num_str, unit_raw in _COUNT_RE.findall(body):
        key = _canonical_unit(unit_raw)
        requested = int(num_str)
        max_count = COUNT_MAXIMA.get(key)
        value = requested
        if max_count is not None and requested > max_count:
            value = max_count
            adjustments.append(
                f"Requested {requested} {key.replace('_', ' ')} exceeds max "
                f"{max_count}; capped to {max_count}."
            )
        counts[key] = value
    return counts


def parse_scenarios(text: str | None) -> ScenarioParseResult:
    """Parse special-instructions text into an ordered list of insured specs."""
    result = ScenarioParseResult()
    if not text:
        return result

    anchors = list(_INSURED_RE.finditer(text))
    if not anchors:
        return result  # no scenario phrasing

    for i, match in enumerate(anchors):
        start = match.end()
        end = anchors[i + 1].start() if i + 1 < len(anchors) else len(text)
        body = text[start:end]
        counts = _parse_counts(body, result.adjustments)
        if not counts:
            continue
        n_insureds = max(1, int(match.group(1)))
        for _ in range(n_insureds):
            result.specs.append(InsuredSpec(counts=dict(counts)))

    return result
