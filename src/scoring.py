"""Confidence-Score: a single 0–100 number summarising how likely it is
that the analysed VBA source compiles.

The score is intentionally simple and deterministic so CI gates can
threshold on it (e.g. `score >= 90` to allow merging an AI-generated
patch). Severity weights are tuned to favour precision: each hard
compile error subtracts more than a style warning so that a single
real error always drops the score below the 90 % gate.

Components
----------
- Severity penalty: errors -20, warnings -3, info -1.
- Coverage cap: when the analyzed code references library types the
  loaded model does not know about, we cap at 90 to flag uncertainty.
- Floor: the score never goes below 0.
- Ceiling: the score never exceeds 100.

`compile_safe` is True iff there are zero severity=error findings.
"""
from __future__ import annotations

from typing import Iterable


SEVERITY_WEIGHTS = {
    "error": 20,
    "warning": 3,
    "info": 1,
}


def compute_score(issues: Iterable[dict], coverage_uncertain: bool = False) -> tuple[int, dict]:
    """Return (score, breakdown). Pure function."""
    counts = {"error": 0, "warning": 0, "info": 0}
    penalty = 0
    for i in issues:
        sev = i.get("severity", "error")
        if sev not in counts:
            counts[sev] = 0
        counts[sev] += 1
        penalty += SEVERITY_WEIGHTS.get(sev, SEVERITY_WEIGHTS["error"])

    score = 100 - penalty
    if coverage_uncertain and score > 90:
        score = 90  # cap until the user supplies a complete model
    score = max(0, min(100, score))

    breakdown = {
        "starting": 100,
        "penalty_total": penalty,
        "by_severity": counts,
        "weights": dict(SEVERITY_WEIGHTS),
        "coverage_uncertain": coverage_uncertain,
        "final": score,
    }
    return score, breakdown


def is_compile_safe(issues: Iterable[dict]) -> bool:
    """A run is `compile_safe` only when zero severity=error findings exist.
    Warnings and info do not block."""
    for i in issues:
        if i.get("severity", "error") == "error":
            return False
    return True
