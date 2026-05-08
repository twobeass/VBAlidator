"""Regression wall: real-world VBA libraries under `tests/awesome_vba/`
must analyze with **zero new false-positives**.

This is the safety net for every future change to the analyzer. Each
sub-project becomes one test case; new false positives surface immediately.

If an upstream library legitimately needs an external symbol the standard
model does not know about, add it to the BASELINE below with a short
justification (count of currently-tolerated issues for that project).
The goal is to drive every entry to 0 over time.
"""
from __future__ import annotations

from pathlib import Path

import pytest


AWESOME_DIR = Path(__file__).resolve().parent / "awesome_vba"

# Project-name -> (max tolerated errors, reason). Drive these to 0 over time.
# A single source of truth so PRs surface unexpected regressions clearly.
# Initial baseline measured in PR #1 (Phase 0), each value is the current
# ceiling — never a target. Lower me as the analyzer improves.
BASELINE: dict[str, tuple[int, str]] = {
    "JSONBag": (12, "baseline phase-1; ComctlLib types not in std model"),
    "VBA-MemoryTools-master": (18, "baseline phase-2; +1 VBA231 Const-from-variable in LibMemory"),
    "VbTrickTimer-master": (54, "baseline phase-1; many API declares + WithEvents"),
    "stdVBA-master": (342, "baseline phase-2; +1 VBA201 (one cross-procedure label jump)"),
}


def _project_dirs() -> list[Path]:
    if not AWESOME_DIR.is_dir():
        return []
    return sorted(p for p in AWESOME_DIR.iterdir() if p.is_dir())


def _vba_files(project: Path) -> list[Path]:
    return sorted(
        p for p in project.rglob("*")
        if p.suffix.lower() in (".bas", ".cls", ".frm")
    )


PROJECTS = _project_dirs()


@pytest.mark.parametrize("project", PROJECTS, ids=[p.name for p in PROJECTS])
def test_awesome_vba_within_baseline(project, run_files):
    files = _vba_files(project)
    if not files:
        pytest.skip(f"No VBA files in {project}")

    result = run_files(files)
    ceiling, reason = BASELINE.get(project.name, (0, "no baseline; expected clean"))

    assert len(result.errors) <= ceiling, (
        f"Regression in {project.name}: {len(result.errors)} errors exceeds "
        f"baseline {ceiling} ({reason}). Sample messages: {result.messages[:5]!r}"
    )
