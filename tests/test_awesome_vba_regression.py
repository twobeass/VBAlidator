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
    "JSONBag": (1, "baseline phase-3+P2.6 fixup; ComctlLib chains now permissive"),
    "VBA-MemoryTools-master": (17, "baseline phase-3+P2.6 fixup"),
    "VbTrickTimer-master": (5, "baseline phase-3; preprocessor case-fix + numeric type suffix"),
    "stdVBA-master": (180, "baseline iter-4-wave-4; +On…GoTo +CallByName +ReDim member chain"),
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
    # Style-level findings (warning / info) are not part of the baseline.
    # Only hard compile errors gate regression of the analyzer surface.
    hard_errors = [
        e for e in result.errors
        if e.get("severity", "error") == "error"
    ]
    ceiling, reason = BASELINE.get(project.name, (0, "no baseline; expected clean"))

    assert len(hard_errors) <= ceiling, (
        f"Regression in {project.name}: {len(hard_errors)} errors exceeds "
        f"baseline {ceiling} ({reason}). Sample messages: "
        f"{[e.get('message','') for e in hard_errors[:5]]!r}"
    )
