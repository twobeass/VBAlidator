"""Regression wall: real-world VBA libraries under `tests/awesome_vba/`
must analyze with **zero new false-positives**.

This is the safety net for every future change to the analyzer. Each
sub-project becomes one test case; new false positives surface immediately.

If an upstream library legitimately needs an external symbol the standard
model does not know about, add it to the BASELINE below with a short
justification (count of currently-tolerated issues for that project).
The goal is to drive every entry to 0 over time.

Host mapping
------------
The `HOSTS` dict picks which bundled host model (excel/word/access/visio)
is layered on top of std_model for each project. Set `None` for genuinely
host-agnostic libraries (pure VBA + Win32). Picking the right host is the
single biggest lever — `--host excel` alone dropped stdVBA from 180 to 141.
"""
from __future__ import annotations

from pathlib import Path

import pytest


AWESOME_DIR = Path(__file__).resolve().parent / "awesome_vba"

# Project-name -> bundled host model layered on std_model.
# None means "host-agnostic VBA code"; otherwise must match a key of
# the CLI's --host choices (excel/word/access/outlook/visio).
HOSTS: dict[str, str | None] = {
    "JSONBag": None,
    "VBA-MemoryTools-master": None,
    "VbTrickTimer-master": None,
    "stdVBA-master": "excel",
}

# Project-name -> (max tolerated errors, reason). Drive these to 0 over time.
# A single source of truth so PRs surface unexpected regressions clearly.
# Each value is the current *ceiling*, never a target — lower me as the
# analyzer improves. Reasons explain which class of remaining issue each
# baseline still reflects so future readers know what's a real bug vs.
# a known coverage gap.
BASELINE: dict[str, tuple[int, str]] = {
    "JSONBag": (
        0,
        "Clean — fixed by shipping the MSComCtl host model and auto-layering "
        "it whenever a `.frm` in the scan set references the library.",
    ),
    "VBA-MemoryTools-master": (
        2,
        "Down from 17 by suppressing Sub-style implicit-call interpretation "
        "inside expression sub-tokens (the `Round(Timer - t, 3)` quirk). "
        "The 2 remaining are genuine upstream const-initialiser-references-"
        "variable issues in LibMemory.bas that the strict-mode rule flags.",
    ),
    "VbTrickTimer-master": (
        0,
        "Clean — fixed by treating `As Any` / `As Any()` Declare params as "
        "universally-compatible sentinels, plus letting `arr()` with empty "
        "parens keep the array type (VBA's explicit pass-whole-array form, "
        "not an indexed element access).",
    ),
    "stdVBA-master": (
        16,
        "Down from 180 via --host excel + library namespaces + stdError "
        "fixture + MSForms 2.0 host model + Enum<->Long ByRef compat + "
        "suppressing Sub-style implicit-call + Array/Choose/Switch as "
        "ParamArray + `error` as identifier + lexer line-continuation "
        "trailing whitespace. Remaining ≈16 are deep analyzer cases "
        "(member-on-unknown-type, Dim-inside-loop-scope, default-property "
        "Item) earmarked for vbatest Iter-5.",
    ),
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

    host = HOSTS.get(project.name)
    result = run_files(files, host=host)
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
