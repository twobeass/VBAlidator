"""Negative-control: every file under `tests/samples/valid_code/` must parse
clean (zero analyzer errors). New valid samples are picked up automatically.

A sample may declare a tolerance budget by including a single comment of the
form ``' EXPECTED_ERRORS: <N>`` in its first 5 lines. This is intended for
fixtures that intentionally exercise constructs the analyzer is still lenient
about (e.g. RaiseEvent / GoSub coverage placeholders). Drive these to 0 over
time.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest


VALID_DIR = Path(__file__).resolve().parent / "samples" / "valid_code"
EXPECTED_ERRORS_RE = re.compile(r"'\s*EXPECTED_ERRORS:\s*(\d+)", re.IGNORECASE)


def _collect() -> list[Path]:
    if not VALID_DIR.is_dir():
        return []
    return sorted(
        p for p in VALID_DIR.rglob("*")
        if p.suffix.lower() in (".bas", ".cls", ".frm")
    )


def _read_expected_errors(path: Path) -> int:
    try:
        with open(path, "r", encoding="latin-1") as f:
            head = "".join(next(f, "") for _ in range(5))
    except OSError:
        return 0
    m = EXPECTED_ERRORS_RE.search(head)
    return int(m.group(1)) if m else 0


VALID_SAMPLES = _collect()


@pytest.mark.parametrize("sample_path", VALID_SAMPLES, ids=[p.name for p in VALID_SAMPLES])
def test_valid_sample_has_no_errors(sample_path, run_files):
    result = run_files([sample_path])
    expected = _read_expected_errors(sample_path)
    # Severity 'warning' / 'info' findings (Option Explicit, etc.) don't
    # count as compile errors for the valid-sample gate.
    hard_errors = [
        e for e in result.errors
        if e.get("severity", "error") == "error"
    ]

    assert len(hard_errors) <= expected, (
        f"Expected ≤{expected} analyzer errors for valid sample {sample_path.name}, "
        f"got {len(hard_errors)}: {[e.get('message','') for e in hard_errors]!r}"
    )
