"""Tests for Phase 4.5 (round-trip verification) and 4.6 (rule docs)."""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

from src import roundtrip
from src.api import precheck_source
from src.rules import all_rules, get_rule, known_rule_ids
from src.scoring import compute_score, is_compile_safe


ROOT = Path(__file__).resolve().parent.parent


# ------------------------------------------------------------------ Rules


def test_every_known_rule_has_complete_metadata():
    for rule in all_rules():
        assert rule.rule_id, "rule_id must be non-empty"
        assert rule.title, f"{rule.rule_id} must have a title"
        assert rule.severity in {"error", "warning", "info", "compile_verified"}, (
            f"{rule.rule_id} severity '{rule.severity}' is not recognised"
        )
        assert rule.category, f"{rule.rule_id} must have a category"
        assert rule.description, f"{rule.rule_id} must have a description"
        # Phase is optional but present for everything we ship.
        assert rule.phase, f"{rule.rule_id} must have a phase tag"


def test_rule_ids_are_unique():
    ids = [r.rule_id for r in all_rules()]
    assert len(ids) == len(set(ids))


def test_get_rule_returns_known_rule():
    assert get_rule("VBA001").title == "Undefined identifier"
    assert get_rule("nope") is None


def test_all_phase_2_rule_ids_are_registered():
    """Spot-check: every rule_id we emit from the analyzer must have an
    entry in the registry. Bumping a new rule without registering it
    would silently degrade the rule catalogue."""
    expected = {
        "VBA101", "VBA102", "VBA103", "VBA104", "VBA105", "VBA106",
        "VBA201",
        "VBA210", "VBA211",
        "VBA221", "VBA222", "VBA223", "VBA224",
        "VBA230", "VBA231",
        "VBA240",
        "VBA250",
        "VBA300",
        "VBA310",
        "VBA320",
        "VBA330",
        "VBA340", "VBA341",
        "VBA_LEX001", "VBA_LEX002",
    }
    missing = expected - known_rule_ids()
    assert not missing, f"missing rule registrations: {sorted(missing)}"


# ------------------------------------------------------------------ Score


def test_compile_verified_blocks_compile_safe():
    issue = {"severity": "compile_verified", "message": "VBE refused"}
    assert not is_compile_safe([issue])


def test_compile_verified_carries_higher_penalty_than_error():
    score_e, _ = compute_score([{"severity": "error"}])
    score_v, _ = compute_score([{"severity": "compile_verified"}])
    assert score_v < score_e


# ------------------------------------------------------------------ Roundtrip


def test_roundtrip_unavailable_on_non_windows():
    """On Linux / macOS the round-trip module must report unavailability
    rather than crashing or silently passing."""
    if sys.platform == "win32":
        pytest.skip("only checks the non-Windows fallback path")
    assert not roundtrip.is_available()
    reason = roundtrip.availability_reason()
    assert "Windows" in reason or "win" in reason.lower()


def test_verify_compile_raises_unavailable_off_platform():
    if sys.platform == "win32":
        pytest.skip("only checks the non-Windows fallback path")
    with pytest.raises(roundtrip.RoundtripUnavailable):
        roundtrip.verify_compile("Sub S(): End Sub")


def test_precheck_with_roundtrip_off_platform_emits_info():
    """precheck(..., roundtrip=True) should not raise on non-Windows; it
    appends an info-level entry explaining why the round-trip is skipped."""
    if sys.platform == "win32":
        pytest.skip("only checks the non-Windows fallback path")
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(): End Sub\n"
    )
    # We can't use the test fixture run_source because that doesn't
    # propagate roundtrip=True. Use the public API directly.
    from src.api import precheck

    # Write to a temp file because precheck infers the path-ish heuristic.
    import tempfile
    with tempfile.NamedTemporaryFile("w", suffix=".bas", delete=False) as fh:
        fh.write(code)
        path = fh.name
    try:
        result = precheck(path, roundtrip=True)
    finally:
        import os
        os.unlink(path)

    assert any(i.get("rule_id") == "VBA_RT000" for i in result.issues), (
        f"Off-platform round-trip should emit an info VBA_RT000. "
        f"Issues: {result.issues!r}"
    )
    # And the run is still compile_safe (info doesn't block).
    assert result.compile_safe


# ------------------------------------------------------------------ Doc generator


_RULE_DIR = ROOT / "docs" / "rules"


def test_rule_docs_generator_is_idempotent(tmp_path):
    """Running the generator twice must not change any file (no diff)."""
    script = ROOT / "tools" / "generate_rule_docs.py"
    # First pass: ensure docs are up-to-date.
    subprocess.run([sys.executable, str(script)], check=True, cwd=ROOT)
    snapshot = {p.name: p.read_text() for p in _RULE_DIR.iterdir() if p.is_file()}
    # Second pass: must be byte-identical.
    subprocess.run([sys.executable, str(script)], check=True, cwd=ROOT)
    after = {p.name: p.read_text() for p in _RULE_DIR.iterdir() if p.is_file()}
    assert snapshot == after, "generator is not idempotent"


def test_every_registered_rule_has_a_doc_page():
    page_files = {p.stem for p in _RULE_DIR.iterdir() if p.suffix == ".md" and p.stem != "index"}
    assert page_files == known_rule_ids(), (
        f"docs/rules/ is out of sync with the registry. "
        f"Run `python tools/generate_rule_docs.py`. "
        f"Diff: registered_only={known_rule_ids() - page_files} "
        f"docs_only={page_files - known_rule_ids()}"
    )


def test_index_lists_every_rule():
    text = (_RULE_DIR / "index.md").read_text()
    for rule_id in known_rule_ids():
        assert re.search(rf"`{rule_id}`", text), f"{rule_id} missing from index.md"


def test_each_rule_page_has_required_sections():
    """Every per-rule page must include Description, fail/ok examples and
    a fix hint. This guards against shipping new rules with empty docs."""
    for rule in all_rules():
        page = (_RULE_DIR / f"{rule.rule_id}.md").read_text()
        assert "## Description" in page
        assert "## How to fix" in page
        # Most pages have both examples; allow empty when description-only.
        if rule.fail_example:
            assert "## Failing example" in page
        if rule.ok_example:
            assert "## Compliant example" in page


# ------------------------------------------------------------------ End-to-end


def test_precheck_decorates_legacy_findings_with_phase_4_5_severity_aware_score():
    """A code sample with 1 error + 1 warning lands at the well-known
    score `100 - 20 - 3 = 77`, demonstrating that the scoring function
    weights warnings differently from errors."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Sub S()\n"
        "    typo = 1\n"   # undefined identifier (error)
        "End Sub\n"
    )
    result = precheck_source(code)
    # Errors: 1 (VBA001)  Warnings: 1 (VBA320)  → 100 - 20 - 3 = 77
    assert result.score == 77, (
        f"Expected score 77 (1 error -20, 1 warning -3), got {result.score}. "
        f"Issues: {result.issues!r}"
    )
