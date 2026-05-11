"""Tests for Phase 4.5 (round-trip verification) and 4.6 (rule docs)."""
from __future__ import annotations

import os
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
    """Running the generator twice must not change any file (no diff).

    Reads use explicit UTF-8 — the catalogue contains emoji severity
    badges (🔴 / 🟡 / 🔵) and Path.read_text would fall back to
    cp1252 on Windows without it.
    """
    script = ROOT / "tools" / "generate_rule_docs.py"
    # Subprocess uses utf-8 stdout to avoid Windows console encoding issues
    # when the script prints status. The Python interpreter itself reads
    # the script as UTF-8 via PYTHONUTF8.
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    subprocess.run([sys.executable, str(script)], check=True, cwd=ROOT, env=env)
    snapshot = {
        p.name: p.read_text(encoding="utf-8")
        for p in _RULE_DIR.iterdir() if p.is_file()
    }
    subprocess.run([sys.executable, str(script)], check=True, cwd=ROOT, env=env)
    after = {
        p.name: p.read_text(encoding="utf-8")
        for p in _RULE_DIR.iterdir() if p.is_file()
    }
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
    text = (_RULE_DIR / "index.md").read_text(encoding="utf-8")
    for rule_id in known_rule_ids():
        assert re.search(rf"`{rule_id}`", text), f"{rule_id} missing from index.md"


def test_each_rule_page_has_required_sections():
    """Every per-rule page must include Description, fail/ok examples and
    a fix hint. This guards against shipping new rules with empty docs."""
    for rule in all_rules():
        page = (_RULE_DIR / f"{rule.rule_id}.md").read_text(encoding="utf-8")
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


# ---- Roundtrip compile-trigger tiered strategy (Class A regression) ---

class _FakeCompileFn:
    """Callable that simulates a hidden COM method via late-binding."""
    def __init__(self, behaviour):
        # behaviour: 'attribute_error' | 'clean' | 'compile_error'
        self.behaviour = behaviour
        self.call_count = 0

    def __call__(self):
        self.call_count += 1
        if self.behaviour == "attribute_error":
            raise AttributeError("<unknown>.Compile")
        if self.behaviour == "clean":
            return None
        if self.behaviour == "compile_error":
            # Mimic pywin32 com_error: (hresult, source, excepinfo, argerr)
            raise Exception((-2147352567, "Microsoft Excel", (
                1004, "VBA", "Compile error: Expected: end of statement",
                None, 0, -2146826276,
            ), None))
        raise RuntimeError("unknown test behaviour")


class _FakeVBProject:
    def __init__(self, compile_behaviour):
        self.Compile = _FakeCompileFn(compile_behaviour)


def test_attempt_compile_clean_compile_returns_empty():
    """Strategy 1: VBProject.Compile() returns cleanly → no issues."""
    from src.roundtrip import _attempt_compile
    vbproj = _FakeVBProject("clean")
    issues = _attempt_compile(vbproj, app=None, wb=None, host="excel", comp_name="M")
    assert issues == []


def test_attempt_compile_real_compile_error_yields_vba_rt001():
    """Strategy 1: VBProject.Compile() raises a real VBE error →
    one VBA_RT001 issue with the parsed description."""
    from src.roundtrip import _attempt_compile
    vbproj = _FakeVBProject("compile_error")
    issues = _attempt_compile(vbproj, app=None, wb=None, host="excel", comp_name="M")
    assert len(issues) == 1
    i = issues[0]
    assert i["rule_id"] == "VBA_RT001"
    assert i["severity"] == "compile_verified"
    # The description gets extracted from the com_error excepinfo tuple.
    assert "Compile error" in i["message"]
    assert "Expected" in i["message"]


def test_attempt_compile_hidden_method_falls_through_to_info():
    """Strategy 1 raises AttributeError (the bug we just fixed).
    Without an app/wb to run Strategy 2, we land on the info notice —
    crucially NOT a false-positive VBA_RT001."""
    from src.roundtrip import _attempt_compile
    vbproj = _FakeVBProject("attribute_error")
    # Pass app=None so Strategy 2's VBComponents.Add call fails and we
    # fall through to Strategy 3.
    issues = _attempt_compile(vbproj, app=None, wb=None, host="excel", comp_name="M")
    assert len(issues) == 1
    i = issues[0]
    assert i["rule_id"] == "VBA_RT000", (
        "AttributeError on Compile must NOT become a false VBA_RT001. "
        f"Got: {i!r}"
    )
    assert i["severity"] == "info"


def test_attempt_compile_strategy1_swallows_method_not_found_text():
    """When `Compile` raises with the well-known 'method not found'
    wording, that's still an indirection failure, not a real compile
    error — Strategy 1 must return None so Strategy 2 / 3 take over."""
    from src.roundtrip import _attempt_compile

    class _MNFCompile:
        def __init__(self):
            self.Compile = self._fn

        def _fn(self):
            raise Exception((-1, "VBA", (
                438, "VBA", "Method or data member not found",
                None, 0, 0,
            ), None))

    issues = _attempt_compile(_MNFCompile(), app=None, wb=None, host="excel", comp_name="M")
    assert all(i["rule_id"] != "VBA_RT001" for i in issues), (
        "'Method or data member not found' is an indirection failure, "
        "not a compile error."
    )


def test_vba_error_description_handles_plain_exception():
    """Plain `Exception("...")` (no excepinfo tuple) falls back to str()."""
    from src.roundtrip import _vba_error_description
    assert _vba_error_description(Exception("oops")) == "oops"


def test_vba_error_description_unwraps_com_error_tuple():
    """pywin32 com_error has the user-facing message at excepinfo[2]."""
    from src.roundtrip import _vba_error_description
    com_error = Exception((
        -2147352567, "Excel",
        (1004, "VBA", "Compile error: Sub or Function not defined", None, 0, 0),
        None,
    ))
    desc = _vba_error_description(com_error)
    assert "Sub or Function not defined" in desc


def test_format_macro_ref_quotes_workbook_name_for_excel():
    """Workbook names may contain spaces — Excel needs single quotes."""
    from src.roundtrip import _format_macro_ref

    class _Wb:
        Name = "My Workbook With Spaces.xlsm"

    ref = _format_macro_ref("excel", _Wb(), "Mod1", "Probe")
    assert ref == "'My Workbook With Spaces.xlsm'!Mod1.Probe"


def test_format_macro_ref_word_omits_workbook():
    from src.roundtrip import _format_macro_ref

    class _Doc:
        Name = "Doc.docm"

    ref = _format_macro_ref("word", _Doc(), "Mod1", "Probe")
    assert ref == "Mod1.Probe"  # Word doesn't prefix the document name
