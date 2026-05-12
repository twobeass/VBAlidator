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


def test_attempt_compile_hidden_method_falls_through_to_inconclusive():
    """Strategy 1 raises AttributeError (Compile is hidden on modern
    Office). With Strategy 2 also unavailable (no app/wb attached),
    we land on `VBA_RT002` inconclusive — crucially NOT a false-positive
    VBA_RT001, AND distinct from VBA_RT000 (platform-missing) because
    VBE *was* reachable here; we just couldn't drive a compile."""
    from src.roundtrip import _attempt_compile
    vbproj = _FakeVBProject("attribute_error")
    # Pass app=None so Strategy 2's VBComponents.Add call fails and we
    # fall through to Strategy 3.
    issues = _attempt_compile(vbproj, app=None, wb=None, host="excel", comp_name="M")
    assert len(issues) == 1
    i = issues[0]
    assert i["rule_id"] == "VBA_RT002", (
        "AttributeError on Compile must NOT become a false VBA_RT001, "
        "and must be distinct from VBA_RT000 (platform-missing). "
        f"Got: {i!r}"
    )
    assert i["severity"] == "warning"


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


# ---- Class-A fixes: Attribute strip + Inconclusive vs Unavailable -----


def test_strip_export_directives_removes_attribute_header():
    """The classic `.bas` export header MUST be stripped before
    injection — VBE rejects Attribute statements inside a module body."""
    from src.roundtrip import _strip_export_directives
    raw = (
        'Attribute VB_Name = "ValidCode"\n'
        "Sub TestValid()\n"
        "    Dim i As Integer\n"
        "End Sub\n"
    )
    out = _strip_export_directives(raw)
    assert "Attribute " not in out, f"Attribute survived: {out!r}"
    assert "VB_Name" not in out
    assert "Sub TestValid()" in out, f"Body lost: {out!r}"
    assert out.startswith("Sub"), (
        "Leading blank lines should be eaten so the resulting module "
        "isn't prefixed by spurious blanks."
    )


def test_strip_export_directives_removes_class_begin_block():
    """`.cls` files use the BEGIN…END attribute block — strip it too."""
    from src.roundtrip import _strip_export_directives
    raw = (
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "END\n"
        'Attribute VB_Name = "MyClass"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Option Explicit\n"
        "\n"
        "Public Sub Greet()\n"
        "    Debug.Print \"hi\"\n"
        "End Sub\n"
    )
    out = _strip_export_directives(raw)
    assert "VERSION " not in out
    assert "BEGIN" not in out
    assert "MultiUse" not in out
    assert "Attribute " not in out
    assert out.startswith("Option Explicit"), out


def test_strip_export_directives_preserves_blanks_inside_body():
    """Blank lines *after* the header start are part of the code and
    must be kept — only header padding is collapsed."""
    from src.roundtrip import _strip_export_directives
    raw = (
        'Attribute VB_Name = "M"\n'
        "\n"
        "Sub First()\n"
        "End Sub\n"
        "\n"
        "Sub Second()\n"
        "End Sub\n"
    )
    out = _strip_export_directives(raw)
    # The blank between First and Second stays.
    assert out.count("\n\n") >= 1
    assert "Sub First" in out and "Sub Second" in out


def test_strip_export_directives_idempotent_on_clean_input():
    """Stripping code without a header must return it unchanged."""
    from src.roundtrip import _strip_export_directives
    clean = "Sub S()\n    Debug.Print 1\nEnd Sub\n"
    assert _strip_export_directives(clean) == clean


def test_inconclusive_issue_distinct_from_unavailable():
    """`VBA_RT000` (unavailable) and `VBA_RT002` (inconclusive) must
    have different rule_ids and severities so callers can tell whether
    VBE was reachable at all."""
    from src.roundtrip import _make_inconclusive_issue, _make_unavailable_issue
    u = _make_unavailable_issue("no pywin32")
    i = _make_inconclusive_issue("all strategies failed")
    assert u["rule_id"] == "VBA_RT000"
    assert u["severity"] == "info"
    assert i["rule_id"] == "VBA_RT002"
    assert i["severity"] == "warning"
    # Distinct categories or messages — but both flagged as roundtrip.
    assert u["category"] == i["category"] == "roundtrip"


def test_attempt_compile_falls_through_to_inconclusive_not_unavailable():
    """Critical regression: when both Strategy 1 and Strategy 2 fail
    (the case the UAT runner hit), we must emit `VBA_RT002`, not
    `VBA_RT000`. The previous build emitted RT000 which is wrong on a
    Windows host that *has* VBE."""
    from src.roundtrip import _attempt_compile

    class _Fake:
        def __init__(self):
            self.Compile = self._compile
            self.VBComponents = self  # for `vbproj.VBComponents.Add`

        def _compile(self):
            raise AttributeError("<unknown>.Compile")

        def Add(self, _kind):
            raise RuntimeError("VBComponents.Add refused (simulated)")

    issues = _attempt_compile(_Fake(), app=None, wb=None, host="excel", comp_name="M")
    assert len(issues) == 1
    assert issues[0]["rule_id"] == "VBA_RT002", issues[0]
    assert issues[0]["severity"] == "warning"


def test_run_with_timeout_returns_function_result_when_fast():
    """The timeout wrapper must be a no-op when the worker completes
    within the budget."""
    from src.roundtrip import _run_with_timeout
    result = _run_with_timeout(lambda: [{"ok": True}], (), timeout_s=5, host="excel")
    assert result == [{"ok": True}]


def test_run_with_timeout_emits_inconclusive_on_overrun():
    """When the worker blocks past `timeout_s`, we must surface a
    `VBA_RT002` inconclusive issue rather than hanging the parent.
    Office-kill is best-effort (taskkill missing on Linux) — that's
    OK, the parent still returns."""
    from src.roundtrip import _run_with_timeout
    import time

    def _slow():
        time.sleep(10)
        return []

    out = _run_with_timeout(_slow, (), timeout_s=0.2, host="excel")
    assert len(out) == 1
    assert out[0]["rule_id"] == "VBA_RT002"
    assert "exceeded" in out[0]["message"].lower()


def test_run_with_timeout_propagates_worker_exception():
    """If the worker raises within the budget, that exception must
    propagate — we don't want to silently swallow real bugs."""
    from src.roundtrip import _run_with_timeout

    def _boom():
        raise RuntimeError("real bug")

    with pytest.raises(RuntimeError, match="real bug"):
        _run_with_timeout(_boom, (), timeout_s=5, host="excel")


# ---- Worker-thread COM apartment init (Class A #4) --------------------


def test_run_with_timeout_calls_coinitialize_when_pythoncom_present(monkeypatch):
    """A worker thread does not inherit the main thread's COM apartment.
    Without `CoInitialize()` the first COM Dispatch raises HRESULT
    0x800401F0 (CO_E_NOTINITIALIZED). Verify the wrapper initialises
    and tears down the apartment for the worker.
    """
    import sys
    import types
    from src import roundtrip

    calls: list[str] = []

    fake_pythoncom = types.ModuleType("pythoncom")
    fake_pythoncom.CoInitialize = lambda: calls.append("init")  # type: ignore[attr-defined]
    fake_pythoncom.CoUninitialize = lambda: calls.append("uninit")  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    result = roundtrip._run_with_timeout(
        lambda: [{"ok": True}], (), timeout_s=5, host="excel",
    )
    assert result == [{"ok": True}]
    # Init and uninit must be paired, and init must precede the target call.
    assert calls == ["init", "uninit"], calls


def test_run_with_timeout_uninitializes_com_even_on_worker_exception(monkeypatch):
    """When the worker raises, CoUninitialize still runs (finally
    block) so we don't leak apartment references on every failing
    invocation."""
    import sys
    import types
    from src import roundtrip

    calls: list[str] = []
    fake_pythoncom = types.ModuleType("pythoncom")
    fake_pythoncom.CoInitialize = lambda: calls.append("init")  # type: ignore[attr-defined]
    fake_pythoncom.CoUninitialize = lambda: calls.append("uninit")  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    def _boom():
        raise RuntimeError("worker crashed")

    with pytest.raises(RuntimeError, match="worker crashed"):
        roundtrip._run_with_timeout(_boom, (), timeout_s=5, host="excel")
    # Init succeeded → uninit must follow.
    assert calls == ["init", "uninit"], calls


def test_run_with_timeout_works_without_pythoncom(monkeypatch):
    """On non-Windows hosts pythoncom isn't installed — the wrapper
    must still work for the unit tests that exercise the surrounding
    plumbing."""
    import sys
    from src import roundtrip

    # Force pythoncom-import failure by injecting a sentinel that raises.
    monkeypatch.setitem(sys.modules, "pythoncom", None)

    # NOTE: monkeypatching to None makes `import pythoncom` raise
    # ImportError; the wrapper catches that and proceeds without
    # COM init.
    result = roundtrip._run_with_timeout(
        lambda: [{"ok": True}], (), timeout_s=5, host="excel",
    )
    assert result == [{"ok": True}]
