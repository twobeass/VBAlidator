"""Tests for the Premium-Prechecker API (P4.1–P4.4):
- precheck() function and PrecheckResult dataclass
- compute_score() weighting
- normalize_issues() rule-ID inference
- JSON v2 report shape
- --host model auto-loading (Excel)
"""
from __future__ import annotations

from pathlib import Path

import pytest

from src.api import PrecheckResult, precheck, precheck_source
from src.reporting import build_report_v2, normalize_issue, normalize_issues
from src.scoring import compute_score, is_compile_safe


# ---- Score ---------------------------------------------------------------


def test_score_clean_input_is_100():
    score, _ = compute_score([])
    assert score == 100


def test_score_one_error_drops_below_threshold():
    issues = [{"severity": "error", "message": "x"}]
    score, breakdown = compute_score(issues)
    assert score == 80
    assert breakdown["by_severity"]["error"] == 1
    assert breakdown["penalty_total"] == 20


def test_score_warnings_count_less_than_errors():
    score_w, _ = compute_score([{"severity": "warning"}])
    score_e, _ = compute_score([{"severity": "error"}])
    assert score_w > score_e


def test_score_floors_at_zero():
    score, _ = compute_score([{"severity": "error"}] * 100)
    assert score == 0


def test_score_coverage_uncertain_caps_at_90():
    score, _ = compute_score([], coverage_uncertain=True)
    assert score == 90


def test_compile_safe_only_errors_block():
    assert is_compile_safe([])
    assert is_compile_safe([{"severity": "warning"}])
    assert is_compile_safe([{"severity": "info"}])
    assert not is_compile_safe([{"severity": "error"}])


# ---- Reporting -----------------------------------------------------------


def test_normalize_issue_legacy_undefined():
    n = normalize_issue({"file": "a.bas", "line": 5, "message": "Undefined identifier 'foo' in 'S'."})
    assert n["rule_id"] == "VBA001"
    assert n["severity"] == "error"
    assert n["category"] == "name_resolution"


def test_normalize_issue_legacy_unreachable_is_warning():
    n = normalize_issue({"line": 9, "message": "Unreachable code detected in 'S'."})
    assert n["rule_id"] == "VBA009"
    assert n["severity"] == "warning"


def test_normalize_issue_keeps_explicit_rule_id():
    n = normalize_issue({
        "rule_id": "VBA210",
        "severity": "error",
        "message": "`Set` used on non-object …",
    })
    assert n["rule_id"] == "VBA210"
    assert n["severity"] == "error"
    assert n["category"] == "assignment"


def test_normalize_issues_idempotent():
    first = normalize_issues([{"message": "Undefined identifier 'x'"}])
    second = normalize_issues(first)
    assert second == first


def test_build_report_v2_shape():
    issues = [
        {"file": "a.bas", "line": 1, "message": "Undefined identifier 'x' in 'S'."},
        {"file": "a.bas", "line": 2, "rule_id": "VBA210", "severity": "error", "message": "Set on scalar"},
        {"file": "b.bas", "line": 7, "rule_id": "VBA320", "severity": "warning", "message": "Option Explicit"},
    ]
    report = build_report_v2(issues, files_scanned=2, score=77, compile_safe=False)
    assert report["version"] == "2.0"
    assert report["summary"]["files_scanned"] == 2
    assert report["summary"]["score"] == 77
    assert report["summary"]["compile_safe"] is False
    assert report["summary"]["errors"] == 2
    assert report["summary"]["warnings"] == 1
    # files grouped + sorted
    paths = [f["path"] for f in report["files"]]
    assert paths == sorted(paths)
    assert all("issues" in f for f in report["files"])
    # flat list also present for jq consumers
    assert isinstance(report["issues"], list)
    assert len(report["issues"]) == 3


# ---- precheck() public API ----------------------------------------------


def test_precheck_clean_inline_source():
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim x As Long
    x = 1
End Sub
"""
    result = precheck_source(code)
    assert isinstance(result, PrecheckResult)
    assert result.compile_safe
    assert result.score == 100
    assert result.errors == []
    assert bool(result) is True


def test_precheck_inline_source_with_undefined():
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    typo = 1
End Sub
"""
    result = precheck_source(code)
    assert not result.compile_safe
    assert result.score < 100
    assert any(e["rule_id"] == "VBA001" for e in result.errors)
    assert bool(result) is False


def test_precheck_result_json_is_v2():
    code = """Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim x As Long
    x = 1
End Sub
"""
    result = precheck_source(code)
    j = result.json()
    assert j["version"] == "2.0"
    assert j["summary"]["compile_safe"]
    assert j["summary"]["score"] == 100


def test_precheck_demo_directory_works():
    """Run on the demo fixtures and verify the result is well-formed."""
    repo = Path(__file__).resolve().parent.parent
    result = precheck(repo / "tests" / "demo")
    assert result.files_scanned >= 1
    j = result.json()
    assert "summary" in j
    assert isinstance(j["files"], list)


def test_precheck_strict_vs_non_strict():
    """`strict=False` means warnings do not affect compile_safe / score."""
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    x = 1
End Sub
"""
    # Missing Option Explicit triggers a warning.
    strict = precheck_source(code)
    assert strict.warnings, "Missing Option Explicit must produce VBA320 warning"
    # `precheck_source` always operates strict — non-strict gating is
    # exposed via `precheck()` and the CLI `--no-strict` flag, which is
    # exercised in the CLI smoke test of ci.yml.


# ---- Host models --------------------------------------------------------


def test_excel_host_model_resolves_workbook(tmp_path):
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim wb As Workbook\n"
        "    Set wb = ActiveWorkbook\n"
        "    wb.Save\n"
        "End Sub\n",
    )
    result = precheck(bas, host="excel")
    assert all(e["rule_id"] != "VBA001" for e in result.errors), (
        f"With --host=excel, ActiveWorkbook + Workbook.Save must resolve. "
        f"Errors: {result.errors!r}"
    )


def test_no_host_yields_more_undefined(tmp_path):
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim wb As Object\n"
        "    Set wb = ActiveWorkbook\n"  # ActiveWorkbook is in std_model
        "End Sub\n",
    )
    result_no_host = precheck(bas)
    # ActiveWorkbook IS in std_model so it must resolve even without --host.
    assert all("ActiveWorkbook" not in e.get("message", "") for e in result_no_host.errors)


def test_word_host_loads_documents_class(tmp_path):
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim doc As Document\n"
        "    Set doc = ActiveDocument\n"
        "    doc.Save\n"
        "End Sub\n",
    )
    result = precheck(bas, host="word")
    assert all(e["rule_id"] != "VBA001" for e in result.errors), (
        f"Word host must resolve ActiveDocument + Document. Errors: {result.errors!r}"
    )


def test_unknown_host_does_not_crash(tmp_path):
    bas = tmp_path / "M.bas"
    bas.write_text('Attribute VB_Name = "M"\nOption Explicit\nSub S()\nEnd Sub\n')
    # Silently no-ops; user just gets the std_model coverage.
    result = precheck(bas, host="not_a_real_host")
    assert result.compile_safe


# ---- MSComCtl auto-layer ----------------------------------------------------


def _MSCOMCTL_FORM(call: str) -> str:
    """A minimal `.frm` referencing the Microsoft Common Controls library
    plus the inline VBA snippet `call` inside `Form_Resize`."""
    return (
        "VERSION 5.00\n"
        "Object = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0\"; \"MSCOMCTL.OCX\"\n"
        "Begin VB.Form Form1\n"
        "   Caption = \"Test\"\n"
        "   Begin ComctlLib.TreeView TreeView1\n"
        "      Height = 2535\n"
        "   End\n"
        "End\n"
        'Attribute VB_Name = "Form1"\n'
        "Option Explicit\n"
        "Private Sub Form_Resize()\n"
        f"    {call}\n"
        "End Sub\n"
    )


def test_mscomctl_autolayers_when_frm_references_comctllib(tmp_path):
    """A `.frm` declaring a `ComctlLib.TreeView` control must auto-layer
    `models/mscomctl.json` so member access on the control (e.g.
    `TreeView1.Move`) resolves without the user passing `--host mscomctl`."""
    frm = tmp_path / "Form1.frm"
    frm.write_text(_MSCOMCTL_FORM("TreeView1.Move 0, 0, 100, 100"), encoding="latin-1")
    result = precheck(frm)
    member_errors = [e for e in result.errors if "Move" in e.get("message", "")]
    assert not member_errors, (
        f"TreeView.Move must resolve via auto-layered mscomctl model. "
        f"Errors: {result.errors!r}"
    )


def test_mscomctl_resolves_listview_and_progressbar_members(tmp_path):
    """The other shipped control classes (ListView, ProgressBar, …) also
    pick up the VB6 container-control base members."""
    frm = tmp_path / "Form2.frm"
    frm.write_text(
        "VERSION 5.00\n"
        "Begin VB.Form Form2\n"
        "   Begin ComctlLib.ListView LV1\n"
        "      Height = 1000\n"
        "   End\n"
        "   Begin ComctlLib.ProgressBar PB1\n"
        "      Height = 200\n"
        "   End\n"
        "End\n"
        'Attribute VB_Name = "Form2"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    LV1.Visible = True\n"
        "    PB1.Width = 200\n"
        "End Sub\n",
        encoding="latin-1",
    )
    result = precheck(frm)
    member_errors = [
        e for e in result.errors
        if "Visible" in e.get("message", "") or "Width" in e.get("message", "")
    ]
    assert not member_errors, (
        f"ListView.Visible / ProgressBar.Width must resolve. Errors: {result.errors!r}"
    )


def test_mscomctl_autolayer_explicit_host_choice(tmp_path):
    """`--host mscomctl` is also a valid CLI choice — the model loads
    even when no `.frm` triggers the auto-layer (covers the explicit
    opt-in path for users analyzing `.bas` modules that use the controls
    via late binding)."""
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(t As TreeView)\n"
        "    t.Move 0, 0, 100, 100\n"
        "End Sub\n",
    )
    result = precheck(bas, host="mscomctl")
    member_errors = [e for e in result.errors if "Move" in e.get("message", "")]
    assert not member_errors, (
        f"With --host=mscomctl, TreeView.Move must resolve. "
        f"Errors: {result.errors!r}"
    )


# ---- MSForms auto-layer -----------------------------------------------------


def test_msforms_autolayers_when_source_references_namespace(tmp_path):
    """A `.cls` using `TypeOf ctrl Is MSForms.UserForm` must auto-layer
    `models/msforms.json` so the namespace lookup resolves — without
    requiring `--host msforms` explicitly. Regression for stdVBA-master's
    `stdUIElement.cls`."""
    cls = tmp_path / "C.cls"
    cls.write_text(
        "VERSION 1.0 CLASS\n"
        'Attribute VB_Name = "C"\n'
        "Option Explicit\n"
        "Sub S(ctrl As Object)\n"
        "    If TypeOf ctrl Is MSForms.UserForm Then\n"
        "    End If\n"
        "    If TypeOf ctrl Is MSForms.CommandButton Then\n"
        "    End If\n"
        "    If TypeOf ctrl Is MSForms.TextBox Then\n"
        "    End If\n"
        "End Sub\n",
        encoding="latin-1",
    )
    result = precheck(cls)
    msforms_errors = [e for e in result.errors if "MSForms" in e.get("message", "")]
    assert not msforms_errors, (
        f"`MSForms.<Class>` references must resolve via auto-layered "
        f"msforms model. Got: {result.errors!r}"
    )


def test_msforms_control_base_members_resolve(tmp_path):
    """The abstract `MSForms.Control` base picks up the VB6 container-
    control member set (`Caption`, `Left`, `Top`, `Width`, `Height`, …)
    so library code like `this.Control.Caption` resolves. Regression for
    stdVBA-master's `stdUIElement.cls::Caption` property."""
    cls = tmp_path / "C.cls"
    cls.write_text(
        "VERSION 1.0 CLASS\n"
        'Attribute VB_Name = "C"\n'
        "Option Explicit\n"
        "Private ctrl As MSForms.Control\n"
        "Sub S()\n"
        "    Dim s As String\n"
        "    s = ctrl.Caption\n"
        "    ctrl.Left = 10\n"
        "    ctrl.Top = 20\n"
        "    ctrl.Width = 100\n"
        "    ctrl.Height = 50\n"
        "End Sub\n",
        encoding="latin-1",
    )
    result = precheck(cls)
    member_errors = [
        e for e in result.errors
        if "not found in type 'MSForms.Control'" in e.get("message", "")
    ]
    assert not member_errors, (
        f"MSForms.Control must carry VB6 base members. Got: {result.errors!r}"
    )


def test_msforms_autolayer_explicit_host_choice(tmp_path):
    """`--host msforms` is a valid CLI choice for `.bas` modules that use
    UserForms via late binding (no `MSForms.` prefix to trigger the
    auto-layer)."""
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(btn As CommandButton)\n"
        "    btn.Caption = \"Click\"\n"
        "End Sub\n",
    )
    result = precheck(bas, host="msforms")
    member_errors = [e for e in result.errors if "Caption" in e.get("message", "")]
    assert not member_errors, (
        f"With --host=msforms, CommandButton.Caption must resolve. "
        f"Errors: {result.errors!r}"
    )


# ---- Defines -----------------------------------------------------------


def test_defines_can_disable_ptrsafe_check(tmp_path):
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        'Private Declare Function GetTickCount Lib "kernel32" () As Long\n'
        "Sub S()\nEnd Sub\n",
    )
    # Default: VBA300 fires.
    default_result = precheck(bas)
    assert any(e["rule_id"] == "VBA300" for e in default_result.errors)

    # With WIN64=False the rule is suppressed.
    relaxed = precheck(bas, defines={"WIN64": False, "VBA7": False})
    assert all(e["rule_id"] != "VBA300" for e in relaxed.errors)


# ---- Issue list filters --------------------------------------------------


def test_result_buckets_split_by_severity():
    code = """
Attribute VB_Name = "M"
Sub S()
    typo = 1
End Sub
"""
    result = precheck_source(code)
    # We expect at least one error (typo) AND one warning (missing Option Explicit).
    assert result.errors, "must surface errors"
    assert result.warnings, "must surface VBA320 warning"
    # Error/warning sets are disjoint.
    err_msgs = {e["message"] for e in result.errors}
    warn_msgs = {w["message"] for w in result.warnings}
    assert not err_msgs & warn_msgs


# ---- Models JSON well-formed --------------------------------------------


@pytest.mark.parametrize("host", ["excel", "word", "access", "outlook", "visio", "mscomctl", "msforms"])
def test_host_model_json_loads(host):
    """Every shipped host model must be valid JSON with the expected sections."""
    import json
    path = Path(__file__).resolve().parent.parent / "src" / "models" / f"{host}.json"
    assert path.is_file(), f"Missing host model {path}"
    data = json.loads(path.read_text())
    assert "globals" in data or "classes" in data, (
        f"{host}.json must declare globals or classes"
    )
    assert isinstance(data.get("globals", {}), dict)
    assert isinstance(data.get("classes", {}), dict)


# ---- vba_model.json auto-load -------------------------------------------


def _write_minimal_model(path):
    """Write a tiny custom model defining a class+global that nothing
    else in the std/host catalogues knows about — so we can prove the
    auto-load path actually merged it."""
    import json
    path.write_text(json.dumps({
        "globals": {
            "MyCustomGlobal": {"type": "Long", "kind": "Constant"},
        },
        "classes": {
            "MyAddinClass": {
                "members": {"DoStuff": {"type": "Sub"}},
            },
        },
    }))


def test_vba_model_json_auto_load_from_input_dir(tmp_path):
    """Drop a `vba_model.json` next to the input folder and verify the
    documented auto-load behaviour wires it through to the analyser."""
    workdir = tmp_path / "myproject"
    workdir.mkdir()
    (workdir / "M.bas").write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim x As Long\n"
        "    x = MyCustomGlobal\n"
        "End Sub\n"
    )
    _write_minimal_model(workdir / "vba_model.json")

    result = precheck(workdir)
    assert all(
        "MyCustomGlobal" not in e.get("message", "") for e in result.errors
    ), (
        "Auto-loaded vba_model.json must register MyCustomGlobal. "
        f"Errors: {result.errors!r}"
    )


def test_vba_model_json_auto_load_from_cwd(tmp_path, monkeypatch):
    """Auto-load also picks up a `vba_model.json` from the CWD even
    when the input lives in a different folder."""
    code_dir = tmp_path / "code"
    code_dir.mkdir()
    bas = code_dir / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim x As Long\n"
        "    x = MyCustomGlobal\n"
        "End Sub\n"
    )
    cwd = tmp_path / "cwd"
    cwd.mkdir()
    _write_minimal_model(cwd / "vba_model.json")
    monkeypatch.chdir(cwd)

    result = precheck(bas)
    assert all(
        "MyCustomGlobal" not in e.get("message", "") for e in result.errors
    ), (
        "CWD-level vba_model.json must be auto-loaded. "
        f"Errors: {result.errors!r}"
    )


def test_vba_model_json_explicit_model_path_wins(tmp_path):
    """Explicit `model_path=` overrides any auto-load candidate so
    callers can pin a specific model from a CI config."""
    workdir = tmp_path / "code"
    workdir.mkdir()
    (workdir / "M.bas").write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim x As Long\n"
        "    x = OnlyInExplicitModel\n"
        "End Sub\n"
    )

    # The auto-load candidate registers a *different* global so we can
    # tell which one actually got loaded.
    _write_minimal_model(workdir / "vba_model.json")

    explicit = tmp_path / "explicit.json"
    import json as _json
    explicit.write_text(_json.dumps({
        "globals": {"OnlyInExplicitModel": {"type": "Long"}}
    }))

    result = precheck(workdir, model_path=explicit)
    # The explicit model knows OnlyInExplicitModel.
    assert all(
        "OnlyInExplicitModel" not in e.get("message", "") for e in result.errors
    )


def test_inline_source_does_not_trigger_auto_load(tmp_path, monkeypatch):
    """For inline source strings there is no input directory — we must
    not accidentally promote any nearby `vba_model.json`."""
    monkeypatch.chdir(tmp_path)
    _write_minimal_model(tmp_path / "vba_model.json")
    result = precheck_source(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(): End Sub\n",
    )
    # `precheck_source` is the inline path — it must be hermetic.
    assert result.compile_safe
    # The custom-model global must NOT have been registered (sanity:
    # `precheck_source` doesn't take a path-like argument so the
    # auto-load heuristic shouldn't fire).
    # Touching MyCustomGlobal would resolve only if it was loaded.
    inline_with_use = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim x As Long\n"
        "    x = MyCustomGlobal\n"
        "End Sub\n"
    )
    result2 = precheck_source(inline_with_use)
    assert any(
        "MyCustomGlobal" in e.get("message", "") for e in result2.errors
    ), "precheck_source must not auto-load nearby vba_model.json"


# ---- CLI surface (UAT §0, §2) -------------------------------------------


def test_version_flag_prints_package_version():
    """`vbalidator --version` must exit 0 and print `vbalidator <version>`
    so UAT §0 can pin the installed build."""
    import subprocess
    import sys
    out = subprocess.run(
        [sys.executable, "-m", "src.main", "--version"],
        capture_output=True, text=True, check=False,
    )
    assert out.returncode == 0, (out.returncode, out.stderr)
    assert "vbalidator" in out.stdout.lower()
    from src import __version__
    assert __version__ in out.stdout


def test_quiet_flag_suppresses_stdout_on_clean_input(tmp_path):
    """UAT §2 row 1: `vbalidator --quiet --no-strict <clean>` exits 0 and
    writes the JSON report — stdout is silent so CI tools that pipe
    stdout don't get banner noise."""
    import subprocess
    import sys
    bas = tmp_path / "M.bas"
    bas.write_text(
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(): End Sub\n",
        encoding="utf-8",
    )
    report = tmp_path / "report.json"
    out = subprocess.run(
        [sys.executable, "-m", "src.main", str(bas),
         "--quiet", "--no-strict", "--output", str(report)],
        capture_output=True, text=True, check=False,
    )
    assert out.returncode == 0, (out.returncode, out.stdout, out.stderr)
    # --quiet means stdout is empty (or only colorama reset codes).
    visible = "".join(ch for ch in out.stdout if ch.isprintable() and ch != ' ')
    assert visible == "", f"--quiet must silence stdout. Got: {out.stdout!r}"
    # JSON report still produced.
    import json
    d = json.loads(report.read_text(encoding="utf-8"))
    assert d["version"] == "2.0"
    assert d["summary"]["score"] == 100
    assert d["summary"]["compile_safe"] is True


def test_pipeline_errors_route_to_stderr_under_quiet(tmp_path):
    """Hard errors (invalid input path) must reach stderr even with
    --quiet, so CI logs surface real failures."""
    import subprocess
    import sys
    out = subprocess.run(
        [sys.executable, "-m", "src.main", "/nonexistent/path",
         "--quiet", "--output", str(tmp_path / "x.json")],
        capture_output=True, text=True, check=False,
    )
    assert out.returncode == 2, out.returncode
    # Error message lands on stderr, NOT stdout (which would corrupt
    # CI pipes).
    assert out.stdout == "", f"stdout should be empty: {out.stdout!r}"
    assert "does not exist" in out.stderr.lower()


# ---- Public import contract (UAT §5) ------------------------------------


def test_vbalidator_package_top_level_imports():
    """UAT §5: after `pip install vbalidator`, the user-facing import is
    `from vbalidator import precheck, precheck_source, PrecheckResult`.

    The implementation still lives under `src/` (renamed in a future
    PR); a compatibility shim at `vbalidator/__init__.py` re-exports
    the public surface so the import name matches the PyPI
    distribution name.
    """
    import importlib
    mod = importlib.import_module("vbalidator")
    for name in ("precheck", "precheck_source", "PrecheckResult", "__version__"):
        assert hasattr(mod, name), f"vbalidator.{name} missing — UAT §5 broken"

    # Identity check: the re-exported precheck must BE the same object
    # `from src import precheck` returns. Otherwise monkeypatches in
    # one namespace silently miss the other.
    from src import precheck as src_precheck
    from vbalidator import precheck as vba_precheck
    assert src_precheck is vba_precheck, (
        "vbalidator.precheck must alias src.precheck, not duplicate it"
    )


def test_vbalidator_submodule_imports_resolve():
    """Inner-module imports must resolve too — tooling that does
    `from vbalidator.rules import all_rules` shouldn't have to know
    about the src/ implementation detail."""
    import importlib
    for sub in ("api", "scoring", "reporting", "rules", "roundtrip",
                "analyzer", "lexer", "parser", "preprocessor", "config"):
        importlib.import_module(f"vbalidator.{sub}")


def test_vbalidator_version_matches_pyproject():
    """`vbalidator.__version__` is the canonical published version and
    must equal `src.__version__`. python-semantic-release writes both."""
    from src import __version__ as src_version
    from vbalidator import __version__ as pkg_version
    assert src_version == pkg_version
