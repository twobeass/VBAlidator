"""Phase 3.8 — conditional-compilation directives (VBA190-VBA192)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_else_without_if(run_source):
    r = run_source("#Else\n    x = 1\n#End If\n")
    assert "VBA190" in _ids(r)


def test_end_if_without_if(run_source):
    r = run_source("Sub S()\nEnd Sub\n#End If\n")
    assert "VBA190" in _ids(r)


def test_missing_end_if(run_source):
    r = run_source("#If Win64 Then\n    Dim x As Long\n")
    assert "VBA191" in _ids(r)


def test_balanced_conditional_ok(run_source):
    r = run_source("#If Win64 Then\n    Dim x As Long\n#Else\n    Dim y As Long\n#End If\n")
    assert "VBA190" not in _ids(r)
    assert "VBA191" not in _ids(r)


def test_invalid_const_no_value(run_source):
    r = run_source("#Const FOO\n")
    assert "VBA192" in _ids(r)


def test_valid_const_ok(run_source):
    r = run_source("#Const FOO = 1\n")
    assert "VBA192" not in _ids(r)
