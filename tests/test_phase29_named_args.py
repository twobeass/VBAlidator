"""Phase 2.9 — named-argument validation at call sites (VBA140/141/142)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_named_arg_not_found(run_source):
    r = run_source(
        "Sub S(ByVal a As Long)\nEnd Sub\n"
        "Sub T()\n    Call S(b:=1)\nEnd Sub\n"
    )
    assert "VBA140" in _ids(r)


def test_named_arg_found_ok(run_source):
    r = run_source(
        "Sub S(ByVal a As Long)\nEnd Sub\n"
        "Sub T()\n    Call S(a:=1)\nEnd Sub\n"
    )
    assert "VBA140" not in _ids(r)


def test_named_arg_duplicate(run_source):
    r = run_source(
        "Sub S(ByVal a As Long, ByVal b As Long)\nEnd Sub\n"
        "Sub T()\n    Call S(a:=1, a:=2)\nEnd Sub\n"
    )
    assert "VBA141" in _ids(r)


def test_named_args_with_paramarray(run_source):
    r = run_source(
        "Sub S(ParamArray a() As Variant)\nEnd Sub\n"
        "Sub T()\n    Call S(a:=1)\nEnd Sub\n"
    )
    assert "VBA142" in _ids(r)


def test_positional_paramarray_ok(run_source):
    r = run_source(
        "Sub S(ParamArray a() As Variant)\nEnd Sub\n"
        "Sub T()\n    Call S(1, 2, 3)\nEnd Sub\n"
    )
    assert "VBA142" not in _ids(r)
