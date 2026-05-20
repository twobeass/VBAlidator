"""Phase 2.10 — parameter-list well-formedness (VBA130-VBA136)."""
from __future__ import annotations

import pytest


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


CASES = [
    ("VBA130", "Sub S(Optional a As Long, b As Long)\nEnd Sub\n"),
    ("VBA131", "Sub S(ParamArray a() As Variant, b As Long)\nEnd Sub\n"),
    ("VBA132", "Sub S(ParamArray a() As Variant, ParamArray b() As Variant)\nEnd Sub\n"),
    ("VBA133", "Sub S(Optional a As Long, ParamArray b() As Variant)\nEnd Sub\n"),
    ("VBA134", "Sub S(ParamArray a() As Long)\nEnd Sub\n"),
    ("VBA136", "Sub S(ByVal a() As Long)\nEnd Sub\n"),
]


@pytest.mark.parametrize("rule_id, src", CASES, ids=[c[0] for c in CASES])
def test_param_list_violation(rule_id, src, run_source):
    result = run_source(src)
    assert rule_id in _ids(result), (
        f"Expected {rule_id}. Got: {[(e['rule_id'], e['message']) for e in result.errors]!r}"
    )


def test_udt_byval_param(run_source):
    r = run_source(
        "Type T\n    x As Long\nEnd Type\n"
        "Sub S(ByVal p As T)\nEnd Sub\n"
    )
    assert "VBA137" in _ids(r)


def test_udt_byref_param_ok(run_source):
    r = run_source(
        "Type T\n    x As Long\nEnd Type\n"
        "Sub S(ByRef p As T)\nEnd Sub\n"
    )
    assert "VBA137" not in _ids(r)


def test_enum_byval_param_ok(run_source):
    r = run_source(
        "Enum E\n    A = 1\nEnd Enum\n"
        "Sub S(ByVal p As E)\nEnd Sub\n"
    )
    assert "VBA137" not in _ids(r)


def test_well_formed_param_lists_are_clean(run_source):
    src = (
        "Sub A(a As Long, Optional b As Long, Optional c As String)\nEnd Sub\n"
        "Sub B(first As Long, ParamArray rest() As Variant)\nEnd Sub\n"
        "Function F(ByRef arr() As Long) As Long\nEnd Function\n"
        "Sub C(ParamArray items())\nEnd Sub\n"
    )
    result = run_source(src)
    param_ids = {f"VBA13{n}" for n in range(0, 10)}
    assert not (param_ids & _ids(result)), (
        f"Well-formed signatures must not raise VBA13x. Got: "
        f"{[(e['rule_id'], e['message']) for e in result.errors]!r}"
    )
