"""Phase 2.13 — assignment / expression semantics (VBA160/164/165)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_assign_to_const(run_source):
    r = run_source("Const PI As Double = 3.14159\nSub S()\n    PI = 3\nEnd Sub\n")
    assert "VBA160" in _ids(r)


def test_assign_to_enum_member(run_source):
    r = run_source(
        "Enum Color\n    Red = 1\nEnd Enum\n"
        "Sub S()\n    Red = 5\nEnd Sub\n"
    )
    assert "VBA160" in _ids(r)


def test_read_const_ok(run_source):
    r = run_source("Const PI As Double = 3.14159\nSub S()\n    Dim r As Double\n    r = PI\nEnd Sub\n")
    assert "VBA160" not in _ids(r)


def test_assign_to_me(run_source):
    r = run_source("Sub S()\n    Set Me = Nothing\nEnd Sub\n", module_type="Class")
    assert "VBA164" in _ids(r)


def test_new_primitive(run_source):
    r = run_source("Sub S()\n    Dim x As New Long\nEnd Sub\n")
    assert "VBA165" in _ids(r)


def test_new_class_ok(run_source):
    r = run_source("Sub S()\n    Dim c As New Collection\nEnd Sub\n")
    assert "VBA165" not in _ids(r)
