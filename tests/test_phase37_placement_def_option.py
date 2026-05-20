"""Phase 3.7 — Type/Enum/Def/Option placement (VBA180-VBA184)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_statement_inside_enum(run_source):
    r = run_source("Enum E\n    A = 1\n    Dim x As Long\nEnd Enum\n")
    assert "VBA180" in _ids(r)


def test_clean_enum(run_source):
    r = run_source("Enum E\n    A = 1\n    B = 2\nEnd Enum\n")
    assert "VBA180" not in _ids(r)


def test_statement_inside_type(run_source):
    r = run_source("Type T\n    x As Long\n    Sub Foo()\nEnd Type\n")
    assert "VBA181" in _ids(r)


def test_clean_type(run_source):
    r = run_source("Type T\n    x As Long\n    y As String\nEnd Type\n")
    assert "VBA181" not in _ids(r)


def test_deftype_after_declaration(run_source):
    r = run_source("Dim x As Long\nDefInt A-Z\n")
    assert "VBA182" in _ids(r)


def test_deftype_before_declaration_ok(run_source):
    r = run_source("DefInt A-Z\nDim x As Long\n")
    assert "VBA182" not in _ids(r)


def test_duplicate_deftype_range(run_source):
    r = run_source("DefInt A-K\nDefStr H-Z\n")
    assert "VBA183" in _ids(r)


def test_disjoint_deftype_ranges_ok(run_source):
    r = run_source("DefInt A-G\nDefStr H-Z\n")
    assert "VBA183" not in _ids(r)


def test_duplicate_option_explicit(run_source):
    r = run_source("Option Explicit\nOption Explicit\n")
    assert "VBA184" in _ids(r)


def test_single_option_explicit_ok(run_source):
    r = run_source("Option Explicit\n")
    assert "VBA184" not in _ids(r)
