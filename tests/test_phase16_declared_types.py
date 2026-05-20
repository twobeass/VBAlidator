"""Phase 1.6 — declared-type resolution (VBA120/122/123)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_unknown_local_type(run_source):
    r = run_source("Sub S()\n    Dim x As Custmer\nEnd Sub\n")
    assert "VBA120" in _ids(r)


def test_known_builtin_type_ok(run_source):
    r = run_source("Sub S()\n    Dim x As Long\n    Dim y As String\nEnd Sub\n")
    assert "VBA120" not in _ids(r)


def test_source_udt_type_ok(run_source):
    r = run_source(
        "Type Point\n    x As Long\n    y As Long\nEnd Type\n"
        "Sub S()\n    Dim p As Point\nEnd Sub\n"
    )
    assert "VBA120" not in _ids(r)


def test_qualified_type_lenient(run_source):
    r = run_source("Sub S()\n    Dim d As Scripting.Dictionary\nEnd Sub\n")
    assert "VBA120" not in _ids(r)


def test_unknown_arg_type(run_source):
    r = run_source("Sub S(ByVal p As Wibble)\nEnd Sub\n")
    assert "VBA120" in _ids(r)


def test_unknown_member_type(run_source):
    r = run_source("Type T\n    a As Nonexistent\nEnd Type\n")
    assert "VBA120" in _ids(r)


def test_empty_type(run_source):
    r = run_source("Type Empty1\nEnd Type\n")
    assert "VBA122" in _ids(r)


def test_empty_enum(run_source):
    r = run_source("Enum Empty2\nEnd Enum\n")
    assert "VBA123" in _ids(r)


def test_nonempty_type_ok(run_source):
    r = run_source("Type Point\n    x As Long\nEnd Type\n")
    assert "VBA122" not in _ids(r)
