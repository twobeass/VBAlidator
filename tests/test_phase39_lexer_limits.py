"""Phase 3.9 — lexer-level limits (VBA_LEX004-VBA_LEX006)."""
from __future__ import annotations


def _ids(result):
    return {e.get("rule_id") for e in result.errors}


def test_unterminated_string(run_source):
    r = run_source('Sub S()\n    Dim x As String\n    x = "hello\nEnd Sub\n')
    assert "VBA_LEX004" in _ids(r)


def test_missing_end_bracket(run_source):
    r = run_source("Sub S()\n    Dim r As Range\n    Set r = [A1\nEnd Sub\n")
    assert "VBA_LEX004" in _ids(r)


def test_terminated_string_ok(run_source):
    r = run_source('Sub S()\n    Dim x As String\n    x = "hello"\nEnd Sub\n')
    assert "VBA_LEX004" not in _ids(r)


def test_identifier_too_long(run_source):
    name = "a" * 300
    r = run_source(f"Sub S()\n    Dim {name} As Long\nEnd Sub\n")
    assert "VBA_LEX005" in _ids(r)


def test_normal_identifier_ok(run_source):
    r = run_source("Sub S()\n    Dim total As Long\nEnd Sub\n")
    assert "VBA_LEX005" not in _ids(r)


def test_line_too_long(run_source):
    long_line = "    x = " + " + ".join(["1"] * 600)  # well over 1023 chars
    r = run_source(f"Sub S()\n{long_line}\nEnd Sub\n")
    assert "VBA_LEX006" in _ids(r)


def test_normal_line_ok(run_source):
    r = run_source("Sub S()\n    x = 1 + 2 + 3\nEnd Sub\n")
    assert "VBA_LEX006" not in _ids(r)


def test_unterminated_string_does_not_swallow_following_lines(run_source):
    # Regression for the STRING regex that used to match across newlines:
    # an unterminated string must NOT pair with a quote on a later line and
    # mask the error. Each unterminated quote must surface VBA_LEX004.
    src = 'Sub S()\n    a = "x\n    b = "y\n    c = 5\nEnd Sub\n'
    r = run_source(src)
    assert "VBA_LEX004" in _ids(r)
