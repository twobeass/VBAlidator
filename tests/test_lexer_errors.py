"""Regression tests for the Lexer MISMATCH-token hardening.

Previously, characters that did not match any token spec were silently
dropped (``continue``) which produced a partial token stream and masked
encoding/typo issues. They are now collected on the lexer and surfaced via
the analyzer's error list with a stable rule_id.
"""
from __future__ import annotations

from src.lexer import Lexer, LexerError


def test_lexer_records_unexpected_character():
    # Use € (Euro) — never a valid VBA token. Keep `@` out of the test
    # because numeric type-suffix `1@` (Currency) is legal as of Phase 3.
    code = "Dim x As Long\nx = 1€\n"
    lex = Lexer(code)
    list(lex.tokenize())
    assert lex.errors, "Lexer should record unexpected character '€'"
    assert all(isinstance(e, LexerError) for e in lex.errors)
    assert any(e.char == "€" for e in lex.errors)


def test_lexer_error_to_dict_shape():
    err = LexerError("€", line=2, column=7)
    payload = err.to_dict(filename="Module1.bas")
    assert payload["file"] == "Module1.bas"
    assert payload["line"] == 2
    assert payload["column"] == 7
    assert payload["rule_id"] == "VBA_LEX001"
    assert payload["severity"] == "error"
    assert "€" in payload["message"]


def test_lexer_errors_surface_through_pipeline(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    x = 1€
End Sub
"""
    result = run_source(code)
    assert result.lexer_errors, "Pipeline should collect lexer errors"
    assert any(e.get("rule_id") == "VBA_LEX001" for e in result.errors), (
        f"Lexer errors must be surfaced as analyzer issues. Got: {result.errors!r}"
    )


def test_file_number_syntax_not_a_lexer_error(run_source):
    """VBA's `#1` / `#2` file-number argument used by Open / Print / Put /
    Close I/O statements must tokenize cleanly. Previously it fell through
    to the MISMATCH path and produced VBA_LEX001 false positives.
    """
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub WriteFile(path As String, body As String)
    Open path For Binary As #1
    Put #1, , body
    Close #1
End Sub
"""
    result = run_source(code)
    lex_errs = [e for e in result.errors if e.get("rule_id") == "VBA_LEX001"]
    assert not lex_errs, (
        f"File-number `#N` must not trigger lexer errors. Got: {lex_errs!r}"
    )


def test_clean_source_has_no_lexer_errors(run_source):
    code = """
Attribute VB_Name = "M"
Sub Clean()
    Dim x As Long
    x = 1
End Sub
"""
    result = run_source(code)
    assert result.lexer_errors == [], (
        f"Clean source must produce no lexer errors. Got: {result.lexer_errors!r}"
    )
