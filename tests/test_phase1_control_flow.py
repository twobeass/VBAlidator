"""Direct AST tests for Phase-1 control-flow nodes.

The fixture-driven tests in test_compile_error_samples.py and
test_valid_samples.py also cover these constructs end-to-end. These tests
focus on the AST shape itself so any future parser regression is caught
before it becomes an analyzer regression.
"""
from __future__ import annotations

import pytest

from src.lexer import Lexer
from src.parser import (
    CaseClauseNode,
    DoNode,
    EraseNode,
    ForNode,
    RedimNode,
    SelectNode,
    VBAParser,
)


def _parse(code: str):
    tokens = list(Lexer(code).tokenize())
    return VBAParser(tokens).parse_module()


def test_for_loop_body_captured():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim i As Long
    For i = 1 To 10
        i = i + 1
    Next i
End Sub
"""
    proc = _parse(code).procedures[0]
    fors = [n for n in proc.body if isinstance(n, ForNode)]
    assert len(fors) == 1
    f = fors[0]
    assert f.kind == "counter"
    assert f.body, "For body must contain at least one statement"


def test_for_each_body_captured():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim arr As Variant
    Dim x As Variant
    For Each x In arr
        x = x + 1
    Next x
End Sub
"""
    proc = _parse(code).procedures[0]
    fors = [n for n in proc.body if isinstance(n, ForNode)]
    assert len(fors) == 1 and fors[0].kind == "each"
    assert fors[0].body


def test_do_loop_body_captured_top_tested():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    Do While x < 10
        x = x + 1
    Loop
End Sub
"""
    proc = _parse(code).procedures[0]
    dos = [n for n in proc.body if isinstance(n, DoNode)]
    assert len(dos) == 1
    assert dos[0].condition_position == "top"
    assert dos[0].body


def test_do_loop_body_captured_bottom_tested():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
"""
    proc = _parse(code).procedures[0]
    dos = [n for n in proc.body if isinstance(n, DoNode)]
    assert len(dos) == 1
    assert dos[0].condition_position == "bottom"
    assert dos[0].condition_tokens, "Bottom-tested loop must capture condition tokens"


def test_while_wend_body_captured():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    While x < 10
        x = x + 1
    Wend
End Sub
"""
    proc = _parse(code).procedures[0]
    dos = [n for n in proc.body if isinstance(n, DoNode)]
    assert len(dos) == 1
    assert dos[0].kind == "while_wend"


def test_select_case_clauses_captured():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    Select Case x
        Case 1, 2, 3
            x = 0
        Case Is < 10
            x = 10
        Case 11 To 20
            x = 20
        Case Else
            x = -1
    End Select
End Sub
"""
    proc = _parse(code).procedures[0]
    sels = [n for n in proc.body if isinstance(n, SelectNode)]
    assert len(sels) == 1
    cases = sels[0].cases
    assert len(cases) == 4
    assert all(isinstance(c, CaseClauseNode) for c in cases)
    assert sum(1 for c in cases if c.is_else) == 1


def test_for_body_undefined_identifier_is_detected(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim i As Long
    For i = 1 To 10
        notDeclared = i
    Next i
End Sub
"""
    result = run_source(code)
    assert result.has_message_containing("notDeclared"), (
        f"Identifier inside For body must be analyzed. Errors: {result.messages!r}"
    )


def test_select_case_body_undefined_identifier_is_detected(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    Select Case x
        Case 1
            wrongVar = 1
    End Select
End Sub
"""
    result = run_source(code)
    assert result.has_message_containing("wrongVar"), (
        f"Identifier inside Select body must be analyzed. Errors: {result.messages!r}"
    )


def test_redim_node_parsed():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim arr() As Long
    ReDim Preserve arr(1 To 10)
End Sub
"""
    proc = _parse(code).procedures[0]
    redims = [n for n in proc.body if isinstance(n, RedimNode)]
    assert len(redims) == 1
    r = redims[0]
    assert r.preserve is True
    assert len(r.targets) == 1
    target = r.targets[0]
    name_token = target[0]
    assert name_token.value.lower() == "arr"
    # P2.6/#20: targets now include chain_tokens for dotted-target support.
    if len(target) >= 4:
        chain = target[3]
        assert [t.value.lower() for t in chain] == ["arr"]


def test_erase_node_parsed():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim arr() As Long
    Erase arr
End Sub
"""
    proc = _parse(code).procedures[0]
    erases = [n for n in proc.body if isinstance(n, EraseNode)]
    assert len(erases) == 1
    assert len(erases[0].targets) == 1
    assert erases[0].targets[0].value.lower() == "arr"


def test_string_suffix_resolves(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim s As String
    s = Left$("hello", 3)
    s = Right$(s, 2)
    s = Mid$(s, 1, 1)
End Sub
"""
    result = run_source(code)
    hard = [e for e in result.errors if e.get("severity", "error") == "error"]
    assert hard == [], (
        f"$-suffixed standard functions must resolve. Errors: {[e['message'] for e in hard]!r}"
    )


def test_bracket_identifier_tokenizes():
    """`[A1]` is VBA's foreign-name escape; lexer must tokenize it as one unit."""
    code = 'Set x = [A1]\n'
    tokens = list(Lexer(code).tokenize())
    types = [t.type for t in tokens if t.type not in ("NEWLINE", "EOF")]
    assert "BRACKET_IDENTIFIER" in types, f"Expected BRACKET_IDENTIFIER token, got {types!r}"


def test_unexpected_character_still_reports_error():
    """The Phase-0 MISMATCH→error pathway must still fire for genuinely
    invalid characters even though `$` and `[...]` are now accepted."""
    code = "Dim x As Long\nx = 1€\n"  # Euro sign
    lex = Lexer(code)
    list(lex.tokenize())
    assert any(e.char == "€" for e in lex.errors), (
        f"Euro sign must still surface as lexer error. Errors: {[e.char for e in lex.errors]!r}"
    )


@pytest.mark.timeout(5)
def test_standalone_end_statement_inside_if_block_does_not_hang():
    """Regression: `parse_block` previously prefix-matched the standalone
    `End` *statement* (program terminator) against multi-word block markers
    like `End If`, returning early and trapping callers in an infinite loop
    once a label followed. Reproduced by stdVBA's `stdError.cls::Raise`."""
    code = """
Attribute VB_Name = "M"
Public Function Raise() As Long
    If True Then
      End
    End If
    Exit Function
ErrorOccurred:
    Raise = 1
End Function
"""
    module = _parse(code)
    # Must reach the procedure and capture both the standalone End and
    # the label-tagged recovery branch without hanging.
    assert len(module.procedures) == 1
    assert module.procedures[0].body, "Function body must not be empty"
