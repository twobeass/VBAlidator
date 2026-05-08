"""Regression tests for the single-line `If` parsing bug.

Before the fix, `parse_if_stmt` for single-line Ifs collected condition and
body tokens but returned ``None``, which silently discarded them. The
analyzer therefore never inspected single-line If statements — undefined
identifiers, type mismatches and argument-count problems on those lines
went unreported.
"""
from __future__ import annotations

from src.parser import IfNode, ProcedureNode, StatementNode, VBAParser
from src.lexer import Lexer


def _parse_module_for(code: str):
    tokens = list(Lexer(code).tokenize())
    return VBAParser(tokens).parse_module()


def _find_if_in_proc(proc: ProcedureNode) -> IfNode:
    for node in proc.body:
        if isinstance(node, IfNode):
            return node
    raise AssertionError(f"No IfNode found in procedure {proc.name}")


def test_single_line_if_returns_ifnode():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    If x = 1 Then x = 2
End Sub
"""
    module = _parse_module_for(code)
    proc = module.procedures[0]
    if_node = _find_if_in_proc(proc)
    assert isinstance(if_node, IfNode)
    assert if_node.true_block, "Single-line If body must be captured"
    assert if_node.condition_tokens, "Condition tokens must be present"


def test_single_line_if_else_branch_captured():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    If x = 1 Then x = 2 Else x = 3
End Sub
"""
    module = _parse_module_for(code)
    proc = module.procedures[0]
    if_node = _find_if_in_proc(proc)
    assert if_node.else_block is not None, "Single-line If/Else must capture else branch"
    assert if_node.else_block, "Else branch must contain at least one statement"


def test_single_line_if_chained_colon_statements():
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    If x = 1 Then x = 2 : x = 3 : x = 4
End Sub
"""
    module = _parse_module_for(code)
    proc = module.procedures[0]
    if_node = _find_if_in_proc(proc)
    # Each colon-separated statement should be its own StatementNode.
    assert len(if_node.true_block) >= 3, (
        f"Expected ≥3 colon-separated statements in body, got {len(if_node.true_block)}"
    )
    for n in if_node.true_block:
        assert isinstance(n, StatementNode)


def test_single_line_if_body_undefined_identifier_is_detected(run_source):
    """The end-to-end regression: an undefined identifier inside a single-line
    If body must now be flagged by the analyzer (it was silently dropped before).
    """
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    If x = 1 Then notDeclaredVariable = 2
End Sub
"""
    result = run_source(code)
    assert result.has_message_containing("notDeclaredVariable"), (
        f"Undefined identifier inside single-line If body must be reported. "
        f"Errors: {result.messages!r}"
    )
