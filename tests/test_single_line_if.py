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


# ---- Colon-as-statement-separator inside a Sub body (Iter-2 regression) ----


def test_single_line_colon_separator_does_not_swallow_undefined(run_source):
    """`Sub S(): Dim x: x = UndefinedName: End Sub` must NOT silently
    drop the body. Pre-fix `analyze_expression_info` treated *every*
    `IDENT :` pair anywhere in the token stream as a label definition
    (skipping the lookup), so `UndefinedName` was hidden purely because
    a statement-separator `:` happened to follow it.

    A label-definition only exists at the start of a statement
    (token-index 0). Mid-stream colons are statement separators and
    must not suppress identifier resolution."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(): Dim x As Long: x = UndefinedName: End Sub\n"
    )
    result = run_source(code)
    # `run_source` returns the raw analyzer error list; rule_id is set
    # only by `normalize_issues` in the api layer. Match the legacy
    # message pattern instead — `normalize_issues` would translate this
    # into a VBA001 (verified separately by test_phase4_api.py).
    assert any(
        "Undefined identifier 'UndefinedName'" in e.get("message", "")
        for e in result.errors
    ), (
        f"Colon-separated single-line Sub body must surface the "
        f"undefined-identifier error for UndefinedName. "
        f"Errors: {result.errors!r}"
    )


def test_single_line_colon_separator_matches_multiline_behaviour(run_source):
    """Identical code in colon form vs. multiline form must produce the
    same set of `rule_id`s. Tightens the regression: any future divergence
    between the two normalisations is a bug."""
    multi = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim x As Long\n"
        "    x = UndefinedName\n"
        "End Sub\n"
    )
    colon = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S(): Dim x As Long: x = UndefinedName: End Sub\n"
    )
    multi_ids = sorted(e.get("rule_id") for e in run_source(multi).errors)
    colon_ids = sorted(e.get("rule_id") for e in run_source(colon).errors)
    assert multi_ids == colon_ids, (
        f"colon-form must report the same errors as the multiline form. "
        f"multi={multi_ids!r}  colon={colon_ids!r}"
    )


def test_label_at_statement_start_still_recognised(run_source):
    """The fix narrows label-detection to position 0 of a statement's
    token list. Guard against the over-fix: a real `LabelName:` at the
    start of a statement must continue to be recognised so jump-target
    validation (VBA201) keeps working."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    GoTo Skip\n"
        "    Debug.Print \"unreached\"\n"
        "Skip:\n"
        "    Debug.Print \"reached\"\n"
        "End Sub\n"
    )
    result = run_source(code)
    # No VBA201: the label exists.
    assert all(e.get("rule_id") != "VBA201" for e in result.errors), (
        f"Real label-definition must not become VBA201 after the colon-"
        f"detection fix. Errors: {result.errors!r}"
    )
