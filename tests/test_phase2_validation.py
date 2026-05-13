"""Direct tests for Phase 2.1 (jumps), 2.2 (Set/Let), 2.3 (property arity)."""
from __future__ import annotations


def test_undefined_goto_reported(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    GoTo NoSuch
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA201" for e in result.errors), (
        f"GoTo to undefined label must produce VBA201. Errors: {result.errors!r}"
    )


def test_on_error_goto_zero_is_accepted(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    On Error GoTo 0
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA201" for e in result.errors), (
        f"`On Error GoTo 0` is the standard reset and must not flag. Errors: {result.errors!r}"
    )


def test_on_error_goto_minus_one_is_accepted(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    On Error GoTo -1
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA201" for e in result.errors)


def test_resume_next_is_accepted(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    On Error Resume Next
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA201" for e in result.errors)


def test_resume_to_undefined_label(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    On Error GoTo Handler
    Exit Sub
Handler:
    Resume Nope
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA201" and "Resume" in e.get("message", "")
        for e in result.errors
    ), f"Resume to unknown label must produce VBA201. Errors: {result.errors!r}"


def test_label_inside_for_body_is_visible(run_source):
    """Labels declared anywhere in the procedure must be reachable for jumps."""
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim i As Long
    For i = 1 To 10
InnerLabel:
        i = i + 1
    Next i
    GoTo InnerLabel
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA201" for e in result.errors), (
        f"Label in nested block must be visible for outer GoTo. Errors: {result.errors!r}"
    )


def test_set_on_scalar_is_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim s As String
    Set s = "x"
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA210" for e in result.errors), (
        f"Set on scalar must produce VBA210. Errors: {result.errors!r}"
    )


def test_missing_set_on_object_is_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim col As Object
    col = CreateObject("Scripting.Dictionary")
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA211" for e in result.errors), (
        f"Object assignment without Set must produce VBA211. Errors: {result.errors!r}"
    )


def test_variant_does_not_trigger_set_let(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim v As Variant
    v = 1
    Set v = Nothing
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") not in ("VBA210", "VBA211") for e in result.errors), (
        f"Variant LHS must accept either form. Errors: {result.errors!r}"
    )


def test_dotted_lhs_does_not_trigger_set_let(run_source):
    """`obj.prop = expr` is intentionally skipped to avoid false positives
    on properties whose member type is not in the standard model."""
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    col.SomeMember = 1
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") not in ("VBA210", "VBA211") for e in result.errors), (
        f"Dotted LHS must not be flagged. Errors: {result.errors!r}"
    )


def test_property_arity_mismatch():
    """Property Let with wrong arg count must surface VBA222."""
    from src.analyzer import Analyzer
    from src.config import Config
    from src.lexer import Lexer
    from src.parser import VBAParser

    code = """VERSION 1.0 CLASS
Attribute VB_Name = "WrongLet"
Public Property Get Foo() As Long
    Foo = 1
End Property
Public Property Let Foo(ByVal a As Long, ByVal b As Long)
End Property
"""
    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="WrongLet.cls")
    module = parser.parse_module()
    module.module_type = "Class"
    module.filename = "WrongLet.cls"

    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    assert any(e.get("rule_id") == "VBA222" for e in errors), (
        f"Property Let with wrong arity must produce VBA222. Errors: {errors!r}"
    )


def test_property_let_with_object_param_flagged():
    """Property Let whose value parameter is object-typed should suggest Property Set."""
    from src.analyzer import Analyzer
    from src.config import Config
    from src.lexer import Lexer
    from src.parser import VBAParser

    code = """VERSION 1.0 CLASS
Attribute VB_Name = "LetObject"
Public Property Get Item() As Object
    Set Item = Nothing
End Property
Public Property Let Item(ByVal RHS As Object)
End Property
"""
    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="LetObject.cls")
    module = parser.parse_module()
    module.module_type = "Class"
    module.filename = "LetObject.cls"

    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    assert any(e.get("rule_id") == "VBA224" for e in errors), (
        f"Property Let on Object param should produce VBA224. Errors: {errors!r}"
    )


def test_valid_property_passes():
    from src.analyzer import Analyzer
    from src.config import Config
    from src.lexer import Lexer
    from src.parser import VBAParser

    code = """VERSION 1.0 CLASS
Attribute VB_Name = "GoodProp"
Public Property Get Name() As String
    Name = "x"
End Property
Public Property Let Name(ByVal RHS As String)
End Property
"""
    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="GoodProp.cls")
    module = parser.parse_module()
    module.module_type = "Class"
    module.filename = "GoodProp.cls"

    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    prop_errors = [e for e in errors if e.get("rule_id", "").startswith("VBA22")]
    assert not prop_errors, (
        f"Well-formed Get/Let pair must produce no property errors. Got: {prop_errors!r}"
    )
