"""Direct tests for Phase 3.1–3.4, 3.6 (Implements / Events / PtrSafe /
Enum-uniqueness / Option Explicit warning)."""
from __future__ import annotations

from src.analyzer import Analyzer
from src.config import Config
from src.lexer import Lexer
from src.parser import VBAParser


# ---- P3.3 PtrSafe -------------------------------------------------------


def test_ptrsafe_required_on_64bit(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA300" for e in result.errors), (
        f"Missing PtrSafe must produce VBA300. Errors: {result.errors!r}"
    )


def test_ptrsafe_present_is_accepted(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA300" for e in result.errors), (
        f"PtrSafe-tagged Declare must not flag VBA300. Errors: {result.errors!r}"
    )


def test_ptrsafe_can_be_disabled_via_define():
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
"""
    config = Config()
    config.definitions["WIN64"] = False
    config.definitions["VBA7"] = False

    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="M.bas")
    module = parser.parse_module()
    module.module_type = "Module"
    module.filename = "M.bas"

    analyzer = Analyzer(config)
    analyzer.add_module(module)
    errors = analyzer.analyze()
    assert all(e.get("rule_id") != "VBA300" for e in errors), (
        "When WIN64=False the PtrSafe rule must not fire."
    )


# ---- P3.4 Enum uniqueness ----------------------------------------------


def test_duplicate_enum_member_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Public Enum Colors
    Red = 1
    Green = 2
    Red = 3
End Enum
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA310" for e in result.errors), (
        f"Duplicate enum member must produce VBA310. Errors: {result.errors!r}"
    )


def test_unique_enum_members_pass(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Public Enum Status
    Ready = 0
    Busy = 1
    Done = 2
End Enum
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA310" for e in result.errors)


# ---- P3.6 Option Explicit ----------------------------------------------


def test_missing_option_explicit_warns(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
End Sub
"""
    result = run_source(code)
    matched = [e for e in result.errors if e.get("rule_id") == "VBA320"]
    assert matched, f"Missing Option Explicit must warn. Errors: {result.errors!r}"
    assert matched[0].get("severity") == "warning"


def test_option_explicit_present_no_warning(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
End Sub
"""
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA320" for e in result.errors)


# ---- P3.2 RaiseEvent ----------------------------------------------------


def test_raise_event_unknown_target_flagged(run_source):
    code = """
VERSION 1.0 CLASS
Attribute VB_Name = "C"
Option Explicit
Sub Trigger()
    RaiseEvent NotDeclared
End Sub
"""
    result = run_source(code, module_type="Class")
    assert any(e.get("rule_id") == "VBA340" for e in result.errors), (
        f"RaiseEvent with no matching Event must produce VBA340. Errors: {result.errors!r}"
    )


def test_raise_event_arity_mismatch_flagged():
    code = """
VERSION 1.0 CLASS
Attribute VB_Name = "C"
Option Explicit
Public Event Changed(ByVal name As String, ByVal value As Variant)

Sub Trigger()
    RaiseEvent Changed("only-one-arg")
End Sub
"""
    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="C.cls")
    module = parser.parse_module()
    module.module_type = "Class"
    module.filename = "C.cls"
    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    assert any(e.get("rule_id") == "VBA341" for e in errors), (
        f"RaiseEvent arity mismatch must produce VBA341. Errors: {errors!r}"
    )


def test_raise_event_correct_arity_passes():
    code = """
VERSION 1.0 CLASS
Attribute VB_Name = "C"
Option Explicit
Public Event Changed(ByVal name As String, ByVal value As Variant)

Sub Trigger()
    RaiseEvent Changed("a", 1)
End Sub
"""
    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="C.cls")
    module = parser.parse_module()
    module.module_type = "Class"
    module.filename = "C.cls"
    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    bad = [e for e in errors if e.get("rule_id", "").startswith("VBA34")]
    assert not bad, f"Well-formed RaiseEvent must not flag. Errors: {bad!r}"


# ---- P3.1 Implements ----------------------------------------------------


def test_implements_missing_method_flagged():
    iface = """
VERSION 1.0 CLASS
Attribute VB_Name = "IShape"
Option Explicit
Public Sub Draw()
End Sub
Public Function Area() As Double
End Function
"""
    impl = """
VERSION 1.0 CLASS
Attribute VB_Name = "Square"
Option Explicit
Implements IShape

' Only IShape_Draw is provided — IShape_Area is missing.
Public Sub IShape_Draw()
End Sub
"""
    analyzer = Analyzer(Config())
    for src, name, mtype in [(iface, "IShape", "Class"), (impl, "Square", "Class")]:
        tokens = list(Lexer(src).tokenize())
        module = VBAParser(tokens, filename=f"{name}.cls").parse_module()
        module.module_type = mtype
        module.filename = f"{name}.cls"
        analyzer.add_module(module)
    errors = analyzer.analyze()
    assert any(e.get("rule_id") == "VBA330" for e in errors), (
        f"Missing implementing method must produce VBA330. Errors: {errors!r}"
    )


def test_implements_complete_passes():
    iface = """
VERSION 1.0 CLASS
Attribute VB_Name = "IShape"
Option Explicit
Public Sub Draw()
End Sub
"""
    impl = """
VERSION 1.0 CLASS
Attribute VB_Name = "Square"
Option Explicit
Implements IShape

Public Sub IShape_Draw()
End Sub
"""
    analyzer = Analyzer(Config())
    for src, name, mtype in [(iface, "IShape", "Class"), (impl, "Square", "Class")]:
        tokens = list(Lexer(src).tokenize())
        module = VBAParser(tokens, filename=f"{name}.cls").parse_module()
        module.module_type = mtype
        module.filename = f"{name}.cls"
        analyzer.add_module(module)
    errors = analyzer.analyze()
    bad = [e for e in errors if e.get("rule_id") == "VBA330"]
    assert not bad, f"Complete implementation must not flag. Errors: {bad!r}"


# ---- Lexer / preprocessor cleanup --------------------------------------


def test_preprocessor_is_case_insensitive(run_source):
    """`#If Vba7 Then` (mixed case) used to evaluate to False because
    the lookup was case-sensitive against the upper-cased defines map.
    """
    code = """
Attribute VB_Name = "M"
Option Explicit

#If Vba7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
"""
    result = run_source(code)
    # Either branch is fine, but the chosen one must not emit VBA300
    # (missing PtrSafe). Pre-fix, the Else branch was selected and VBA300
    # fired even though VBA7 was in the defaults.
    assert all(e.get("rule_id") != "VBA300" for e in result.errors), (
        f"Vba7 (mixed-case) should be a truthy define. Errors: {result.errors!r}"
    )


def test_currency_literal_lex():
    """`50023612.1134@` is a Currency literal, not garbage."""
    code = "x = 50023612.1134@\n"
    lex = Lexer(code)
    list(lex.tokenize())
    assert not lex.errors, f"Currency literal must lex cleanly. Errors: {lex.errors!r}"


# ---- VBA350 — End Sub/Function/Property terminator mismatch -----------


def test_end_sub_closing_function_flagged(run_source):
    """A Function closed with `End Sub` (or vice versa) must surface
    VBA350. VBE rejects this at compile time; common AI-generator slip
    after refactoring a signature without updating the terminator."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Function ComputeValue() As Long\n"
        "    ComputeValue = 42\n"
        "End Sub\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA350" for e in result.errors), (
        f"`End Sub` closing a Function must produce VBA350. "
        f"Errors: {result.errors!r}"
    )


def test_end_function_closing_sub_flagged(run_source):
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub DoStuff()\n"
        "    Debug.Print 1\n"
        "End Function\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA350" for e in result.errors), result.errors


def test_matching_end_terminator_is_clean(run_source):
    """Guard against over-detection: properly matched terminators must
    not flag VBA350."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub A()\n"
        "End Sub\n"
        "Function B() As Long\n"
        "    B = 1\n"
        "End Function\n"
        "Property Get C() As Long\n"
        "    C = 1\n"
        "End Property\n"
    )
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA350" for e in result.errors), result.errors
