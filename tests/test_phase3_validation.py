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


# ---- VBA360 / VBA361 — statement placement (P3.5) ---------------------


def test_type_inside_procedure_flagged(run_source):
    """`Type ... End Type` is module-only. Declaring it inside a Sub
    is a hard VBA compile error."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Type Point\n"
        "        x As Long\n"
        "    End Type\n"
        "End Sub\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA360" for e in result.errors), result.errors


def test_enum_inside_procedure_flagged(run_source):
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Enum Colors\n"
        "        Red = 1\n"
        "    End Enum\n"
        "End Sub\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA360" for e in result.errors), result.errors


def test_declare_inside_procedure_flagged(run_source):
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        '    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'
        "End Sub\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA360" for e in result.errors), result.errors


def test_module_level_type_passes(run_source):
    """Symmetric guard: same `Type ... End Type` at module scope is fine."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "\n"
        "Type Point\n"
        "    x As Long\n"
        "End Type\n"
        "\n"
        "Sub S()\n"
        "    Dim p As Point\n"
        "End Sub\n"
    )
    result = run_source(code)
    assert all(e.get("rule_id") != "VBA360" for e in result.errors), result.errors


def test_executable_at_module_level_flagged(run_source):
    """Free-standing `Debug.Print` at module top is illegal in VBA."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        'Debug.Print "illegal at module level"\n'
        "Sub S(): End Sub\n"
    )
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA361" for e in result.errors), result.errors


def test_class_module_header_does_not_trigger_vba361(run_source):
    """The `.cls` export header (VERSION + BEGIN…END attribute block) is
    a serialisation artefact, not VBA code — must not trip VBA361 even
    though the BEGIN block contains assignment-like lines."""
    code = (
        "VERSION 1.0 CLASS\n"
        "BEGIN\n"
        "  MultiUse = -1  'True\n"
        "  Persistable = 0\n"
        "  DataBindingBehavior = 0\n"
        "END\n"
        'Attribute VB_Name = "MyClass"\n'
        "Option Explicit\n"
        "Public Sub Greet()\n"
        "End Sub\n"
    )
    result = run_source(code, module_type="Class")
    bad = [e for e in result.errors if e.get("rule_id") == "VBA361"]
    assert not bad, f"`.cls` header block must not produce VBA361. Got: {bad!r}"


def test_declare_any_byref_accepts_any_concrete_type(run_source):
    """`As Any` in a Declare statement is VBA's universally-compatible
    sentinel — the analyzer must not flag callers passing concrete types."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

Sub S()
    Dim b As Byte
    Dim l As Long
    Dim s As String
    CopyMemory b, l, 1
    CopyMemory s, b, 1
End Sub
"""
    result = run_source(code)
    mismatches = [e for e in result.errors if "ByRef argument type mismatch" in (e.get("message") or "")]
    assert not mismatches, (
        f"`As Any` ByRef params must accept any concrete type. Got: {mismatches!r}"
    )


def test_declare_any_array_byref_accepts_concrete_array(run_source):
    """`As Any()` is the array form of the `Any` sentinel — same contract,
    must accept any concrete array type (regression for VbTrickTimer's
    `DupArray` declaration)."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Declare PtrSafe Sub DupArray Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination() As Any, ByRef pSA As Any, _
    Optional ByVal Length As LongPtr = 8)

Sub S()
    Dim bData() As Byte
    Dim tSAMap As Long
    DupArray bData, VarPtr(tSAMap)
End Sub
"""
    result = run_source(code)
    mismatches = [e for e in result.errors if "ByRef argument type mismatch" in (e.get("message") or "")]
    assert not mismatches, (
        f"`As Any()` ByRef params must accept any concrete array type. Got: {mismatches!r}"
    )


def test_array_accepts_unlimited_args(run_source):
    """VBA's `Array()` is a ParamArray — callers may pass arbitrarily
    many values. Regression for stdAcc's `CreateLookupDict(Array(...))`
    populated with 122 name/value pairs."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim v As Variant\n"
        "    v = Array(" + ", ".join(str(i) for i in range(150)) + ")\n"
        "End Sub\n"
    )
    result = run_source(code)
    arg_errors = [e for e in result.errors if "Argument count mismatch for 'Array'" in (e.get("message") or "")]
    assert not arg_errors, (
        f"`Array()` is a ParamArray — must accept any arg count. Got: {arg_errors!r}"
    )


def test_choose_and_switch_accept_unlimited_args(run_source):
    """Same as Array — `Choose(idx, c1, c2, ...)` and
    `Switch(cond1, val1, cond2, val2, ...)` are ParamArray-style."""
    code = (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n"
        "Sub S()\n"
        "    Dim v As Variant\n"
        "    v = Choose(1, " + ", ".join(f'"v{i}"' for i in range(120)) + ")\n"
        "    v = Switch(" + ", ".join(f'True, "v{i}"' for i in range(60)) + ")\n"
        "End Sub\n"
    )
    result = run_source(code)
    arg_errors = [e for e in result.errors if "Argument count mismatch for" in (e.get("message") or "")]
    assert not arg_errors, (
        f"`Choose`/`Switch` are ParamArray — must accept any arg count. "
        f"Got: {arg_errors!r}"
    )


def test_enum_arg_compatible_with_long_byref_param(run_source):
    """VBA accepts `Dim x As MyEnum: Inner x` against
    `Sub Inner(ByRef p As Long)` because enums are Long under the
    hood. Regression for stdVBA-style code."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Public Enum ParentType
    ptStandard = 1
End Enum

Sub Inner(ByRef p As Long)
End Sub

Sub Caller()
    Dim x As ParentType
    Inner x
End Sub
"""
    result = run_source(code)
    mismatches = [e for e in result.errors if "ByRef argument type mismatch" in (e.get("message") or "")]
    assert not mismatches, (
        f"Enum-typed argument must accept Long ByRef param. Got: {mismatches!r}"
    )


def test_long_arg_compatible_with_enum_byref_param(run_source):
    """Reverse direction: `Dim x As Long: Inner x` against
    `Sub Inner(ByRef p As MyEnum)` — also valid VBA, and the dominant
    pattern in stdVBA's `stdLambda.cls` opcode emitter."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Public Enum IInstruction
    iOp1 = 1
    iOp2 = 2
End Enum

Sub Emit(ByRef kInstruction As IInstruction)
End Sub

Sub Caller()
    Dim x As Long
    x = 1
    Emit x
End Sub
"""
    result = run_source(code)
    mismatches = [e for e in result.errors if "ByRef argument type mismatch" in (e.get("message") or "")]
    assert not mismatches, (
        f"Long-typed argument must accept Enum ByRef param. Got: {mismatches!r}"
    )


def test_array_passed_with_empty_parens_keeps_array_type(run_source):
    """`arr()` with empty parens is VBA's explicit pass-whole-array syntax,
    not an indexed element access — the type must stay the array, not
    collapse to the element. Regression for VbTrickTimer's
    `FindSignature(bData(), bTemplate(), bMask())`."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Function FindSignature(ByRef bData() As Byte, ByRef bSign() As Byte) As Long
    FindSignature = 0
End Function

Sub S()
    Dim bData() As Byte
    Dim bSign() As Byte
    Dim lIndex As Long
    lIndex = FindSignature(bData(), bSign())
End Sub
"""
    result = run_source(code)
    mismatches = [e for e in result.errors if "ByRef argument type mismatch" in (e.get("message") or "")]
    assert not mismatches, (
        f"Empty-paren `arr()` must keep the array type. Got: {mismatches!r}"
    )


def test_subcall_inside_expression_does_not_steal_args(run_source):
    """Sub-style implicit calls (`MsgBox "hi"`, `Debug.Print 1, 2`) are
    legal at statement level but NOT inside an expression. The classic
    failure: `Round(Timer - t, 3)` was attributing `Timer - t, 3` to
    `Timer` (a 0-arg built-in) instead of letting the comma split
    `Round`'s args. Regression for VBA-MemoryTools `DemoLibMemory.bas`."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim t As Double
    t = Timer
    Debug.Print "elapsed: " & VBA.Round(Timer - t, 3)
    Debug.Print "elapsed: " & Round(Timer - t, 3)
End Sub
"""
    result = run_source(code)
    hard = [e for e in result.errors if e.get("severity", "error") == "error"]
    assert not hard, (
        f"`Round(Timer - t, 3)` must split args at the comma, not interpret "
        f"`Timer` as a 2-arg call. Got: {hard!r}"
    )


def test_statement_level_subcall_still_works(run_source):
    """Sub-style implicit calls at the start of a statement must keep
    working — the fix above must not regress `MsgBox "Hi"` /
    `Debug.Print x, y` style code."""
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub Greet(s As String, n As Long)
End Sub

Sub S()
    Dim n As Long
    n = 3
    Debug.Print "hi", n
    Greet "world", n
End Sub
"""
    result = run_source(code)
    hard = [e for e in result.errors if e.get("severity", "error") == "error"]
    assert not hard, (
        f"Statement-level Sub-style calls must stay valid. Got: {hard!r}"
    )


def test_enum_member_is_valid_const_initialiser(run_source):
    """Source-declared `Enum` members are compile-time Long constants in
    VBA — they may appear on the RHS of `Const X = MyEnum.Member * 4`.
    Regression for VBA-MemoryTools' `LibMemory.bas::EmptyArray`."""
    code = """
Attribute VB_Name = "M"
Option Explicit

Public Enum FADF
    FADF_AUTO = &H1
    FADF_HAVEVARTYPE = &H80
End Enum

Sub S()
    Const fFeaturesHi As Long = FADF_HAVEVARTYPE * &H10000
End Sub
"""
    result = run_source(code)
    const_errors = [e for e in result.errors if "non-constant" in (e.get("message") or "")]
    assert not const_errors, (
        f"Enum members must count as constant on the RHS of `Const X = …`. "
        f"Got: {const_errors!r}"
    )
