"""Single source of truth for every VBAlidator rule.

The registry is consumed by:
- src/reporting.py    → default severity / category for new findings
- tools/generate_rule_docs.py → docs/rules/<id>.md generation
- tests/test_rules_registry.py → coverage check (every rule_id emitted
  by the analyzer must have an entry here)

Adding a new rule
-----------------
1. Add an entry below with rule_id, title, severity, category, description,
   fail_example, ok_example and fix_hint.
2. Use the rule_id when calling `self.errors.append({...})` in analyzer.
3. Run `python tools/generate_rule_docs.py` to refresh the catalogue.

The IDs follow the roadmap numbering:
- VBA000–VBA099  legacy / pre-rule-id analyzer findings (mapped from
                 message patterns in src/reporting.py)
- VBA100–VBA199  Phase 1 (control flow: For/Do/While/Select/ReDim/Erase)
- VBA200–VBA299  Phase 2 (jumps / Set-Let / properties / operators /
                 const / fixed-string / etc.)
- VBA300–VBA399  Phase 3 (PtrSafe / enum / option / interface / events)
- VBA_LEX*       lexer-level diagnostics
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Rule:
    rule_id: str
    title: str
    severity: str          # error | warning | info
    category: str
    description: str
    fail_example: str = ""
    ok_example: str = ""
    fix_hint: str = ""
    phase: str = ""        # e.g. "0", "1", "2.4", "3.3"
    tags: tuple = field(default_factory=tuple)


# ---- Registry -----------------------------------------------------------

_RULES: list[Rule] = [
    # -- Lexer ---------------------------------------------------------
    Rule(
        rule_id="VBA_LEX001",
        title="Unexpected character",
        severity="error",
        category="lexer",
        phase="0",
        description=(
            "The lexer encountered a character that is not part of any "
            "valid VBA token. Often an encoding mistake (smart quotes, "
            "Euro sign, …) introduced by copy-paste."
        ),
        fail_example="x = 1€",
        ok_example="x = 1",
        fix_hint="Replace the offending character with its ASCII equivalent or remove it.",
    ),
    Rule(
        rule_id="VBA_LEX002",
        title="Invalid date literal",
        severity="error",
        category="lexer",
        phase="2.7",
        description=(
            "A `#…#` literal does not parse as any of the recognised VBA "
            "date / time formats, or its month / day / hour fields are "
            "out of range."
        ),
        fail_example="d = #2025-13-45#",
        ok_example="d = #2025-01-15#",
        fix_hint="Use m/d/y, yyyy-mm-dd, d-mmm-y, or 'MMMM d, y' format with valid date components.",
    ),

    # -- Round-trip verification (Phase 4.5) -------------------------
    Rule(
        rule_id="VBA_RT000",
        title="Round-trip verification unavailable",
        severity="info",
        category="roundtrip",
        phase="4.5",
        description=(
            "The runtime could not even attempt a VBE round-trip — usually "
            "because we're not on Windows, pywin32 is missing, or Office "
            "is not installed. Static analysis remains the authoritative "
            "result; this is informational only."
        ),
        fix_hint=(
            "Install Microsoft Office and `pip install pywin32` to enable "
            "round-trip verification, or simply drop `--roundtrip` from "
            "the invocation."
        ),
    ),
    Rule(
        rule_id="VBA_RT001",
        title="VBE round-trip compile error",
        severity="compile_verified",
        category="roundtrip",
        phase="4.5",
        description=(
            "The actual VBE compiler refused the source. This is the "
            "strongest possible verdict — a real Office host has rejected "
            "the code, so the static analyser's pass / fail call is "
            "confirmed dynamically."
        ),
        fix_hint=(
            "Open the source in the VBE manually to see the full error "
            "message; the round-trip report includes the VBE description."
        ),
    ),
    Rule(
        rule_id="VBA_RT002",
        title="Round-trip verification inconclusive",
        severity="warning",
        category="roundtrip",
        phase="4.5",
        description=(
            "The runtime tried to drive the VBE compiler but no trigger "
            "succeeded — `VBProject.Compile()` is hidden on modern Office, "
            "and the probe-Sub via `Application.Run` either timed out or "
            "failed with an unrecognised description. Distinct from "
            "`VBA_RT000`: VBE *was* reachable, we just couldn't reach a "
            "verdict."
        ),
        fix_hint=(
            "See TODO.md §A2 for the open work on Strategy 3 (VBE menu-bar "
            "invocation). In the meantime: rely on the static analyser, "
            "which remains the authoritative answer."
        ),
    ),

    # -- Legacy analyzer findings ------------------------------------
    Rule(
        rule_id="VBA001",
        title="Undefined identifier",
        severity="error",
        category="name_resolution",
        phase="0",
        description="A symbol is referenced that is not declared in any visible scope.",
        fail_example="Sub S()\n    typo = 1\nEnd Sub",
        ok_example="Sub S()\n    Dim count As Long\n    count = 1\nEnd Sub",
        fix_hint="Declare the variable with `Dim` or fix the spelling.",
    ),
    Rule(
        rule_id="VBA002",
        title="Member not found",
        severity="error",
        category="member_access",
        phase="0",
        description="A `obj.Member` access references a member that is not part of the object's type.",
        fail_example="Dim r As Range\nr.NotAMember = 1",
        ok_example="Dim r As Range\nr.Value = 1",
        fix_hint="Check the type's members in the object browser; use `--host` to load the matching model.",
    ),
    Rule(
        rule_id="VBA003",
        title="Duplicate declaration",
        severity="error",
        category="declaration",
        phase="0",
        description="Two `Dim`/`Const` declarations define the same name in the same scope.",
        fail_example="Dim x As Long\nDim x As String",
        ok_example="Dim x As Long\nDim y As String",
        fix_hint="Rename one of them or remove the duplicate.",
    ),
    Rule(
        rule_id="VBA004",
        title="Invalid `.member` reference without With",
        severity="error",
        category="member_access",
        phase="0",
        description="A leading-dot member reference appears outside any `With` block.",
        fail_example=".Value = 1",
        ok_example="With r\n    .Value = 1\nEnd With",
        fix_hint="Wrap the access in a `With` block or qualify the receiver explicitly.",
    ),
    Rule(
        rule_id="VBA005",
        title="Expected Array or Procedure",
        severity="error",
        category="type",
        phase="0",
        description="A scalar variable is invoked with `(args)` as if it were a function or array.",
        fail_example="Dim i As Integer\ni(1)",
        ok_example="Dim arr() As Integer\nReDim arr(1 To 10)\narr(1) = 0",
        fix_hint="Declare the variable as an array or call the right procedure.",
    ),
    Rule(
        rule_id="VBA006",
        title="Argument count mismatch",
        severity="error",
        category="signature",
        phase="0",
        description="A call passes too few or too many arguments for the target procedure.",
        fail_example="MsgBox()  ' MsgBox needs at least 1 arg",
        ok_example='MsgBox "hello"',
        fix_hint="Match the procedure signature; check Optional / ParamArray markers.",
    ),
    Rule(
        rule_id="VBA007",
        title="ByRef argument type mismatch",
        severity="error",
        category="signature",
        phase="0",
        description="A `ByRef` parameter receives a variable whose type does not match the parameter type.",
        fail_example=(
            "Sub Inc(ByRef n As Long): n = n + 1: End Sub\n"
            "Dim s As String: Inc s"
        ),
        ok_example=(
            "Sub Inc(ByRef n As Long): n = n + 1: End Sub\n"
            "Dim n As Long: Inc n"
        ),
        fix_hint="Use a temporary variable of the matching type, or change the parameter to `ByVal` if a copy is acceptable.",
    ),
    Rule(
        rule_id="VBA008",
        title="Exit statement type mismatch",
        severity="error",
        category="control_flow",
        phase="0",
        description="`Exit Sub` used inside a Function (or vice versa).",
        fail_example="Function F() As Long\n    Exit Sub\nEnd Function",
        ok_example="Function F() As Long\n    Exit Function\nEnd Function",
        fix_hint="Use the matching `Exit` form for the procedure type.",
    ),
    Rule(
        rule_id="VBA009",
        title="Unreachable code",
        severity="warning",
        category="control_flow",
        phase="0",
        description="Code follows an unconditional `Exit` / `GoTo` / `End` and can never run.",
        fail_example="Sub S()\n    Exit Sub\n    Debug.Print 1\nEnd Sub",
        ok_example="Sub S()\n    Debug.Print 1\n    Exit Sub\nEnd Sub",
        fix_hint="Remove the dead code or move it before the unconditional jump.",
    ),
    Rule(
        rule_id="VBA010",
        title="Syntax error",
        severity="error",
        category="syntax",
        phase="0",
        description="Generic syntax error (unexpected `End X`, missing `Then`, stray block terminator …).",
        fail_example="If x > 0\n    Debug.Print x\nEnd If",
        ok_example="If x > 0 Then\n    Debug.Print x\nEnd If",
        fix_hint="Check the surrounding tokens — typical causes are missing `Then` or mismatched block keywords.",
    ),

    # -- Phase 1: control-flow ---------------------------------------
    Rule(
        rule_id="VBA101",
        title="ReDim target undefined",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`ReDim` is applied to a name that has not been declared as a dynamic array.",
        fail_example="ReDim notDeclared(1 To 10)",
        ok_example="Dim arr() As Long\nReDim arr(1 To 10)",
        fix_hint="Declare the array with `Dim arr() As Type` first.",
    ),
    Rule(
        rule_id="VBA102",
        title="ReDim target not a variable",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`ReDim` was applied to something other than an array variable (procedure, class, …).",
        fail_example="ReDim Foo(1 To 10)  ' Foo is a Sub",
        ok_example="Dim arr() As Long\nReDim arr(1 To 10)",
        fix_hint="Check the name resolves to a dynamic-array variable.",
    ),
    Rule(
        rule_id="VBA103",
        title="ReDim target is not a dynamic array",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`ReDim` requires its target to be declared as a dynamic array (`Dim x() As …`).",
        fail_example="Dim x As Long\nReDim x(1 To 10)",
        ok_example="Dim x() As Long\nReDim x(1 To 10)",
        fix_hint="Add empty parentheses to the `Dim` declaration.",
    ),
    Rule(
        rule_id="VBA104",
        title="Erase target undefined",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`Erase` is applied to a name that has not been declared.",
        fail_example="Erase notDeclared",
        ok_example="Dim arr() As Long\nReDim arr(1 To 5)\nErase arr",
        fix_hint="Declare the array variable first.",
    ),
    Rule(
        rule_id="VBA105",
        title="Erase target not a variable",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`Erase` was applied to something other than an array variable.",
        fail_example="Erase Foo  ' Foo is a Sub",
        ok_example="Dim arr() As Long\nErase arr",
        fix_hint="Check that the name refers to an array variable.",
    ),
    Rule(
        rule_id="VBA106",
        title="Erase target is not an array",
        severity="error",
        category="declaration",
        phase="1.4",
        description="`Erase` requires an array. Scalar variables cannot be erased.",
        fail_example="Dim s As String\nErase s",
        ok_example="Dim arr() As Long\nErase arr",
        fix_hint="Use `s = vbNullString` for a String, or declare an array.",
    ),

    # -- Phase 2.1: jumps -------------------------------------------
    Rule(
        rule_id="VBA201",
        title="Jump target is not a label",
        severity="error",
        category="jump",
        phase="2.1",
        description="A `GoTo` / `On Error GoTo` / `Resume` / `GoSub` references a label that is not declared in the procedure.",
        fail_example="Sub S()\n    GoTo NoSuch\nEnd Sub",
        ok_example="Sub S()\n    GoTo Skip\nSkip:\nEnd Sub",
        fix_hint="Declare the label or fix the spelling. Special forms `On Error GoTo 0`, `On Error GoTo -1`, and `On Error Resume Next` do not need a target.",
    ),

    # -- Phase 2.2: Set vs Let --------------------------------------
    Rule(
        rule_id="VBA210",
        title="`Set` used on a scalar target",
        severity="error",
        category="assignment",
        phase="2.2",
        description="`Set` is only valid for Object / Class / Variant references. Scalar types use plain `=` (or `Let`).",
        fail_example="Dim s As String\nSet s = \"hello\"",
        ok_example="Dim s As String\ns = \"hello\"",
        fix_hint="Drop `Set` (or use `Let s = …`).",
    ),
    Rule(
        rule_id="VBA211",
        title="Object assignment without `Set`",
        severity="error",
        category="assignment",
        phase="2.2",
        description="Assigning to an Object-typed variable without `Set` is a compile error in VBA.",
        fail_example="Dim col As Object\ncol = CreateObject(\"Scripting.Dictionary\")",
        ok_example="Dim col As Object\nSet col = CreateObject(\"Scripting.Dictionary\")",
        fix_hint="Prepend `Set` to the assignment.",
    ),

    # -- Phase 2.3: property arity ----------------------------------
    Rule(
        rule_id="VBA221",
        title="Property Let/Set has zero parameters",
        severity="error",
        category="property",
        phase="2.3",
        description="`Property Let`/`Property Set` must declare at least the value parameter.",
        fail_example="Property Let Name()\nEnd Property",
        ok_example="Property Let Name(ByVal RHS As String)\nEnd Property",
        fix_hint="Add the assigned-value parameter.",
    ),
    Rule(
        rule_id="VBA222",
        title="Property Let/Set arity disagrees with Get",
        severity="error",
        category="property",
        phase="2.3",
        description="`Property Let`/`Set` parameter count must equal `Property Get`'s parameter count + 1 (the value parameter).",
        fail_example=(
            "Property Get Foo() As Long: End Property\n"
            "Property Let Foo(ByVal a As Long, ByVal b As Long): End Property"
        ),
        ok_example=(
            "Property Get Foo() As Long: End Property\n"
            "Property Let Foo(ByVal RHS As Long): End Property"
        ),
        fix_hint="Add or remove parameters until Let/Set has Get-arg-count + 1 parameters.",
    ),
    Rule(
        rule_id="VBA223",
        title="Property Set value parameter is scalar (use Let)",
        severity="error",
        category="property",
        phase="2.3",
        description="`Property Set` is for Object-typed RHS values. Scalar types use `Property Let`.",
        fail_example="Property Set Name(ByVal RHS As String): End Property",
        ok_example="Property Let Name(ByVal RHS As String): End Property",
        fix_hint="Rename the accessor to `Property Let`.",
    ),
    Rule(
        rule_id="VBA224",
        title="Property Let value parameter is object (use Set)",
        severity="error",
        category="property",
        phase="2.3",
        description="`Property Let` is for scalar RHS values. Object types use `Property Set`.",
        fail_example="Property Let Item(ByVal RHS As Object): End Property",
        ok_example="Property Set Item(ByVal RHS As Object): End Property",
        fix_hint="Rename the accessor to `Property Set`.",
    ),

    # -- Phase 2.5: Const expression --------------------------------
    Rule(
        rule_id="VBA230",
        title="`Const` initialiser calls a function",
        severity="error",
        category="const_expression",
        phase="2.5",
        description="`Const` initialisers must be constant expressions — function calls are not allowed.",
        fail_example="Const X As Long = MsgBox(\"x\")",
        ok_example="Const X As Long = 42",
        fix_hint="Replace the call with a literal or another `Const` reference.",
    ),
    Rule(
        rule_id="VBA231",
        title="`Const` initialiser references a non-constant",
        severity="error",
        category="const_expression",
        phase="2.5",
        description="`Const` initialisers cannot reference variables — only literals, other constants, or enum members.",
        fail_example=(
            "Dim runtimeValue As Long\n"
            "Const X As Long = runtimeValue"
        ),
        ok_example=(
            "Const A As Long = 5\n"
            "Const X As Long = A + 1"
        ),
        fix_hint="Use a literal or another `Const` / Enum member as the initialiser.",
    ),

    # -- Phase 2.4: operator types ----------------------------------
    Rule(
        rule_id="VBA240",
        title="Arithmetic operator between string and numeric literal",
        severity="error",
        category="operator_type",
        phase="2.4",
        description=(
            "Operators like `-`, `*`, `/`, `\\`, `^`, `Mod` require numeric operands. "
            "Use `&` for string concatenation."
        ),
        fail_example='x = "abc" - 1',
        ok_example='x = "value: " & 1',
        fix_hint="Use `&` for concatenation. `+` is bidirectional in VBA but can silently coerce — prefer `&` for strings.",
    ),

    # -- Phase 2.8: fixed-length string -----------------------------
    Rule(
        rule_id="VBA250",
        title="Fixed-length String at procedure level",
        severity="error",
        category="declaration",
        phase="2.8",
        description="`Dim s As String * N` is only legal at module / UDT level, not inside a procedure.",
        fail_example=(
            "Sub S()\n"
            "    Dim s As String * 10\n"
            "End Sub"
        ),
        ok_example=(
            "Public name As String * 10\n"
            "Sub S(): End Sub"
        ),
        fix_hint="Move the declaration to module level, or use a regular `As String`.",
    ),

    # -- Phase 3.3: PtrSafe -----------------------------------------
    Rule(
        rule_id="VBA300",
        title="`Declare` missing `PtrSafe`",
        severity="error",
        category="platform",
        phase="3.3",
        description=(
            "On 64-bit Office (VBA7+) every `Declare` of a Win32 API entry "
            "must carry the `PtrSafe` attribute. Without it the compiler "
            "refuses to load the module."
        ),
        fail_example='Private Declare Function GetTickCount Lib "kernel32" () As Long',
        ok_example=(
            "#If VBA7 Then\n"
            "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\n"
            "#Else\n"
            "Private Declare Function GetTickCount Lib \"kernel32\" () As Long\n"
            "#End If"
        ),
        fix_hint="Add `PtrSafe` after `Declare`. For dual 32/64-bit support wrap in `#If VBA7 Then` / `#Else`.",
    ),

    # -- Phase 3.4: enum uniqueness ---------------------------------
    Rule(
        rule_id="VBA310",
        title="Duplicate Enum member name",
        severity="error",
        category="enum",
        phase="3.4",
        description="Within a single Enum block all member names must be unique.",
        fail_example=(
            "Public Enum Colors\n"
            "    Red = 1\n"
            "    Red = 2\n"
            "End Enum"
        ),
        ok_example=(
            "Public Enum Colors\n"
            "    Red = 1\n"
            "    Blue = 2\n"
            "End Enum"
        ),
        fix_hint="Rename one of the duplicate members.",
    ),

    # -- Phase 3.6: Option Explicit ---------------------------------
    Rule(
        rule_id="VBA320",
        title="Module is missing `Option Explicit`",
        severity="warning",
        category="style",
        phase="3.6",
        description=(
            "Without `Option Explicit` typo'd variable names silently "
            "create new Variant variables — the #1 source of typo-induced "
            "bugs in VBA, and a common AI-generation pitfall."
        ),
        fail_example="Sub S()\n    typo = 1\nEnd Sub",
        ok_example=(
            "Option Explicit\n"
            "Sub S()\n"
            "    Dim count As Long\n"
            "    count = 1\n"
            "End Sub"
        ),
        fix_hint="Add `Option Explicit` as the first non-comment line of the module.",
    ),

    # -- Phase 3.1: Implements --------------------------------------
    Rule(
        rule_id="VBA330",
        title="Class is missing an interface method",
        severity="error",
        category="interface",
        phase="3.1",
        description=(
            "When a class declares `Implements <Interface>`, every public "
            "Sub / Function / Property of that interface must have a "
            "matching `<Interface>_<Member>` method."
        ),
        fail_example=(
            "' IShape:\n"
            "Public Sub Draw(): End Sub\n"
            "' Square:\n"
            "Implements IShape\n"
            "' missing IShape_Draw"
        ),
        ok_example=(
            "' Square:\n"
            "Implements IShape\n"
            "Public Sub IShape_Draw(): End Sub"
        ),
        fix_hint="Add the missing `Interface_Member` methods.",
    ),

    # -- Phase 3.2: Events ------------------------------------------
    Rule(
        rule_id="VBA340",
        title="`RaiseEvent` target has no matching `Event` declaration",
        severity="error",
        category="events",
        phase="3.2",
        description="Events can only be raised from the class that declares them.",
        fail_example="Sub Trigger()\n    RaiseEvent NotDeclared\nEnd Sub",
        ok_example=(
            "Public Event Changed()\n"
            "Sub Trigger()\n"
            "    RaiseEvent Changed\n"
            "End Sub"
        ),
        fix_hint="Declare `Public Event <Name>(...)` at module level, or fix the event name.",
    ),
    Rule(
        rule_id="VBA341",
        title="`RaiseEvent` argument count mismatch",
        severity="error",
        category="events",
        phase="3.2",
        description="The number of arguments passed to `RaiseEvent` must match the event's parameter list.",
        fail_example=(
            "Public Event Changed(ByVal name As String, ByVal value As Variant)\n"
            "Sub Trigger()\n"
            "    RaiseEvent Changed(\"only-one-arg\")\n"
            "End Sub"
        ),
        ok_example=(
            "Public Event Changed(ByVal name As String, ByVal value As Variant)\n"
            "Sub Trigger()\n"
            "    RaiseEvent Changed(\"a\", 1)\n"
            "End Sub"
        ),
        fix_hint="Pass the matching number of arguments (respecting Optional / ParamArray).",
    ),
]


_RULES_BY_ID = {r.rule_id: r for r in _RULES}


def get_rule(rule_id: str) -> Rule | None:
    return _RULES_BY_ID.get(rule_id)


def all_rules() -> list[Rule]:
    """Return all registered rules sorted by id."""
    return sorted(_RULES, key=lambda r: r.rule_id)


def known_rule_ids() -> set[str]:
    return set(_RULES_BY_ID.keys())


__all__ = ["Rule", "get_rule", "all_rules", "known_rule_ids"]
