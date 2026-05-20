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
    Rule(
        rule_id="VBA_LEX004",
        title="Unterminated string or missing end bracket",
        severity="error",
        category="lexer",
        phase="3.9",
        description=(
            "A string literal has no closing `\"`, or a bracket-quoted name "
            "`[…]` has no closing `]`."
        ),
        fail_example='s = "hello',
        ok_example='s = "hello"',
        fix_hint="Add the missing closing quote or `]`.",
    ),
    Rule(
        rule_id="VBA_LEX005",
        title="Identifier too long",
        severity="error",
        category="lexer",
        phase="3.9",
        description="VBA identifiers (names) are limited to 255 characters.",
        fail_example="Dim aaaa…(>255 chars) As Long",
        ok_example="Dim total As Long",
        fix_hint="Shorten the identifier to 255 characters or fewer.",
    ),
    Rule(
        rule_id="VBA_LEX006",
        title="Line too long",
        severity="error",
        category="lexer",
        phase="3.9",
        description="A physical source line exceeds VBA's 1023-character limit.",
        fail_example="' a single line longer than 1023 characters …",
        ok_example="' wrap with the line-continuation character `_`",
        fix_hint="Break the line using the `_` line-continuation character.",
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

    # -- Phase 1.6: declared-type resolution ------------------------
    Rule(
        rule_id="VBA120",
        title="User-defined type not defined",
        severity="warning",
        category="type",
        phase="1.6",
        description=(
            "A declared type name does not resolve to any built-in type, "
            "source-declared `Type`/`Enum`/class, or host-model class. "
            "Library-qualified names (`Scripting.Dictionary`) are accepted "
            "leniently so un-modelled hosts don't false-positive."
        ),
        fail_example="Dim x As Custmer   ' typo for Customer",
        ok_example="Dim x As Customer",
        fix_hint="Define the `Type`/`Enum`/class, fix the spelling, or load the matching host model with `--host`.",
    ),
    Rule(
        rule_id="VBA122",
        title="User-defined type without members",
        severity="error",
        category="type",
        phase="1.6",
        description="A `Type … End Type` block must declare at least one member.",
        fail_example="Type Empty\nEnd Type",
        ok_example="Type Point\n    x As Long\nEnd Type",
        fix_hint="Add at least one field, or remove the empty Type.",
    ),
    Rule(
        rule_id="VBA123",
        title="Empty Enum not allowed",
        severity="error",
        category="type",
        phase="1.6",
        description="An `Enum … End Enum` block must declare at least one member.",
        fail_example="Enum Empty\nEnd Enum",
        ok_example="Enum Color\n    Red\nEnd Enum",
        fix_hint="Add at least one member, or remove the empty Enum.",
    ),

    # -- Phase 1.5: block-terminator pairing ------------------------
    Rule(
        rule_id="VBA110",
        title="For without Next",
        severity="error",
        category="syntax",
        phase="1.5",
        description=(
            "A `For` / `For Each` loop body reached the end of the procedure "
            "(or module) without a matching `Next`. VBE refuses to compile an "
            "unterminated loop."
        ),
        fail_example="For i = 1 To 10\n    Debug.Print i",
        ok_example="For i = 1 To 10\n    Debug.Print i\nNext i",
        fix_hint="Add the matching `Next` (optionally `Next i`) that closes the loop.",
    ),
    Rule(
        rule_id="VBA111",
        title="Do without Loop",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `Do` block has no matching `Loop`.",
        fail_example="Do While x < 10\n    x = x + 1",
        ok_example="Do While x < 10\n    x = x + 1\nLoop",
        fix_hint="Add the `Loop` (optionally `Loop While`/`Loop Until`) that closes the `Do`.",
    ),
    Rule(
        rule_id="VBA112",
        title="While without Wend",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `While` block has no matching `Wend`.",
        fail_example="While x < 10\n    x = x + 1",
        ok_example="While x < 10\n    x = x + 1\nWend",
        fix_hint="Add the `Wend` that closes the `While`.",
    ),
    Rule(
        rule_id="VBA113",
        title="Expected End With",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `With` block has no matching `End With`.",
        fail_example="With rng\n    .Value = 1",
        ok_example="With rng\n    .Value = 1\nEnd With",
        fix_hint="Add the `End With` that closes the block.",
    ),
    Rule(
        rule_id="VBA114",
        title="Select Case without End Select",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `Select Case` block has no matching `End Select`.",
        fail_example="Select Case x\n    Case 1\n        Debug.Print 1",
        ok_example="Select Case x\n    Case 1\n        Debug.Print 1\nEnd Select",
        fix_hint="Add the `End Select` that closes the block.",
    ),
    Rule(
        rule_id="VBA115",
        title="Block If without End If",
        severity="error",
        category="syntax",
        phase="1.5",
        description=(
            "A block `If … Then` (the multi-line form, where `Then` is the "
            "last token on the line) reached the end of the procedure without "
            "a matching `End If`."
        ),
        fail_example="If x > 0 Then\n    Debug.Print x",
        ok_example="If x > 0 Then\n    Debug.Print x\nEnd If",
        fix_hint="Add the `End If`, or convert to a single-line `If x > 0 Then Debug.Print x`.",
    ),
    Rule(
        rule_id="VBA116",
        title="Next without For",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `Next` appears with no `For` loop open in scope.",
        fail_example="    Debug.Print i\nNext i",
        ok_example="For i = 1 To 10\n    Debug.Print i\nNext i",
        fix_hint="Remove the stray `Next` or add the opening `For`.",
    ),
    Rule(
        rule_id="VBA117",
        title="Loop without Do",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `Loop` appears with no `Do` block open in scope.",
        fail_example="    x = x + 1\nLoop",
        ok_example="Do\n    x = x + 1\nLoop While x < 10",
        fix_hint="Remove the stray `Loop` or add the opening `Do`.",
    ),
    Rule(
        rule_id="VBA118",
        title="Wend without While",
        severity="error",
        category="syntax",
        phase="1.5",
        description="A `Wend` appears with no `While` block open in scope.",
        fail_example="    x = x + 1\nWend",
        ok_example="While x < 10\n    x = x + 1\nWend",
        fix_hint="Remove the stray `Wend` or add the opening `While`.",
    ),
    Rule(
        rule_id="VBA119",
        title="Block terminator without matching opener",
        severity="error",
        category="syntax",
        phase="1.5",
        description=(
            "A block terminator (`Else`, `ElseIf`, `End If`, `End Select`, "
            "`End With`, …) appears with no matching opening statement in "
            "scope — typically the symptom of an earlier mis-nested block."
        ),
        fail_example="Debug.Print 1\nEnd If",
        ok_example="If x > 0 Then\n    Debug.Print 1\nEnd If",
        fix_hint="Fix the surrounding block nesting; an opener is missing or an earlier block closed too early.",
    ),

    # -- Phase 2.10: parameter-list well-formedness -----------------
    Rule(
        rule_id="VBA130",
        title="Required parameter after Optional",
        severity="error",
        category="signature",
        phase="2.10",
        description=(
            "Once a parameter is declared `Optional`, every parameter after "
            "it must also be `Optional` (the trailing `ParamArray` aside)."
        ),
        fail_example="Sub S(Optional a As Long, b As Long): End Sub",
        ok_example="Sub S(Optional a As Long, Optional b As Long): End Sub",
        fix_hint="Mark the trailing parameters `Optional`, or move the required ones before the first Optional.",
    ),
    Rule(
        rule_id="VBA131",
        title="ParamArray is not the last parameter",
        severity="error",
        category="signature",
        phase="2.10",
        description="`ParamArray` must be the final parameter of a procedure.",
        fail_example="Sub S(ParamArray a() As Variant, b As Long): End Sub",
        ok_example="Sub S(b As Long, ParamArray a() As Variant): End Sub",
        fix_hint="Move the `ParamArray` to the end of the parameter list.",
    ),
    Rule(
        rule_id="VBA132",
        title="More than one ParamArray",
        severity="error",
        category="signature",
        phase="2.10",
        description="A procedure may declare at most one `ParamArray`.",
        fail_example="Sub S(ParamArray a() As Variant, ParamArray b() As Variant): End Sub",
        ok_example="Sub S(ParamArray a() As Variant): End Sub",
        fix_hint="Keep a single `ParamArray` and pass the rest explicitly.",
    ),
    Rule(
        rule_id="VBA133",
        title="ParamArray combined with Optional",
        severity="error",
        category="signature",
        phase="2.10",
        description="A procedure cannot mix `Optional` parameters with a `ParamArray`.",
        fail_example="Sub S(Optional a As Long, ParamArray b() As Variant): End Sub",
        ok_example="Sub S(a As Long, ParamArray b() As Variant): End Sub",
        fix_hint="Drop the `Optional` markers or the `ParamArray` — they are mutually exclusive.",
    ),
    Rule(
        rule_id="VBA134",
        title="ParamArray element type is not Variant",
        severity="error",
        category="signature",
        phase="2.10",
        description="A `ParamArray` must be declared as an array of `Variant`.",
        fail_example="Sub S(ParamArray a() As Long): End Sub",
        ok_example="Sub S(ParamArray a() As Variant): End Sub",
        fix_hint="Declare the ParamArray as `ParamArray name() As Variant`.",
    ),
    Rule(
        rule_id="VBA137",
        title="User-defined type passed ByVal",
        severity="error",
        category="signature",
        phase="2.10",
        description="A user-defined `Type` parameter must be passed `ByRef`, not `ByVal`.",
        fail_example="Type T\n    x As Long\nEnd Type\nSub S(ByVal p As T)\nEnd Sub",
        ok_example="Type T\n    x As Long\nEnd Type\nSub S(ByRef p As T)\nEnd Sub",
        fix_hint="Change the parameter to `ByRef` (the default).",
    ),
    Rule(
        rule_id="VBA136",
        title="Array parameter passed ByVal",
        severity="error",
        category="signature",
        phase="2.10",
        description="Array parameters must be passed `ByRef`; `ByVal` arrays are illegal in VBA.",
        fail_example="Sub S(ByVal a() As Long): End Sub",
        ok_example="Sub S(ByRef a() As Long): End Sub",
        fix_hint="Change the parameter mechanism to `ByRef` (the default).",
    ),

    # -- Phase 2.14: type-declaration character --------------------
    Rule(
        rule_id="VBA270",
        title="Type-declaration character does not match declared type",
        severity="error",
        category="declaration",
        phase="2.14",
        description=(
            "A legacy type-declaration suffix (`$` String, `%` Integer, `@` "
            "Currency) on a variable name must agree with its explicit `As` "
            "type."
        ),
        fail_example="Dim count% As Long",
        ok_example="Dim count As Long   ' or: Dim count%",
        fix_hint="Drop the suffix, or make the `As` type match it.",
    ),

    # -- Phase 3.10: object-module member restrictions -------------
    Rule(
        rule_id="VBA370",
        title="Constant/array/Declare as Public member of object module",
        severity="error",
        category="placement",
        phase="3.10",
        description=(
            "Constants, arrays and `Declare` statements cannot be Public "
            "members of an object module (Class / Form). They are only legal "
            "in standard modules, or as Private members."
        ),
        fail_example="' in a .cls:\nPublic Const MAX As Long = 10",
        ok_example="' in a .cls:\nPrivate Const MAX As Long = 10",
        fix_hint="Move it to a standard module, or mark it Private and expose it via a Property Get.",
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

    # -- Phase 2.9: named arguments at call sites -------------------
    Rule(
        rule_id="VBA140",
        title="Named argument not found",
        severity="error",
        category="signature",
        phase="2.9",
        description="A `name:=value` call argument names a parameter the target procedure does not declare.",
        fail_example="Sub S(ByVal a As Long)\nEnd Sub\nSub T()\n    Call S(b:=1)\nEnd Sub",
        ok_example="Sub S(ByVal a As Long)\nEnd Sub\nSub T()\n    Call S(a:=1)\nEnd Sub",
        fix_hint="Use a parameter name that exists in the procedure signature.",
    ),
    Rule(
        rule_id="VBA141",
        title="Named argument specified more than once",
        severity="error",
        category="signature",
        phase="2.9",
        description="The same named argument appears twice in one call.",
        fail_example="Call S(a:=1, a:=2)",
        ok_example="Call S(a:=1, b:=2)",
        fix_hint="Pass each named argument at most once.",
    ),
    Rule(
        rule_id="VBA142",
        title="Named arguments with ParamArray",
        severity="error",
        category="signature",
        phase="2.9",
        description="A procedure declaring a `ParamArray` cannot be called using named arguments.",
        fail_example="Sub S(ParamArray a() As Variant)\nEnd Sub\nSub T()\n    Call S(a:=1)\nEnd Sub",
        ok_example="Sub S(ParamArray a() As Variant)\nEnd Sub\nSub T()\n    Call S(1, 2, 3)\nEnd Sub",
        fix_hint="Pass arguments positionally when the procedure has a ParamArray.",
    ),

    # -- Phase 2.12: procedure / property consistency ---------------
    Rule(
        rule_id="VBA151",
        title="Code after End Sub/Function/Property",
        severity="error",
        category="syntax",
        phase="2.12",
        description=(
            "Only comments may appear after `End Sub`, `End Function`, or "
            "`End Property` on the same line."
        ),
        fail_example="Sub S()\nEnd Sub: x = 1",
        ok_example="Sub S()\nEnd Sub   ' done",
        fix_hint="Move the trailing statement to its own line.",
    ),
    Rule(
        rule_id="VBA150",
        title="Ambiguous name detected (duplicate procedure)",
        severity="error",
        category="declaration",
        phase="2.12",
        description=(
            "Two procedures share the same name in one module. Property "
            "Get/Let/Set may share a name, but two of the same accessor kind "
            "(or two Subs/Functions) may not."
        ),
        fail_example="Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub",
        ok_example="Sub Foo()\nEnd Sub\nSub Bar()\nEnd Sub",
        fix_hint="Rename or remove the duplicate procedure.",
    ),
    Rule(
        rule_id="VBA152",
        title="Property Get/Let type inconsistent",
        severity="error",
        category="property",
        phase="2.12",
        description=(
            "A `Property Get` return type and the matching `Property Let` "
            "value-parameter type must agree."
        ),
        fail_example=(
            "Property Get Foo() As Long\nEnd Property\n"
            "Property Let Foo(ByVal v As String)\nEnd Property"
        ),
        ok_example=(
            "Property Get Foo() As Long\nEnd Property\n"
            "Property Let Foo(ByVal v As Long)\nEnd Property"
        ),
        fix_hint="Make the Get return type and the Let value-parameter type identical.",
    ),
    Rule(
        rule_id="VBA153",
        title="Implements method signature mismatch",
        severity="error",
        category="interface",
        phase="2.12",
        description=(
            "A method implementing an interface member must match the "
            "member's parameter count."
        ),
        fail_example=(
            "' IShape: Public Sub Draw(ByVal scale As Double)\n"
            "Implements IShape\n"
            "Private Sub IShape_Draw()\nEnd Sub"
        ),
        ok_example=(
            "Implements IShape\n"
            "Private Sub IShape_Draw(ByVal scale As Double)\nEnd Sub"
        ),
        fix_hint="Match the interface member's parameter list exactly.",
    ),

    # -- Phase 2.13: assignment / expression semantics --------------
    Rule(
        rule_id="VBA160",
        title="Assignment to constant not permitted",
        severity="error",
        category="assignment",
        phase="2.13",
        description="A `Const` or `Enum` member is read-only and cannot be assigned to.",
        fail_example="Const PI As Double = 3.14159\nPI = 3",
        ok_example="Const PI As Double = 3.14159\nDim r As Double\nr = PI",
        fix_hint="Assign to a variable, not a constant or Enum member.",
    ),
    Rule(
        rule_id="VBA164",
        title="Invalid use of Me keyword",
        severity="error",
        category="assignment",
        phase="2.13",
        description="`Me` refers to the current instance and cannot be assigned to.",
        fail_example="Set Me = Nothing",
        ok_example="Set obj = Me",
        fix_hint="Use a regular object variable as the assignment target.",
    ),
    Rule(
        rule_id="VBA165",
        title="Invalid use of New keyword",
        severity="error",
        category="declaration",
        phase="2.13",
        description="`New` requires a creatable class; primitives, `Object` and `Variant` cannot be `New`-ed.",
        fail_example="Dim x As New Long",
        ok_example="Dim c As New Collection",
        fix_hint="Use `New` only with a class type, or drop `New` for primitives.",
    ),

    # -- Phase 2.11: loop / label context ---------------------------
    Rule(
        rule_id="VBA170",
        title="Exit For not within For...Next",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="`Exit For` is only valid inside a `For` / `For Each` loop.",
        fail_example="Sub S()\n    Exit For\nEnd Sub",
        ok_example="Sub S()\n    For i = 1 To 10\n        Exit For\n    Next i\nEnd Sub",
        fix_hint="Remove the `Exit For` or place it inside a `For` loop.",
    ),
    Rule(
        rule_id="VBA171",
        title="Exit Do not within Do...Loop",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="`Exit Do` is only valid inside a `Do...Loop`.",
        fail_example="Sub S()\n    Exit Do\nEnd Sub",
        ok_example="Sub S()\n    Do\n        Exit Do\n    Loop\nEnd Sub",
        fix_hint="Remove the `Exit Do` or place it inside a `Do` loop.",
    ),
    Rule(
        rule_id="VBA172",
        title="Next control variable mismatch",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="The variable on `Next x` must match the loop's `For` counter.",
        fail_example="For i = 1 To 10\n    Debug.Print i\nNext j",
        ok_example="For i = 1 To 10\n    Debug.Print i\nNext i",
        fix_hint="Use the same variable as the `For` counter, or a bare `Next`.",
    ),
    Rule(
        rule_id="VBA173",
        title="For control variable already in use",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="A nested `For` loop reuses the counter variable of an enclosing loop.",
        fail_example="For i = 1 To 3\n    For i = 1 To 3\n    Next i\nNext i",
        ok_example="For i = 1 To 3\n    For j = 1 To 3\n    Next j\nNext i",
        fix_hint="Use a distinct counter variable for the inner loop.",
    ),
    Rule(
        rule_id="VBA174",
        title="For Each control variable must be Variant or object",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="A `For Each` control variable cannot be a primitive scalar type.",
        fail_example="Dim i As Long\nFor Each i In coll\nNext i",
        ok_example="Dim v As Variant\nFor Each v In coll\nNext v",
        fix_hint="Declare the control variable as `Variant`, `Object`, or a class type.",
    ),
    Rule(
        rule_id="VBA175",
        title="Duplicate label",
        severity="error",
        category="control_flow",
        phase="2.11",
        description="The same line label is declared more than once in a procedure.",
        fail_example="Sub S()\nDone:\n    Exit Sub\nDone:\nEnd Sub",
        ok_example="Sub S()\nDone:\n    Exit Sub\nEnd Sub",
        fix_hint="Rename or remove the duplicate label.",
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

    # NOTE on Phase 2.6 — Deep member-chain typing landed but does NOT
    # introduce a new rule ID. The existing VBA002 ("Member not found")
    # diagnostic was already the right umbrella; P2.6 just made the
    # chain walker accurate enough that VBA002 fires at the correct hop
    # for arbitrarily deep chains (array index, function return, etc.).
    # See `tests/test_phase26_member_chain.py` for coverage.

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

    # -- Phase 3 (continued) — terminator mismatch -------------------
    Rule(
        rule_id="VBA350",
        title="Procedure terminator does not match its kind",
        severity="error",
        category="syntax",
        phase="3.5",
        description=(
            "`End Sub` closes a Sub, `End Function` closes a Function, "
            "`End Property` closes a Property. Mixing them is a hard VBE "
            "compile error and a common AI-generator slip after editing a "
            "procedure's signature."
        ),
        fail_example=(
            "Function F() As Long\n"
            "    F = 1\n"
            "End Sub\n"
        ),
        ok_example=(
            "Function F() As Long\n"
            "    F = 1\n"
            "End Function\n"
        ),
        fix_hint="Use the End form that matches the procedure declaration.",
    ),

    # -- Phase 3.5 — Statement placement -----------------------------
    Rule(
        rule_id="VBA360",
        title="Module-only declaration inside a procedure body",
        severity="error",
        category="placement",
        phase="3.5",
        description=(
            "`Type`, `Enum`, `Declare`, `Option`, `Implements` and the "
            "`DefInt`/`DefStr`/… family are only legal at module scope. "
            "VBE refuses to compile a Sub/Function/Property that contains "
            "one of these declarations."
        ),
        fail_example=(
            "Sub S()\n"
            "    Type Point\n"
            "        x As Long\n"
            "        y As Long\n"
            "    End Type\n"
            "End Sub\n"
        ),
        ok_example=(
            "Type Point\n"
            "    x As Long\n"
            "    y As Long\n"
            "End Type\n"
            "\n"
            "Sub S()\n"
            "    Dim p As Point\n"
            "End Sub\n"
        ),
        fix_hint="Move the declaration outside any procedure, between the module attributes and the first Sub/Function.",
    ),
    Rule(
        rule_id="VBA361",
        title="Executable statement at module level",
        severity="error",
        category="placement",
        phase="3.5",
        description=(
            "Code outside a Sub/Function/Property runs implicitly at "
            "module load — but VBA only permits *declarative* statements "
            "(`Public`/`Private`/`Dim`/`Const`/`Type`/`Enum`/`Declare`/"
            "`Option`/`Implements`/`Attribute`/`DefXxx`) in that position. "
            "A stray `Debug.Print` or assignment at module top is a hard "
            "compile error."
        ),
        fail_example=(
            'Attribute VB_Name = "M"\n'
            "Option Explicit\n"
            "Debug.Print \"this is illegal at module level\"\n"
            "Sub S(): End Sub\n"
        ),
        ok_example=(
            'Attribute VB_Name = "M"\n'
            "Option Explicit\n"
            "Sub S()\n"
            "    Debug.Print \"executed when S runs\"\n"
            "End Sub\n"
        ),
        fix_hint="Move the statement into a procedure body — `Sub Main()` is a common entry point, or wrap it in an `Auto_Open` / `Workbook_Open` event handler.",
    ),

    # -- Phase 3.7 — Type / Enum / Def / Option placement -----------
    Rule(
        rule_id="VBA180",
        title="Statement invalid inside Enum block",
        severity="error",
        category="placement",
        phase="3.7",
        description=(
            "An `Enum … End Enum` block may only contain member declarations "
            "(`Name [= value]`). Executable statements or `Dim`/`Sub`/… "
            "declarations inside it are illegal."
        ),
        fail_example="Enum E\n    A = 1\n    Dim x As Long\nEnd Enum",
        ok_example="Enum E\n    A = 1\n    B = 2\nEnd Enum",
        fix_hint="Move the statement outside the Enum; keep only `Name = value` members inside.",
    ),
    Rule(
        rule_id="VBA181",
        title="Statement invalid inside Type block",
        severity="error",
        category="placement",
        phase="3.7",
        description=(
            "A `Type … End Type` block may only contain field declarations "
            "(`Name [(dims)] As Type`). Executable statements or "
            "`Dim`/`Sub`/… declarations inside it are illegal."
        ),
        fail_example="Type T\n    x As Long\n    Sub Foo()\nEnd Type",
        ok_example="Type T\n    x As Long\n    y As String\nEnd Type",
        fix_hint="Move the statement outside the Type; keep only `Name As Type` fields inside.",
    ),
    Rule(
        rule_id="VBA182",
        title="Deftype statement must precede declarations",
        severity="error",
        category="placement",
        phase="3.7",
        description=(
            "`DefInt`/`DefStr`/… statements must appear before any variable, "
            "constant, type or procedure declaration in the module."
        ),
        fail_example="Dim x As Long\nDefInt A-Z",
        ok_example="DefInt A-Z\nDim x As Long",
        fix_hint="Move the `DefXxx` statement to the top of the module, after `Option` but before any declaration.",
    ),
    Rule(
        rule_id="VBA183",
        title="Duplicate Deftype letter range",
        severity="error",
        category="placement",
        phase="3.7",
        description="A letter cannot be covered by more than one `DefXxx` statement.",
        fail_example="DefInt A-K\nDefStr H-Z",
        ok_example="DefInt A-G\nDefStr H-Z",
        fix_hint="Make the `DefXxx` letter ranges disjoint.",
    ),
    Rule(
        rule_id="VBA184",
        title="Duplicate Option statement",
        severity="error",
        category="placement",
        phase="3.7",
        description="The same `Option` statement (Explicit/Compare/Base/Private Module) appears twice.",
        fail_example="Option Explicit\nOption Explicit",
        ok_example="Option Explicit",
        fix_hint="Remove the duplicate `Option` statement.",
    ),

    # -- Phase 3.8 — conditional compilation (#If) ------------------
    Rule(
        rule_id="VBA190",
        title="Conditional directive without matching #If",
        severity="error",
        category="preprocessor",
        phase="3.8",
        description=(
            "`#Else`, `#ElseIf` and `#End If` must be preceded by a matching "
            "`#If`. A directive with no open `#If` is a compile error."
        ),
        fail_example="#Else\n    x = 1\n#End If",
        ok_example="#If Win64 Then\n    x = 1\n#Else\n    x = 2\n#End If",
        fix_hint="Add the opening `#If … Then`, or remove the orphan directive.",
    ),
    Rule(
        rule_id="VBA191",
        title="Missing #End If",
        severity="error",
        category="preprocessor",
        phase="3.8",
        description="Every `#If` conditional-compilation block must be closed with `#End If`.",
        fail_example="#If Win64 Then\n    x = 1",
        ok_example="#If Win64 Then\n    x = 1\n#End If",
        fix_hint="Add the `#End If` that closes the `#If` block.",
    ),
    Rule(
        rule_id="VBA192",
        title="Invalid #Const declaration",
        severity="error",
        category="preprocessor",
        phase="3.8",
        description="A `#Const` directive must have the form `#Const Name = expression`.",
        fail_example="#Const",
        ok_example="#Const DEBUG_BUILD = 1",
        fix_hint="Write `#Const Name = <constant expression>`.",
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
