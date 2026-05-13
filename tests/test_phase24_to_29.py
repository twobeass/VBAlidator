"""Direct tests for Phase 2.4 (operators), 2.5 (const), 2.7 (date),
2.8 (fixed-length string), 2.9 (DefType)."""
from __future__ import annotations


# ---- Phase 2.7 -----------------------------------------------------------


def test_invalid_date_literal_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim d As Date
    d = #2025-13-45#
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA_LEX002" for e in result.errors), (
        f"Invalid date literal must produce VBA_LEX002. Errors: {result.errors!r}"
    )


def test_valid_date_formats_accepted(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim d As Date
    d = #1/1/2020#
    d = #2020-01-01#
    d = #1/1/100#
    d = #January 1, 2020#
    d = #1-Jan-2020#
    d = #12:00:00 AM#
End Sub
"""
    result = run_source(code)
    bad = [e for e in result.errors if e.get("rule_id") == "VBA_LEX002"]
    assert not bad, f"Valid date formats must not flag. Got: {bad!r}"


# ---- Phase 2.5 -----------------------------------------------------------


def test_const_calling_function_is_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Const X As Long = MsgBox("x")
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA230" for e in result.errors), (
        f"Const RHS calling a function must produce VBA230. Errors: {result.errors!r}"
    )


def test_const_referencing_variable_is_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim runtimeValue As Long
    runtimeValue = 1
    Const X As Long = runtimeValue
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA231" for e in result.errors), (
        f"Const RHS referencing a variable must produce VBA231. Errors: {result.errors!r}"
    )


def test_const_referencing_other_const_is_ok(run_source):
    """Const RHS may reference another constant or built-in vb* alias."""
    code = """
Attribute VB_Name = "M"
Sub S()
    Const A As Long = 5
    Const B As Long = A + 1
    Const C As Long = vbRed
End Sub
"""
    result = run_source(code)
    bad = [e for e in result.errors if e.get("rule_id", "").startswith("VBA23")]
    assert not bad, f"Const RHS referencing another Const must not flag. Got: {bad!r}"


# ---- Phase 2.4 -----------------------------------------------------------


def test_arith_string_minus_int_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    x = "abc" - 1
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA240" for e in result.errors), (
        f"`\"abc\" - 1` must produce VBA240. Errors: {result.errors!r}"
    )


def test_arith_string_times_int_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Long
    x = 3 * "y"
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA240" for e in result.errors)


def test_string_concat_with_ampersand_not_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim s As String
    s = "value: " & 1
End Sub
"""
    result = run_source(code)
    bad = [e for e in result.errors if e.get("rule_id") == "VBA240"]
    assert not bad, f"`&` is concat, must not flag. Got: {bad!r}"


def test_plus_between_string_and_int_not_flagged(run_source):
    """+ is bidirectionally coerced in VBA — too lenient to flag."""
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim x As Variant
    x = "1" + 1
End Sub
"""
    result = run_source(code)
    bad = [e for e in result.errors if e.get("rule_id") == "VBA240"]
    assert not bad, f"`+` must not be flagged for string+int. Got: {bad!r}"


# ---- Phase 2.8 -----------------------------------------------------------


def test_fixed_length_string_in_procedure_flagged(run_source):
    code = """
Attribute VB_Name = "M"
Sub S()
    Dim s As String * 10
End Sub
"""
    result = run_source(code)
    assert any(e.get("rule_id") == "VBA250" for e in result.errors), (
        f"`Dim s As String * N` inside a Sub must produce VBA250. Errors: {result.errors!r}"
    )


# ---- Phase 2.9 -----------------------------------------------------------


def test_def_type_applied_to_implicit_typing(run_source):
    """`DefInt I-N` makes Dim i (no `As`) imply Integer.

    The check is structural: we feed a typed assignment that would only be
    valid if the implicit type matched. The analyzer must not report a
    type-related error.
    """
    code = """
Attribute VB_Name = "M"
DefInt I-N
Sub S()
    Dim i, j
    i = 1
    j = i + 1
End Sub
"""
    run_source(code)  # smoke-test the e2e pipeline path
    # Implementation detail: we don't yet emit type-mismatch errors on
    # arbitrary assignments (Phase 2.4 only flags literal-literal). What
    # we *can* assert is that the variable was registered with type
    # Integer, not Variant. Run a second pass and inspect.
    from src.config import Config
    from src.lexer import Lexer
    from src.parser import VBAParser
    from src.analyzer import Analyzer

    tokens = list(Lexer(code).tokenize())
    parser = VBAParser(tokens, filename="M.bas")
    module = parser.parse_module()
    module.module_type = "Module"
    module.filename = "M.bas"

    assert module.def_type_map.get("i") == "Integer", (
        f"DefInt I-N must populate def_type_map. Got: {module.def_type_map!r}"
    )
    assert module.def_type_map.get("n") == "Integer"
    assert "z" not in module.def_type_map  # outside the range

    # Also: passes analysis without type-related errors on the implicit vars.
    analyzer = Analyzer(Config())
    analyzer.add_module(module)
    errors = analyzer.analyze()
    type_errors = [e for e in errors if "i" in e.get("message", "").lower() or "j" in e.get("message", "").lower()]
    assert not any("undefined" in e.get("message", "").lower() for e in type_errors)


def test_def_str_applies_to_string():
    from src.lexer import Lexer
    from src.parser import VBAParser

    code = "DefStr S\nDim sName\n"
    tokens = list(Lexer(code).tokenize())
    module = VBAParser(tokens).parse_module()
    assert module.def_type_map.get("s") == "String"
