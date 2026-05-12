"""Phase 2.6 — UDT/Class member-chain depth.

Verifies that the analyser walks dotted member chains of arbitrary depth
without losing element types through array indexing, function-call
returns, or property-get returns. Each test asserts both directions:

  * a known-good leaf member resolves cleanly,
  * a typo'd leaf member produces a "Member 'X' not found in type 'Y'"
    error at the correct hop.

Regression-anchored by `tests/samples/valid_code/deep_member_chain.bas`
and the new fixtures under
`tests/samples/compile_errors/member_access/`.
"""
from __future__ import annotations


def _err_messages(result):
    return [e.get("message", "") for e in result.errors]


def test_deep_udt_chain_resolves_leaf(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Lvl4
    n As Long
End Type
Private Type Lvl3
    a As Lvl4
End Type
Private Type Lvl2
    b As Lvl3
End Type
Private Type Lvl1
    c As Lvl2
End Type

Sub S()
    Dim x As Lvl1
    Dim n As Long
    n = x.c.b.a.n
End Sub
"""
    result = run_source(code)
    member_errors = [
        e for e in result.errors
        if "not found in type" in e.get("message", "")
    ]
    assert not member_errors, (
        f"Valid 5-deep chain must not produce member-not-found. "
        f"Got: {member_errors!r}"
    )


def test_deep_udt_chain_flags_typo_at_correct_hop(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Lvl4
    n As Long
End Type
Private Type Lvl3
    a As Lvl4
End Type
Private Type Lvl2
    b As Lvl3
End Type
Private Type Lvl1
    c As Lvl2
End Type

Sub S()
    Dim x As Lvl1
    Dim n As Long
    n = x.c.b.a.bogus
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA260"
        and "bogus" in e.get("message", "")
        and "Lvl4" in e.get("message", "")
        for e in result.errors
    ), f"Expected VBA260 on 'bogus' against 'Lvl4'. Errors: {result.errors!r}"


def test_udt_array_member_preserves_element_type(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Cell
    val As Long
End Type
Private Type Row
    cells() As Cell
End Type

Sub S()
    Dim r As Row
    Dim n As Long
    n = r.cells(0).val
End Sub
"""
    result = run_source(code)
    member_errors = [
        e for e in result.errors
        if "not found in type" in e.get("message", "")
    ]
    assert not member_errors, (
        f"`cells() As Cell` must keep `Cell` as the element type. "
        f"Got: {member_errors!r}"
    )


def test_udt_array_member_chain_flags_typo(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Cell
    val As Long
End Type
Private Type Row
    cells() As Cell
End Type

Sub S()
    Dim r As Row
    Dim n As Long
    n = r.cells(0).bogus
End Sub
"""
    result = run_source(code)
    assert any(
        "bogus" in m and "Cell" in m and "not found" in m
        for m in _err_messages(result)
    ), f"Expected 'bogus' flagged against 'Cell'. Errors: {result.errors!r}"


def test_module_level_array_of_udt_preserves_type(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Cell
    val As Long
End Type

Private cells() As Cell

Sub S()
    Dim n As Long
    n = cells(0).bogus
End Sub
"""
    result = run_source(code)
    assert any(
        "bogus" in m and "Cell" in m and "not found" in m
        for m in _err_messages(result)
    ), (
        f"Module-level `Private cells() As Cell` must keep `Cell` "
        f"as the element type. Errors: {result.errors!r}"
    )


def test_function_returning_udt_chains_into_member(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Inner
    n As Long
End Type
Private Type Outer
    inner As Inner
End Type

Function GetOuter() As Outer
End Function

Sub S()
    Dim n As Long
    n = GetOuter().inner.bogus
End Sub
"""
    result = run_source(code)
    assert any(
        "bogus" in m and "Inner" in m and "not found" in m
        for m in _err_messages(result)
    ), f"Chain after function call must keep return-type. Errors: {result.errors!r}"


def test_property_get_returning_udt_chains_into_member(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Inner
    n As Long
End Type

Property Get TheInner() As Inner
End Property

Sub S()
    Dim n As Long
    n = TheInner.bogus
End Sub
"""
    result = run_source(code)
    assert any(
        "bogus" in m and "Inner" in m and "not found" in m
        for m in _err_messages(result)
    ), f"Property Get return must chain. Errors: {result.errors!r}"


def test_deeply_nested_array_of_udt_chain(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Cell
    val As Long
End Type
Private Type RowT
    cells() As Cell
End Type
Private Type Table
    rows() As RowT
End Type
Private Type Book
    tables() As Table
End Type

Sub S()
    Dim b As Book
    Dim n As Long
    n = b.tables(0).rows(0).cells(0).bogus
End Sub
"""
    result = run_source(code)
    assert any(
        "bogus" in m and "Cell" in m and "not found" in m
        for m in _err_messages(result)
    ), f"Depth-5 array-of-UDT chain must flag leaf typo. Errors: {result.errors!r}"


def test_dotted_set_let_validation_on_chain(run_source):
    """P2.6 enables Set/Let validation on dotted LHS for resolved chains.
    A UDT member is not an Object — `Set x.inner = ...` must fire VBA210.
    """
    code = """
Attribute VB_Name = "M"
Option Explicit

Private Type Inner
    n As Long
End Type
Private Type Outer
    inner As Inner
End Type

Sub S()
    Dim x As Outer
    Set x.inner = Nothing
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA210" and "x.inner" in e.get("message", "")
        for e in result.errors
    ), f"Expected VBA210 on `Set x.inner = ...`. Errors: {result.errors!r}"


def test_dotted_set_let_skipped_on_unresolvable_chain(run_source):
    """Members of permissive `Object` LHS must NOT trigger VBA210/211 —
    we can't infer the real member type from declaration alone.
    """
    code = """
Attribute VB_Name = "M"
Option Explicit

Sub S()
    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    col.SomeMember = 1
End Sub
"""
    result = run_source(code)
    assert all(
        e.get("rule_id") not in ("VBA210", "VBA211") for e in result.errors
    ), f"Dotted LHS on Object must not fire Set/Let. Errors: {result.errors!r}"
