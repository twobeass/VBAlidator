"""Iter-4 Wave 4 community-code FP fixes.

Covers:
  * #18 — `On X GoTo lbl1, lbl2, ...` jump-table syntax
  * #19 — `CallByName` + `VbCallType` enum in the VBA runtime model
  * #20 — `ReDim obj.member(...)` with member-access targets
"""
from __future__ import annotations


# ---- #18 — On X GoTo / GoSub jump table ---------------------------------


def test_on_goto_jump_table_with_valid_labels_is_silent(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub JumpTable()
    Dim op As Long
    op = 1
    On op GoTo lblA, lblB, lblC
    Exit Sub
lblA:
    Debug.Print "A"
    Exit Sub
lblB:
    Debug.Print "B"
    Exit Sub
lblC:
    Debug.Print "C"
End Sub
"""
    result = run_source(code)
    hard = [e for e in result.errors if e.get("severity", "error") == "error"]
    assert not hard, f"Valid jump-table must be clean. Got: {hard!r}"


def test_on_goto_jump_table_flags_missing_label(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub Bad()
    Dim op As Long
    op = 1
    On op GoTo realLabel, missingLabel
    Exit Sub
realLabel:
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA201" and "missingLabel" in e.get("message", "")
        for e in result.errors
    ), f"Missing label in jump-table must fire VBA201. Got: {result.errors!r}"


def test_on_goto_jump_table_does_not_emit_vba001_per_label(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub J()
    Dim op As Long: op = 1
    On op GoTo a, b
    Exit Sub
a:
    Exit Sub
b:
End Sub
"""
    result = run_source(code)
    vba001 = [
        e for e in result.errors
        if e.get("rule_id") == "VBA001"
        and any(lbl in e.get("message", "") for lbl in ("'a'", "'b'"))
    ]
    assert not vba001, (
        f"Labels in On…GoTo list must not be flagged as undefined "
        f"identifiers. Got: {vba001!r}"
    )


def test_on_gosub_jump_table_also_recognised(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub J()
    Dim op As Long: op = 1
    On op GoSub a, b
    Exit Sub
a:
    Return
b:
    Return
End Sub
"""
    result = run_source(code)
    bad = [
        e for e in result.errors
        if e.get("rule_id") == "VBA001"
        and any(lbl in e.get("message", "") for lbl in ("'a'", "'b'"))
    ]
    assert not bad, f"On…GoSub labels also must not fire VBA001. Got: {bad!r}"


# ---- #19 — CallByName + VbCallType --------------------------------------


def test_callbyname_is_known_builtin(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S(obj As Object)
    Dim r As Variant
    r = CallByName(obj, "DoIt", VbMethod, 42)
End Sub
"""
    result = run_source(code)
    bad = [
        e for e in result.errors
        if e.get("rule_id") == "VBA001"
        and any(name in e.get("message", "") for name in ("CallByName", "VbMethod"))
    ]
    assert not bad, (
        f"`CallByName` and `VbMethod` must be in the VBA runtime model. "
        f"Got: {bad!r}"
    )


def test_vbcalltype_enum_members_resolve(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim k As Long
    k = vbMethod
    k = vbGet
    k = vbLet
    k = vbSet
End Sub
"""
    result = run_source(code)
    bad = [
        e for e in result.errors
        if e.get("rule_id") == "VBA001"
        and any(n in e.get("message", "") for n in ("vbMethod", "vbGet", "vbLet", "vbSet"))
    ]
    assert not bad, f"VbCallType enum members must resolve. Got: {bad!r}"


def test_qualified_vbcalltype_resolves(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub S()
    Dim k As Long
    k = VbCallType.vbMethod
    k = VbCallType.vbGet Or VbCallType.vbMethod
End Sub
"""
    result = run_source(code)
    hard = [e for e in result.errors if e.get("severity", "error") == "error"]
    assert not hard, f"Qualified VbCallType access must resolve. Got: {hard!r}"


# ---- #20 — ReDim with member-access target ------------------------------


def test_redim_on_udt_array_member_is_silent(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Type Container
    items() As Variant
End Type
Private c As Container
Sub Resize(n As Long)
    ReDim c.items(1 To n)
End Sub
"""
    result = run_source(code)
    bad = [
        e for e in result.errors
        if e.get("rule_id") in ("VBA101", "VBA102", "VBA103")
    ]
    assert not bad, (
        f"`ReDim c.items(...)` on a UDT array member must not fire. "
        f"Got: {bad!r}"
    )


def test_redim_preserve_on_udt_array_member_is_silent(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Type Container
    items() As Variant
End Type
Private c As Container
Sub Resize(n As Long)
    ReDim Preserve c.items(0 To n - 1)
End Sub
"""
    result = run_source(code)
    bad = [
        e for e in result.errors
        if e.get("rule_id") in ("VBA101", "VBA102", "VBA103")
    ]
    assert not bad, f"`ReDim Preserve c.items(...)` must not fire. Got: {bad!r}"


def test_redim_on_udt_scalar_member_fires_vba103(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Private Type Container
    one As Long
End Type
Private c As Container
Sub Bad()
    ReDim c.one(1 To 5)
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA103" and "c.one" in e.get("message", "")
        for e in result.errors
    ), (
        f"ReDim on scalar UDT member must fire VBA103. Got: {result.errors!r}"
    )


def test_redim_on_truly_undefined_target_still_fires_vba101(run_source):
    code = """
Attribute VB_Name = "M"
Option Explicit
Sub Bad()
    ReDim NoSuch(1 To 5)
End Sub
"""
    result = run_source(code)
    assert any(
        e.get("rule_id") == "VBA101" and "NoSuch" in e.get("message", "")
        for e in result.errors
    ), f"Undefined ReDim target must still fire VBA101. Got: {result.errors!r}"
