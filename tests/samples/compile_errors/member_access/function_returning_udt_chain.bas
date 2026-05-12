Attribute VB_Name = "FuncRetUdt"
Option Explicit

Private Type Inner
    n As Long
End Type

Private Type Outer
    inner As Inner
End Type

Function GetOuter() As Outer
End Function

Sub ChainAfterCall()
    Dim n As Long
    n = GetOuter().inner.bogus   ' Inner has no `bogus` — chain after function call must keep type
End Sub
