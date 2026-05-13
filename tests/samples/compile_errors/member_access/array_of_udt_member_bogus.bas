Attribute VB_Name = "ArrayUdtMember"
Option Explicit

Private Type Cell
    val As Long
End Type

Private Type Row
    cells() As Cell
End Type

Sub IndexedMidChain()
    Dim r As Row
    Dim n As Long
    n = r.cells(0).bogus   ' Cell has no `bogus` — array index in chain must preserve element type
End Sub
