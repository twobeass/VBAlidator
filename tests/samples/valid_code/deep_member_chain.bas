Attribute VB_Name = "DeepMemberChain"
Option Explicit

' P2.6 — These deep chains must all type-check cleanly. The analyser
' walks the dotted/indexed/call-resulting chain and resolves each hop
' against UDT/class/global metadata; if any hop loses its element type
' (e.g. an `()`-array member parsed as Variant) a downstream member
' lookup would either misfire or silently degrade. This fixture
' guards against that.

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

Private Type Inner
    n As Long
End Type

Private Type Outer
    inner As Inner
End Type

Function GetBook() As Book
End Function

Function GetOuter() As Outer
End Function

Sub UseDeepChains()
    Dim b As Book
    Dim o As Outer
    Dim n As Long

    n = b.tables(0).rows(0).cells(0).val          ' depth 5, array indices at every hop
    n = GetBook().tables(0).rows(0).cells(0).val  ' depth 5 starting from function call
    n = GetOuter().inner.n                         ' function return -> UDT member -> scalar
    n = o.inner.n                                  ' bare nested chain
End Sub
