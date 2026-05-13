Attribute VB_Name = "ReDimMemberValid"
Option Explicit

' #20 — `ReDim obj.member(...)` is valid when the UDT/class member is
' declared as a dynamic array. The chain walker must traverse the
' member access instead of looking up `member` as a standalone symbol.
Private Type Container
    items() As Variant
    counts() As Long
End Type

Private c As Container

Sub Resize(n As Long)
    ReDim c.items(1 To n)
    ReDim Preserve c.counts(0 To n - 1)
End Sub
