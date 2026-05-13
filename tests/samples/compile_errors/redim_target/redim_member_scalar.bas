Attribute VB_Name = "ReDimMemberScalar"
Option Explicit

' #20 — Even after chain resolution, ReDim on a scalar member is still
' an error. VBA103 must fire because `Container.one` is a Long, not an
' array.
Private Type Container
    one As Long
End Type

Private c As Container

Sub Bad()
    ReDim c.one(1 To 5)
End Sub
