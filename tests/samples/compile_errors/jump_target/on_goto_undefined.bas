Attribute VB_Name = "OnGoToBad"
Option Explicit

' #18 — When the computed-GoTo label list references an undeclared
' label, VBA201 must still fire (one entry resolves, one does not).
Sub Bad()
    Dim op As Long
    op = 1
    On op GoTo realLabel, missingLabel
    Exit Sub
realLabel:
End Sub
