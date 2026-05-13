Attribute VB_Name = "OnGoToValid"
Option Explicit

' #18 — Computed-GoTo jump table is valid VBA syntax. Every label in
' the comma-separated list is defined later in the same procedure.
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
