Attribute VB_Name = "ArgMismatch"
Sub MySub(a As Integer)
End Sub

Sub TestArgs()
    MySub 1, 2 ' Expected 1 arg, got 2
End Sub
