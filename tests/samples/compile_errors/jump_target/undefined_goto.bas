Attribute VB_Name = "UndefinedGoto"
Sub TestGoto()
    GoTo NotALabel
    Exit Sub
RealLabel:
    Debug.Print "ok"
End Sub
