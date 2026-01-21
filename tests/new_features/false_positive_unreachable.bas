Attribute VB_Name = "ReproUnreachable"
Sub TestErrorHandling()
    On Error GoTo ErrHandler

    ' Do something
    Dim x As Integer
    x = 1 / 0

CleanExit:
    Exit Sub
ErrHandler:
    Debug.Print "Error detected"
    Resume CleanExit
End Sub

Sub TestStandardExit()
    Exit Sub
    Debug.Print "This is actually unreachable"
End Sub
