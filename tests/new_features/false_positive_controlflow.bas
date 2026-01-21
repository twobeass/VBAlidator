Attribute VB_Name = "ReproControlFlow"
Sub TestBlockIf()
    Dim x As Integer
    x = 1
    If x = 1 Then
        Exit Sub
        x = 2 ' Should be Unreachable
    End If
    x = 3 ' Should be Reachable (False Positive currently)
End Sub

Sub TestSingleLineIf()
    Dim y As Integer
    If y = 0 Then y = 1: Exit Sub
    y = 2 ' Should be Reachable (False Positive currently)
End Sub

Sub TestLoop()
    Dim i As Integer
    For i = 1 To 10
        If i = 5 Then Exit Sub
    Next i
    i = 0 ' Should be Reachable
End Sub
