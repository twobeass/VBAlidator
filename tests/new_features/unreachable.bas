Attribute VB_Name = "UnreachableTest"
Sub TestUnreachable()
    Dim x As Integer
    x = 1
    Exit Sub
    x = 2 ' Should be flagged as unreachable
End Sub

Sub TestGoTo()
    GoTo MyLabel
    Debug.Print "Unreachable" ' Should be flagged
MyLabel:
    Debug.Print "Reachable"
End Sub

Sub TestEnd()
    End
    Debug.Print "Unreachable"
End Sub
