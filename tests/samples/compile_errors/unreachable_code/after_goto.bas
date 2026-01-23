Attribute VB_Name = "UnreachableGoTo"
Sub TestUnreachableGoTo()
    GoTo MyLabel
    Debug.Print "Unreachable"
MyLabel:
    Debug.Print "Reachable"
End Sub
