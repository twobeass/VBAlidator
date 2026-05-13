Attribute VB_Name = "UndefinedResume"
Sub TestResume()
    On Error GoTo Handler
    Exit Sub
Handler:
    Resume RetryStep
End Sub
