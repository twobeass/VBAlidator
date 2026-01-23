Attribute VB_Name = "MissingThen"
Sub TestMissingThen()
    Dim x As Integer
    x = 10
    If x > 5 ' Missing 'Then'
        Debug.Print "Big"
    End If
End Sub
