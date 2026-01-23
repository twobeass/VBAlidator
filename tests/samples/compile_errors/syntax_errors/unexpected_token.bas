Attribute VB_Name = "UnexpectedToken"
Sub TestUnexpected()
    Dim x As Integer
    x = 10 @ 5 ' @ is not valid here
End Sub
