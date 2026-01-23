Attribute VB_Name = "TypeMismatch"
Sub TestMismatch()
    Dim i As Integer
    i = 5
    i(1) = 10 ' i is integer, not array or procedure
End Sub
