Attribute VB_Name = "BuiltInTest"
Sub Main()
    Dim i As Integer
    i = 10

    Dim o As Object
    Set o = Nothing

    ' ObjPtr expects Object (ByRef strict per our manual update)

    ' OK
    Debug.Print ObjPtr(o)

    ' Error: Integer passed to ByRef Object
    Debug.Print ObjPtr(i)
End Sub
