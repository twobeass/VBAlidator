Attribute VB_Name = "ObjectRepro"

Sub Main()
    Dim c As Collection
    Set c = New Collection

    Dim o As Object
    Set o = c

    ' Error: Expects Object, got Collection
    Call TakesObject(c)

    ' OK: Expects Object, got Object
    Call TakesObject(o)

    ' Error: Expects Collection, got Object
    Call TakesCollection(o)

    ' OK: Expects Collection, got Collection
    Call TakesCollection(c)
End Sub

Sub TakesObject(ByRef obj As Object)
End Sub

Sub TakesCollection(ByRef col As Collection)
End Sub
