Attribute VB_Name = "ReproModule"

Sub Main()
    Dim i As Integer
    i = 10

    ' This should be an error: Expects Long ByRef, passing Integer variable
    Call MySub(i)

    ' This should be OK (ByVal implicit via parens)
    Call MySub((i))

    ' This should be OK (Literal)
    Call MySub(10)
End Sub

Sub MySub(ByRef l As Long)
    ' ...
End Sub
