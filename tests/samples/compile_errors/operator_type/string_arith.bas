Attribute VB_Name = "BadArith"
Sub TestArith()
    Dim x As Long
    x = "hello" - 1        ' string literal in arithmetic
    x = 5 * "abc"          ' string on right side
    x = "a" / "b"          ' both strings
End Sub
