Attribute VB_Name = "ConstFromVar"
Sub TestConst()
    Dim runtimeValue As Long
    runtimeValue = 42
    Const X As Long = runtimeValue  ' references runtime variable
End Sub
