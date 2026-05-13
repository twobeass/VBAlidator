Attribute VB_Name = "ConstWithFunc"
Sub TestConst()
    Const X As Long = MsgBox("hello")  ' function call in Const RHS
End Sub
