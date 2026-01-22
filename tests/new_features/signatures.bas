Attribute VB_Name = "SignaturesTest"
Sub TestSignatures()
    MsgBox "Hello"
    MsgBox "Hello", 1, "Title", "Help", 1000, "Extra" ' Invalid (6 args, max 5)
    MsgBox() ' Invalid (0 args, min 1)
End Sub
