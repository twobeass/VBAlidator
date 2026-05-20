Attribute VB_Name = "NamedNotFound"
Option Explicit

Sub Target(ByVal a As Long)
End Sub

Sub Caller()
    Call Target(b:=1)
End Sub
