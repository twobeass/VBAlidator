Attribute VB_Name = "Module1"
Option Explicit

Public GlobalVar As Integer

Sub TestSub()
    Dim x As Integer
    x = 10
    MsgBox "Hello"
    Debug.Print x
End Sub

Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function
