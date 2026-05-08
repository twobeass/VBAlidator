Attribute VB_Name = "Phase3Valid"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Public Enum Status
    Ready = 0
    Busy = 1
    Done = 2
End Enum

Sub Demo()
    Dim t As Long
    t = GetTickCount()
End Sub
