Attribute VB_Name = "MissingPtrSafe"
Option Explicit

' On 64-bit Office (the modern default) this Declare needs PtrSafe.
Private Declare Function GetTickCount Lib "kernel32" () As Long

Sub TestApi()
End Sub
