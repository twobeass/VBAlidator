Attribute VB_Name = "Module2"
Option Explicit

Sub BadSub()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Error: Invalid member
    ws.InvalidProp = 1
    
    ' Error: Unknown identifier
    y = 20
    
    ' Error: Unknown function
    Call UnknownFunc()
End Sub
