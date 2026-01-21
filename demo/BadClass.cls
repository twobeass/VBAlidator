VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BadClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pValue As String

Public Property Get Value() As String
    Value = pValue
End Property

Sub ClassErrors()
    ' 1. Undefined Member on Me
    Me.InvalidMethod

    ' 2. Undefined Type
    Dim obj As UnknownType
    Set obj = Nothing
End Sub
