VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressionForm
   Caption         =   "RegressionForm"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub TestImplicitControl()
    ' This should NOT be an error in a Form, as it might be a Control
    ' that the parser failed to extract (or we want to be permissive).
    txtValue.Text = "Hello"

    ' But this SHOULD be an error (Strict Member Check) if ExistingModule exists
    ExistingModule.MissingFunc
End Sub
