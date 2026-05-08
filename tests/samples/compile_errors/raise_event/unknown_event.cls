VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "RaiseUnknownEvent"
Option Explicit

' No `Event` declared — RaiseEvent below has no matching declaration.
Sub Trigger()
    RaiseEvent NotDeclared("hello")
End Sub
