VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Phase2PropertyArity"

Public Property Get Name() As String
    Name = "x"
End Property

Public Property Let Name(ByVal RHS As String)
End Property

' Indexed Property: Get takes one arg, Let must take 2.
Public Property Get Value(ByVal Key As String) As Variant
    Value = Nothing
End Property

Public Property Let Value(ByVal Key As String, ByVal RHS As Variant)
End Property

Public Property Set Value(ByVal Key As String, ByVal RHS As Object)
End Property
