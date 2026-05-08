VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "WrongLetArity"

Public Property Get Foo() As Long
    Foo = 1
End Property

' Property Let must have Get-arg-count + 1 = 1 param. Two params is wrong.
Public Property Let Foo(ByVal a As Long, ByVal b As Long)
End Property
