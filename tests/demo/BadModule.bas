Attribute VB_Name = "BadModule"
Option Explicit

Sub DemoErrors()
    ' 1. Undefined Variable
    x = 10

    ' 2. Duplicate Declaration
    Dim y As Integer
    Dim y As String

    ' 3. Const Initialization Error (New Feature)
    Const PI = 3.14
    Const InvalidConst = UndefinedConstant * 2

    ' 4. Invalid Call (Calling a variable like a function)
    Dim z As Integer
    z = 5
    Dim res As Integer
    res = z(10)

    ' 5. Member access on non-object
    Dim i As Integer
    i.SomeMethod

    ' 6. Line continuation check
    Dim a As Integer
    a = 1 + _
        2
End Sub
