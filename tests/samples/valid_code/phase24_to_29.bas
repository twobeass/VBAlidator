Attribute VB_Name = "Phase24To29Valid"
DefInt I-N
DefStr S

' Module-level fixed-length string is fine.
Public name As String * 50

Sub TestDates()
    Dim d As Date
    d = #2020-01-01#
    d = #1/1/2020 12:00:00 AM#
    d = #January 1, 2020#
    d = #1-Jan-2020#
End Sub

Sub TestConsts()
    Const MAX As Long = 100
    Const FACTOR As Double = 2.5 * 3
    Const GREETING As String = "Hello"
    Const RGB_RED As Long = vbRed
End Sub

Sub TestArith()
    Dim x As Long
    x = 5 * 3
    x = 10 - 4
    x = "abc" & 1     ' & is string concat — must NOT flag
    x = 1 + "2"       ' VBA coerces + bidirectionally — also not flagged
End Sub

Sub TestDefType()
    ' i, j, k, l, m, n implicitly typed as Integer (DefInt I-N)
    Dim i, j, k As Integer  ' explicit on k, the others default via DefInt
    i = 1
    j = 2
    k = i + j
    ' s implicitly typed as String (DefStr S)
    Dim s
    s = "hello"
End Sub
