Attribute VB_Name = "MissingProcLevel"
Sub TestProc()
    ' 5. GoSub / Return
    GoSub LocalLabel

    ' 6. RaiseEvent (Syntactically valid as statement, but analyzer?)
    ' Note: RaiseEvent only valid in Class/Form, but Parser should parse it.
    RaiseEvent SomethingHappened("Hello")

    ' 7. LSet / RSet
    Dim s1 As String, s2 As String
    LSet s1 = s2
    RSet s1 = s2

    ' 8. AddressOf Operator
    Dim p As LongPtr
    p = AddressOf TestProc

    Exit Sub
LocalLabel:
    Return
End Sub
