Attribute VB_Name = "MissingProcLevel"
' EXPECTED_ERRORS: 1
' RaiseEvent SomethingHappened: event handler validation is Phase 3 (P3.2);
' until then the analyzer reports 'SomethingHappened' as an undefined
' identifier. Drive this baseline to 0 once Event/RaiseEvent is wired.
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
