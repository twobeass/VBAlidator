Attribute VB_Name = "ControlFlowPhase1"
' Exercises Phase 1 control-flow coverage: For/For Each, Do, While/Wend,
' Select Case, ReDim, Erase. All identifiers are properly declared so the
' analyzer must report 0 errors.

Sub TestControlFlow()
    Dim i As Long
    Dim total As Long
    Dim arr() As Long
    Dim items As Variant

    ReDim arr(1 To 5)
    For i = 1 To 5
        arr(i) = i * 2
    Next i

    ReDim Preserve arr(1 To 10)

    For Each items In arr
        total = total + items
    Next items

    Do While total > 0
        total = total - 1
        If total = 5 Then Exit Do
    Loop

    Do
        total = total + 1
    Loop Until total >= 10

    While total < 20
        total = total + 1
    Wend

    Select Case total
        Case 1, 2, 3
            total = 0
        Case Is < 10
            total = 10
        Case 10 To 20
            total = 20
        Case Else
            total = -1
    End Select

    Erase arr
End Sub

Sub TestStringSuffix()
    Dim s As String
    s = Left$("hello", 3)
    s = Right$(s, 2)
    s = Mid$(s, 1, 1)
    s = UCase$(s)
    s = Trim$(s)
End Sub
