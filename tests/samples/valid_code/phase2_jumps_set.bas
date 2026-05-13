Attribute VB_Name = "Phase2JumpsSet"
' Exercises Phase 2.1 (labels), 2.2 (Set/Let), 2.3 (property arity).
' All constructs are valid; the analyzer must report 0 errors.

Sub TestLabels()
    Dim cond As Boolean
    cond = True
    On Error GoTo Handler
    If cond Then GoTo Skip
    Debug.Print "before skip"
Skip:
    Debug.Print "after skip"
    Exit Sub
Handler:
    Resume Next
End Sub

Sub TestErrorReset()
    On Error GoTo 0
    On Error GoTo -1
    On Error Resume Next
End Sub

Sub TestSetLetCorrect()
    Dim s As String
    Dim n As Long
    Dim col As Object
    s = "hello"          ' scalar assignment, no Set
    n = 42               ' scalar assignment, no Set
    Set col = CreateObject("Scripting.Dictionary")  ' Object → Set
End Sub

Sub TestVariantBothWays()
    Dim v As Variant
    v = 1                ' variant accepts value assignment
    Set v = Nothing      ' variant accepts Set too — must not flag
End Sub
