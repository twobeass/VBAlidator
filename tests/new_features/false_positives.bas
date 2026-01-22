Attribute VB_Name = "FalsePositivesTest"
Sub TestInputBox()
    ' Should not fail with 7 args
    Dim s As String
    s = InputBox("Prompt", "Title", "Default", 100, 100, "Help", 0)
End Sub

Sub TestParamArray(ParamArray args() As Variant)
    ' Definition
End Sub

Sub CallParamArray()
    ' Should not fail with many args
    TestParamArray 1, 2, 3, 4, 5, 6, 7, 8
    TestParamArray
End Sub

Sub TestFSO()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile "a", "b"
    fso.DeleteFile "a"
    Dim f As Object
    Set f = fso.GetFile("a")
End Sub

Sub TestDict()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    Dim k As Variant
    k = d.Keys
    Dim i As Variant
    i = d.Items
    d.CompareMode = 1
End Sub
