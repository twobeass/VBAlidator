Attribute VB_Name = "CreateObjectTest"
Sub TestCreateObject()
    ' Case 1: Variable assignment (remains Object/Variant per declaration)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.InvalidMethod ' This is allowed on Object (Late Binding)

    ' Case 2: With Block (Should use inferred type)
    With CreateObject("Scripting.Dictionary")
        .Add "Key", "Value"
        .InvalidMethod ' Should fail here!
    End With

    ' Case 3: Chained Call (Should use inferred type)
    CreateObject("Scripting.Dictionary").InvalidMethod ' Should fail here!

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile "test.txt"
End Sub
