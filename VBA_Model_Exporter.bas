Attribute VB_Name = "VBA_Model_Exporter"
Option Explicit

' Requires Reference to: Microsoft Scripting Runtime (for FileSystemObject)
' Recommended: TypeLib Information (TLI) for full export

Public Sub ExportModel()
    Dim fso As Object
    Dim jsonFile As Object
    Dim jsonPath As String
    Dim exportData As String
    Dim tliApp As Object
    Dim tliError As String
    
    ' Define output path
    If ThisDocument.Path = "" Then
        MsgBox "Please save the document first.", vbExclamation
        Exit Sub
    End If
    ' Visio's Path usually includes the trailing backslash, but we check to be safe
    jsonPath = ThisDocument.Path
    If Right(jsonPath, 1) <> "\" Then jsonPath = jsonPath & "\"
    jsonPath = jsonPath & "vba_model.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(jsonPath, True)
    
    ' Initialize TLI
    On Error Resume Next
    Set tliApp = CreateObject("TLI.TLIApplication")
    If Err.Number <> 0 Then
        tliError = Err.Description & " (ensure tlbinf32.dll is registered)"
    End If
    On Error GoTo 0
    
    ' Start JSON
    exportData = "{" & vbCrLf
    exportData = exportData & "  ""meta"": {" & vbCrLf
    exportData = exportData & "    ""generated_by"": ""VBA_Model_Exporter""," & vbCrLf
    If tliError <> "" Then
        exportData = exportData & "    ""tli_status"": ""error""," & vbCrLf
        exportData = exportData & "    ""tli_error"": """ & EscapeJson(tliError) & """" & vbCrLf
    Else
        exportData = exportData & "    ""tli_status"": ""success""" & vbCrLf
    End If
    exportData = exportData & "  }," & vbCrLf
    
    ' Export References
    exportData = exportData & ExportReferences()
    
    exportData = exportData & "  ""globals"": {" & vbCrLf
    
    ' Export Application/Global context
    exportData = exportData & "    ""Application"": { ""type"": ""Application"" }," & vbCrLf
    exportData = exportData & "    ""ActivePage"": { ""type"": ""Page"" }," & vbCrLf
    exportData = exportData & "    ""ActiveDocument"": { ""type"": ""Document"" }," & vbCrLf
    exportData = exportData & "    ""ActiveWindow.Selection"": { ""type"": ""Selection"" }" & vbCrLf
    
    exportData = exportData & "  }," & vbCrLf
    exportData = exportData & "  ""classes"": {" & vbCrLf
    
    If Not tliApp Is Nothing Then
        exportData = exportData & ExportTLI(tliApp)
    Else
        exportData = exportData & ExportBasic()
    End If
    
    exportData = exportData & "  }" & vbCrLf
    exportData = exportData & "}"
    
    jsonFile.Write exportData
    jsonFile.Close
    
    MsgBox "Model exported to: " & jsonPath, vbInformation
End Sub

Private Function ExportReferences() As String
    Dim ref As Object
    Dim s As String
    Dim isFirst As Boolean
    
    s = "  ""references"": [" & vbCrLf
    isFirst = True
    
    On Error Resume Next
    For Each ref In Application.VBE.ActiveVBProject.References
        If Not isFirst Then s = s & "," & vbCrLf
        isFirst = False
        
        s = s & "    {" & vbCrLf
        s = s & "      ""name"": """ & ref.Name & """," & vbCrLf
        s = s & "      ""fullpath"": """ & EscapeJson(ref.FullPath) & """," & vbCrLf
        s = s & "      ""guid"": """ & ref.Guid & """," & vbCrLf
        s = s & "      ""major"": " & ref.Major & "," & vbCrLf
        s = s & "      ""minor"": " & ref.Minor & vbCrLf
        s = s & "    }"
    Next
    On Error GoTo 0
    
    s = s & vbCrLf & "  ]," & vbCrLf
    ExportReferences = s
End Function

Private Function ExportBasic() As String
    ' Fallback manual definitions
    ' Fallback manual definitions
    Dim s As String
    s = s & "    ""Page"": { ""members"": { ""Shapes"": { ""type"": ""Shapes"" }, ""Name"": { ""type"": ""String"" } } }," & vbCrLf
    s = s & "    ""Shape"": { ""members"": { ""Text"": { ""type"": ""String"" }, ""NameID"": { ""type"": ""String"" } } }" & vbCrLf
    ExportBasic = s
End Function

Private Function ExportTLI(tliApp As Object) As String
    Dim ref As Object
    Dim typeLib As Object
    Dim s As String
    Dim typeInfo As Object
    Dim member As Object
    Dim memberType As String
    Dim isFirstMember As Boolean
    Dim defaultInterface As Object
    Dim memberItem As Object
    
    s = ""
    
    ' Iterate references in current project
    ' Note: Requires 'Trust access to the VBA project object model'
    On Error Resume Next
    For Each ref In Application.VBE.ActiveVBProject.References
        Set typeLib = tliApp.TypeLibInfoFromFile(ref.FullPath)
        If Not typeLib Is Nothing Then
            
            ' Iterate Classes/Interfaces
            For Each typeInfo In typeLib.CoClasses
                 s = s & "    """ & typeInfo.Name & """: { ""members"": {"
                 
                 ' Iterate Members (simplified)
                 ' Note: TLI is complex, this is a sketch
                 isFirstMember = True
                 
                 ' Get default interface
                 Set defaultInterface = typeInfo.DefaultInterface
                 If Not defaultInterface Is Nothing Then
                     For Each memberItem In defaultInterface.Members
                         If Not isFirstMember Then s = s & ","
                         isFirstMember = False
                         
                         ' Determine return type
                         memberType = "Variant"
                         If Not memberItem.ReturnType Is Nothing Then
                             ' If it's an object, get name, else use VarType
                             memberType = "Object" ' Simplified
                         End If
                         
                         s = s & " """ & memberItem.Name & """: { ""type"": """ & memberType & """ }"
                     Next
                 End If
                 
                 s = s & " } }," & vbCrLf
            Next
            
            ' Also Interfaces
            For Each typeInfo In typeLib.Interfaces
                 s = s & "    """ & typeInfo.Name & """: { ""members"": {"
                 isFirstMember = True
                 For Each member In typeInfo.Members
                     If Not isFirstMember Then s = s & ","
                     isFirstMember = False
                     s = s & " """ & member.Name & """: { ""type"": ""Variant"" }"
                 Next
                 s = s & " } }," & vbCrLf
            Next
            
        End If
    Next
    On Error GoTo 0
    
    ' Trim last comma if needed
    If Right(s, 3) = "," & vbCrLf Then
        s = Left(s, Len(s) - 3) & vbCrLf
    End If
    ExportTLI = s
End Function

Private Function EscapeJson(text As String) As String
    Dim res As String
    res = Replace(text, "\", "\\")
    res = Replace(res, """", "\""")
    EscapeJson = res
End Function
