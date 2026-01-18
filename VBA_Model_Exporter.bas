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
    
    ' Define output path
    jsonPath = ThisWorkbook.Path & Application.PathSeparator & "vba_model.json"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(jsonPath, True)
    
    ' Initialize TLI
    On Error Resume Next
    Set tliApp = CreateObject("TLI.TLIApplication")
    On Error GoTo 0
    
    ' Start JSON
    exportData = "{" & vbCrLf
    exportData = exportData & "  ""meta"": {""generated_by"": ""VBA_Model_Exporter""}," & vbCrLf
    exportData = exportData & "  ""globals"": {" & vbCrLf
    
    ' Export Application/Global context
    exportData = exportData & "    ""Application"": { ""type"": ""Application"" }," & vbCrLf
    exportData = exportData & "    ""ActiveSheet"": { ""type"": ""Worksheet"" }," & vbCrLf
    exportData = exportData & "    ""ActiveWorkbook"": { ""type"": ""Workbook"" }," & vbCrLf
    exportData = exportData & "    ""Selection"": { ""type"": ""Object"" }" & vbCrLf
    
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

Private Function ExportBasic() As String
    ' Fallback manual definitions
    Dim s As String
    s = s & "    ""Worksheet"": { ""members"": { ""Range"": { ""type"": ""Range"" }, ""Name"": { ""type"": ""String"" } } }," & vbCrLf
    s = s & "    ""Range"": { ""members"": { ""Value"": { ""type"": ""Variant"" }, ""Address"": { ""type"": ""String"" } } }" & vbCrLf
    ExportBasic = s
End Function

Private Function ExportTLI(tliApp As Object) As String
    Dim ref As Object
    Dim typeLib As Object
    Dim s As String
    Dim typeInfo As Object
    Dim member As Object
    Dim memberType As String
    
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
                 Dim first As Boolean
                 first = True
                 
                 ' Get default interface
                 Dim defInt As Object
                 Set defInt = typeInfo.DefaultInterface
                 If Not defInt Is Nothing Then
                     Dim m As Object
                     For Each m In defInt.Members
                         If Not first Then s = s & ","
                         first = False
                         
                         ' Determine return type
                         memberType = "Variant"
                         If Not m.ReturnType Is Nothing Then
                             ' If it's an object, get name, else use VarType
                             memberType = "Object" ' Simplified
                         End If
                         
                         s = s & " """ & m.Name & """: { ""type"": """ & memberType & """ }"
                     Next
                 End If
                 
                 s = s & " } }," & vbCrLf
            Next
            
            ' Also Interfaces
            For Each typeInfo In typeLib.Interfaces
                 s = s & "    """ & typeInfo.Name & """: { ""members"": {"
                 first = True
                 For Each member In typeInfo.Members
                     If Not first Then s = s & ","
                     first = False
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
