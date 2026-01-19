Attribute VB_Name = "VBA_Model_Exporter"
' ==========================================================================================
' VBA Reference Exporter Tool
' ==========================================================================================
' Purpose: Exports the list of referenced libraries in the current VBA Project.
'          This is Step 1 of generating the dynamic object model.
'
' INSTRUCTIONS:
' 1. Import this module into your project
' 2. Run the 'ExportReferences' macro
' 3. A 'vba_references.json' file will be created.
' 4. Run 'python tools/generate_model.py' to generate the full model.
' ==========================================================================================

Option Explicit

Public Sub ExportReferences()
    Dim json As String
    json = "{" & vbCrLf
    json = json & "  ""references"": [" & vbCrLf
    
    Dim ref As Object ' Reference
    Dim i As Long
    Dim isFirst As Boolean
    isFirst = True
    
    For Each ref In ThisDocument.VBProject.References
        If Not isFirst Then json = json & "," & vbCrLf
        
        Dim refPath As String
        On Error Resume Next
        refPath = ref.FullPath
        On Error GoTo 0
        
        ' Escape backslashes
        refPath = Replace(refPath, "\", "\\")
        
        json = json & "    {"
        json = json & " ""name"": """ & ref.Name & ""","
        json = json & " ""guid"": """ & ref.GUID & ""","
        json = json & " ""major"": " & ref.Major & ","
        json = json & " ""minor"": " & ref.Minor & ","
        json = json & " ""path"": """ & refPath & """"
        json = json & " }"
        
        isFirst = False
    Next ref
    
    json = json & vbCrLf & "  ]" & vbCrLf
    json = json & "}"
    
    ' Save to file - use references.json initially
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outFile As String
    outFile = ThisDocument.Path & "\vba_references.json"
    
    Dim stream As Object
    Set stream = fso.CreateTextFile(outFile, True)
    stream.Write json
    stream.Close
    
    MsgBox "References exported to: " & outFile & vbCrLf & vbCrLf & "Now run 'python tools/generate_model.py' to complete the process.", vbInformation
End Sub
