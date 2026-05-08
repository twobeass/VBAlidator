Attribute VB_Name = "VBA_Model_Exporter"
' ==========================================================================================
' VBA Reference Exporter — host-agnostic
' ==========================================================================================
' Purpose: Exports the list of referenced libraries in the current VBA
'          project as `vba_references.json`. This is step 1 of the
'          custom-object-model workflow:
'
'             VBA_Model_Exporter.bas    →  vba_references.json
'             tools/generate_model.py   →  vba_model.json
'             vbalidator … --model vba_model.json
'
' Compatible with Excel, Word, Access, PowerPoint, Outlook and any
' other host that exposes `Application.VBE.ActiveVBProject`.
'
' INSTRUCTIONS:
'   1. Open the VBE in your host (Alt+F11).
'   2. File → Import File… → select this `.bas`.
'   3. Run the `ExportReferences` macro (F5 with cursor inside the Sub).
'   4. A `vba_references.json` file is written next to the host file.
'      The exact path is shown in the resulting MsgBox.
'   5. Run `python tools/generate_model.py vba_references.json` to
'      produce the `vba_model.json` consumed by `vbalidator --model`.
'
' Trust-center note: requires "Trust access to the VBA project object
' model" enabled in the Office Trust Center, otherwise `VBE` is hidden.
' ==========================================================================================

Option Explicit

' ---- Public entry point ------------------------------------------------

Public Sub ExportReferences()
    Dim refs As Object
    Set refs = ResolveReferences()
    If refs Is Nothing Then
        MsgBox "Could not access the active VBA project. Enable " & _
               "'Trust access to the VBA project object model' in the " & _
               "Office Trust Center and try again.", _
               vbCritical, "VBAlidator: Reference Exporter"
        Exit Sub
    End If

    Dim outPath As String
    outPath = ResolveOutputPath()

    Dim json As String
    json = BuildJson(refs)

    WriteToFile outPath, json

    MsgBox "References exported to:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "Next: run `python tools/generate_model.py """ & outPath & """`.", _
           vbInformation, "VBAlidator: Reference Exporter"
End Sub

' ---- Implementation ----------------------------------------------------

' Resolve the active VBProject's References collection in a way that
' works across every Office host. Returns Nothing when access is
' blocked by the Trust Center.
Private Function ResolveReferences() As Object
    On Error Resume Next

    ' Path 1: VBE.ActiveVBProject — works in every host with VBE access.
    Dim vbe As Object
    Set vbe = Application.VBE
    If Not vbe Is Nothing Then
        If Not vbe.ActiveVBProject Is Nothing Then
            Set ResolveReferences = vbe.ActiveVBProject.References
            If Err.Number = 0 And Not ResolveReferences Is Nothing Then
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If
    Err.Clear

    ' Path 2: ThisWorkbook.VBProject (Excel)
    Set ResolveReferences = ThisWorkbook.VBProject.References
    If Err.Number = 0 And Not ResolveReferences Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear

    ' Path 3: ThisDocument.VBProject (Word / PowerPoint)
    Set ResolveReferences = ThisDocument.VBProject.References
    If Err.Number = 0 And Not ResolveReferences Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear

    On Error GoTo 0
    Set ResolveReferences = Nothing
End Function

' Best-effort destination for the JSON file. Falls back through:
'   1. Workbook / Document folder when one is available.
'   2. The current host's `Application.Path` (Office install dir is
'      writable on dev machines but not on locked-down ones — caveat).
'   3. The user's TEMP folder, which always works.
Private Function ResolveOutputPath() As String
    Dim folder As String
    folder = TryHostFolder()
    If LenB(folder) = 0 Then
        folder = Environ$("TEMP")
        If LenB(folder) = 0 Then folder = Environ$("USERPROFILE")
    End If
    ResolveOutputPath = folder & "\vba_references.json"
End Function

Private Function TryHostFolder() As String
    On Error Resume Next
    TryHostFolder = ThisWorkbook.Path
    If Err.Number = 0 And LenB(TryHostFolder) > 0 Then Exit Function
    Err.Clear

    TryHostFolder = ThisDocument.Path
    If Err.Number = 0 And LenB(TryHostFolder) > 0 Then Exit Function
    Err.Clear

    TryHostFolder = Application.CurrentProject.Path  ' Access
    If Err.Number = 0 And LenB(TryHostFolder) > 0 Then Exit Function
    Err.Clear

    On Error GoTo 0
    TryHostFolder = ""
End Function

Private Function BuildJson(ByVal refs As Object) As String
    Dim sb As String
    sb = "{" & vbCrLf & "  ""references"": [" & vbCrLf

    Dim ref As Object
    Dim isFirst As Boolean
    isFirst = True

    For Each ref In refs
        If Not isFirst Then sb = sb & "," & vbCrLf

        Dim refPath As String
        On Error Resume Next
        refPath = ref.FullPath
        On Error GoTo 0
        refPath = JsonEscape(refPath)

        sb = sb & "    {" & _
             " ""name"": """ & JsonEscape(ref.Name) & """," & _
             " ""guid"": """ & JsonEscape(ref.GUID) & """," & _
             " ""major"": " & ref.Major & "," & _
             " ""minor"": " & ref.Minor & "," & _
             " ""path"": """ & refPath & """" & _
             " }"

        isFirst = False
    Next ref

    sb = sb & vbCrLf & "  ]" & vbCrLf & "}"
    BuildJson = sb
End Function

Private Function JsonEscape(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, "\", "\\")
    out = Replace(out, """", "\""")
    out = Replace(out, vbCr, "\r")
    out = Replace(out, vbLf, "\n")
    out = Replace(out, vbTab, "\t")
    JsonEscape = out
End Function

Private Sub WriteToFile(ByVal path As String, ByVal contents As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stream As Object
    Set stream = fso.CreateTextFile(path, True, False)   ' ASCII; JSON only contains ASCII after escaping
    stream.Write contents
    stream.Close
End Sub
