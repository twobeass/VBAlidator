Attribute VB_Name = "VbBuiltins"
Option Explicit

' #19 — `CallByName` and the `VbCallType` enum are part of the VBA
' runtime and must be available without explicit declarations or
' references.
Sub UseCallByName(obj As Object)
    Dim result As Variant
    result = CallByName(obj, "DoIt", VbMethod, 42)
    result = CallByName(obj, "Name", VbGet)
    Call CallByName(obj, "Name", VbLet, "new value")
    Call CallByName(obj, "Child", VbSet, Nothing)

    ' Qualified enum access must work too.
    Dim kind As Long
    kind = VbCallType.vbMethod
    kind = VbCallType.vbGet Or VbCallType.vbMethod
End Sub
