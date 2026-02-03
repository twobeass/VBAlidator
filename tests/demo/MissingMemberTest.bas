Attribute VB_Name = "MissingMemberTest"
Sub TestCalls()
    ' This module exists, but the function does not.
    ' It should report "Member 'MissingFunc' not found".
    ExistingModule.MissingFunc
End Sub
