Attribute VB_Name = "MissingModuleLevel"
' 1. DefType statements
DefInt A-Z

' 2. Implements Statement
Implements UnknownInterface

' 3. Event Declarations
Public Event SomethingHappened(ByVal msg As String)

' 4. Friend Scope
Friend Sub FriendMethod()
End Sub

Sub TestCall()
    FriendMethod ' Should be undefined if Friend Sub was ignored
End Sub
