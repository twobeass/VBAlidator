Attribute VB_Name = "MemberNotFound"
Sub TestMember()
    Dim coll As Collection
    Set coll = New Collection
    coll.FooBar ' Collection does not have FooBar
End Sub
