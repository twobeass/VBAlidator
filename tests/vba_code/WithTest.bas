Attribute VB_Name = "WithTest"
Sub TestWith()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws
        .Name = "Test"
        With .Range("A1")
            .Value = 100
        End With
    End With
    
    Dim c As New Collection
    c.Add "Item"
End Sub
