Attribute VB_Name = "VariantTest"
Sub Main()
    Dim i As Integer
    i = 10

    ' Should OK
    Call TakesVariant(i)
End Sub

Sub TakesVariant(ByRef v As Variant)
End Sub
