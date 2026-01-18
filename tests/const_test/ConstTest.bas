Attribute VB_Name = "ConstTest"
#Const ENABLE_LOGGING = True

Sub TestLog()
    #If ENABLE_LOGGING Then
        ' This should be active if #Const works
        Dim x As Integer
        x = "String" ' Type mismatch (though my analyzer only checks var existence mostly, but let's use an undefined var)
        x = UndefinedVarIfTrue
    #Else
        ' This should be active if #Const fails (default False)
        y = UndefinedVarIfFalse
    #End If
End Sub
