Attribute VB_Name = "ModLanguageSpace"
Public Function PrintSpaces(plint As Integer) As String
Dim lInt As Integer
Dim lStr As String
    lStr = AllItems(2)
    If CheckVal(2) Then
        lInt = Val(lStr)
    Else
        lInt = FindItNow(2)
    End If
    
    lStr = Space(lInt)
    
    PrintSpaces = lStr
    
End Function
