Attribute VB_Name = "ModLanguageString"
Public Sub FindInString(plint As Integer)
Dim bReturn As Integer
Dim testString1 As String
Dim testString2 As String

    testString1 = AllItems(2)
    

End Sub

Public Sub ConcatTwoVars()
Dim lStr As String
Dim lStr2 As String
Dim lStr1 As String
Dim lStrFound As Boolean

lStr = AllItems(3)
For i = 0 To MemLimit
    If VarMem(i) = lStr Then
        If TypMem(i) = "String" Then
            lStr = ValMem(i)
            lStrFound = True
            Exit For
        End If
    End If
Next
If Not lStrFound Then
    lStr = ""
    VarMem(MemLimit) = lStr
    ValMem(MemLimit) = ""
    TypMem(MemLimit) = "String"
    MemLimit = MemLimit + 1
End If

lStrFound = False
lStr2 = AllItems(5)
For i = 0 To MemLimit
    If VarMem(i) = lStr2 Then
        If TypMem(i) = "String" Then
            lStr2 = ValMem(i)
            lStrFound = True
            Exit For
        End If
    End If
Next
If Not lStrFound Then
    lStr2 = ""
    VarMem(MemLimit) = lStr2
    ValMem(MemLimit) = ""
    TypMem(MemLimit) = "String"
    MemLimit = MemLimit + 1
End If

lStrFound = False
lStr1 = AllItems(0)
For i = 0 To MemLimit
    If VarMem(i) = lStr1 Then
        TypMem(i) = "String"
        ValMem(i) = lStr + lStr2
        lStrFound = True
        Exit For
    End If
Next
If Not lStrFound Then
    MemLimit = MemLimit + 1
    VarMem(MemLimit) = lStr1
    ValMem(MemLimit) = lStr + lStr2
    TypMem(MemLimit) = "String"
End If

End Sub
