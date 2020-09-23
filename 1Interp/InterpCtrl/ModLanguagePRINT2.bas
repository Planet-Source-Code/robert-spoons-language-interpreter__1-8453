Attribute VB_Name = "ModLanguagePRINT"
Public Function PrintAValue() As String
'This function executes the Print keyword structure
'   with a numeric or variable argument
Dim ltmpAlone As Variant
Dim lRets As String
Dim lret As Integer
Dim i As Integer
Dim lvFound As Integer
Dim tmpi As Integer
Dim sReturn As String

    For i = 0 To MemLimit '- 1
        If VarMem(i) = AllItems(1) Then
            If TypMem(i) = "String" Then
                lRets = ValMem(i)
                lvFound = 1
                Exit For
            ElseIf TypMem(i) = "Integer" Then
                lret = ValMem(i)
                lvFound = 2
                tmpi = i
                Exit For
            End If
        End If
    Next
    If lvFound <> 1 Then
        If Not CheckVal(1) Then
            For i = 0 To MemLimit '- 1
                If VarMem(i) = AllItems(1) Then
                    lret = ValMem(i)
                    vFound = True
                Exit For
                End If
            Next
    
            If Not vFound Then
                lret = 0
            End If
        Else
            lret = Val(AllItems(1))
        End If
    End If
    
    If lvFound = 1 Then
        sReturn = lRets
    Else
        sReturn = CStr(lret)
    End If

    PrintAValue = sReturn
    
End Function

Public Function PrintAString() As String
'This function executes the Print keyword structure
'   with a string argument

Dim lStr As String
Dim i As Integer
Dim allIFlag As Boolean

    If AllItems(1) = Chr(34) Then
        For i = 2 To StructureLen
            If Not allIFlag Then
                If AllItems(i) <> Chr(34) Then
                    lStr = lStr + " " + AllItems(i)
                Else
                    allIFlag = True
                End If
            End If
        Next
    End If
    
    PrintAString = lStr
    v = AllItems(2)
    z = AllItems(3)
    v = 5
End Function

