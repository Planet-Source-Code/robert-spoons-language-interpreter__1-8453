Attribute VB_Name = "ModLngValidation"


Public Function CheckVal(lInt As Integer) As Boolean
'This fuction Checks If item is a Litereal Value or not
'   it will return True if item is a Litteral,
'   or False if item is a Varable.
'
Dim TmpV As String
Dim bReturn As Boolean
On Error GoTo ErrHere
    TmpV = AllItems(lInt)
    If Val(TmpV) = 0 Then
        If TmpV = "0" Then
            bReturn = True
        End If
    ElseIf Len(CStr(Val(TmpV))) = Len(TmpV) Then
        bReturn = True
    End If
    
    CheckVal = bReturn
    x = 5
Exit Function
ErrHere:
ErrString = "Err #400   ERROR IN CHECKVAL (ModLngValidation)"
End Function

'Public Function evaluate(lInt As Variant) As Variant
'Dim lInt As Integer
'Dim itm As Integer
'Dim itmnum As Integer

'If CheckVal(lInt) Then
'        For i = 0 To MemLimit - 1
'            If VarMem(i) = AllItems(0) Then
'                ValMem(i) = Val(AllItems(2))
'                vFound = True
'                itm = i
'                Exit For
'            End If
'        Next
'
'        If Not vFound Then
'            VarMem(MemLimit) = AllItems(0)
'            ValMem(MemLimit) = Val(AllItems(2))
'            itm = MemLimit
'            MemLimit = MemLimit + 1
'        End If
'
'    Else 'Value is a Variable
'        For i = 0 To MemLimit - 1
'            If VarMem(i) = AllItems(2) Then
'                itmFound = True
'                itmnum = i
'                itmval = ValMem(i)
'                Exit For
'            End If
'        Next
'        'x = ValMem(itmnum)
'        If itmnum > -1 Then
'            For i = 0 To MemLimit - 1
'                If VarMem(i) = AllItems(0) Then
'                    ValMem(i) = ValMem(itmnum)
'                    vFound = True
'                    itm = i
'                    Exit For
'                End If
'            Next
'            If Not vFound Then
'                VarMem(MemLimit) = AllItems(0)
'                ValMem(MemLimit) = ValMem(itmnum)
'                itm = MemLimit
'                MemLimit = MemLimit + 1
'            End If
'        Else
'            For i = 0 To MemLimit - 1
'                If VarMem(i) = AllItems(0) Then
'                    ValMem(i) = 0
'                    vFound = True
'                    itm = i
'                    Exit For
'                End If
'            Next
'            If Not vFound Then
'                VarMem(MemLimit) = AllItems(0)
'                ValMem(MemLimit) = 0
'                itm = MemLimit
'                MemLimit = MemLimit + 1
'            End If
'        End If'
'
'End Function

Public Function CheckValue(lStr As Integer) As Boolean
'same as CheckVal, but used with string arg
Dim TmpV As String
Dim bReturn As Boolean
On Error GoTo ErrHere
    TmpV = lStr
    If Val(TmpV) = 0 Then
        If TmpV = "0" Then
            bReturn = True
        End If
    ElseIf Len(CStr(Val(TmpV))) = Len(TmpV) Then
        bReturn = True
    End If
    
    CheckValue = bReturn
Exit Function
ErrHere:
ErrString = "Err #401   ERROR IN CHECKVALUE (ModLngValidation)"
End Function

Public Function ReturnValue(lInt As Integer) As Integer
'returns an items value
Dim TmpV As String
Dim bReturn As Integer
Dim lvFound As Boolean
On Error GoTo ErrHere
    TmpV = AllItems(lInt)
    If Val(TmpV) = 0 Then
        If TmpV = "0" Then
            bReturn = 0
            lvFound = True
        End If
    ElseIf Len(CStr(Val(TmpV))) = Len(TmpV) Then
        bReturn = Val(TmpV)
        lvFound = True
    End If
    
    If Not lvFound Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(lInt) Then
                bReturn = ValMem(i)
                lvFound = True
                Exit For
            End If
        Next
        If Not lvFound Then
            VarMem(MemLimit) = AllItems(lInt)
            ValMem(MemLimit) = 0
            bReturn = 0
            MemLimit = memlimt + 1
        End If
    End If
    
    ReturnValue = bReturn
Exit Function
ErrHere:
ErrString = "Err #402   ERROR IN RETURNVALUE (ModLngValidation)"
End Function

Public Function FindItNow(plint As Integer) As Integer
'returns a vars value
Dim ret As Integer
On Error GoTo ErrHere
For i = 0 To MemLimit
        If VarMem(i) = AllItems(plint) Then
            ret = ValMem(i)
            lvFound = True
            itm = i
            Exit For
        End If
    Next
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(plint)
        ValMem(MemLimit) = 0
        TypMem(MemLimit) = "Integer"
        ret = 0
    End If
    
    FindItNow = ret
Exit Function
ErrHere:
ErrString = "Err #403   ERROR IN FINDITNOW (ModLngValidation)"
End Function

Public Function StoreItNow(plint As Integer, ret As Integer)
'stores a var
Dim lvFound As Boolean
On Error GoTo ErrHere
    For i = 0 To MemLimit
        If VarMem(i) = AllItems(plint) Then
            ValMem(i) = ret
            lvFound = True
            itm = i
            Exit For
        End If
    Next
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(plint)
        ValMem(MemLimit) = ret
        TypMem(MemLimit) = "Integer"
    End If
Exit Function
ErrHere:
ErrString = "Err #404   ERROR IN STOREITNOW (ModLngValidation)"
End Function

