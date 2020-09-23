Attribute VB_Name = "ModLanguageASSIGN"
'This is the assignment module
    
'End Sub

Public Sub AssignTo(plint As Integer)
'this function executes the assignment structure
'for integers

Dim ret As Integer
Dim i As Integer
Dim lvFound As Boolean
    
    If AllItems(2) = "(" Then
        ret = ArithPar(plint)
    Else
        ret = Arithmetic(plint)
    End If
    
    For i = 0 To MemLimit '- 1
        If VarMem(i) = AllItems(0) Then
            ValMem(i) = ret
            lvFound = True
            itm = i
            Exit For
        End If
    Next
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(0)
        ValMem(MemLimit) = ret
        TypMem(MemLimit) = "Integer"
    End If
    
End Sub

Public Sub AssignStrTo(plint As Integer)
'this function executes the assignment structure
'for strings

Dim ret As String
Dim i As Integer
Dim lvFound As Boolean
Dim QuoteCount As Integer

    For i = 2 To StructureLen
        If AllItems(i) = Chr(34) Then
            QuoteCount = QuoteCount + 1
        End If
        If AllItems(i) <> Chr(34) Then
            If QuoteCount < 2 Then
                lStr = lStr + " " + AllItems(i)
            End If
        End If
    Next
        
    For i = 0 To MemLimit
        If VarMem(i) = AllItems(0) Then
            ValMem(i) = lStr
            TypMem(i) = "String"
            lvFound = True
            itm = i
            Exit For
        End If
    Next
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(0)
        ValMem(MemLimit) = lStr
        TypMem(MemLimit) = "String"
    End If
    
End Sub

