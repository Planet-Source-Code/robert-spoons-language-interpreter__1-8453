Attribute VB_Name = "ModLanguageIFTHEN"
Public Sub Conditional(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(3) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(3) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(3))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 = ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalS(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
        ltmpVal2 = LTrim(AllItems(4))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 = ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalNE(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(4) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(4) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(4))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 <> ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalNES(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
    
        ltmpVal2 = LTrim(AllItems(5))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 <> ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalG(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(3) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(3) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(3))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 > ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalGS(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
    
        ltmpVal2 = LTrim(AllItems(4))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 > ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalL(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(3) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(3) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(3))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 < ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalLS(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
    
        ltmpVal2 = LTrim(AllItems(4))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 < ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalGE(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(4) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(4) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(4))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 >= ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalGES(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
    
        ltmpVal2 = LTrim(AllItems(5))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 >= ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalLE(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    If Not CheckVal(4) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(4) Then
                ltmpVal2 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal2 = Val(AllItems(4))
    End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 <= ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

Public Sub ConditionalLES(plint As Integer)
'this function executes the If Then keyword structure

Dim forCount As Integer
Dim tmpVal As Integer
Dim ltmpVal1 As Variant
Dim ltmpVal2 As Variant
Dim bRT As String
Dim MB(100) As Integer
Dim lCompareType As Integer

    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    Do
        forCount = forCount + 1
        
        'For To
        If UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(0)))) = KeyWord(0) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'If Then
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(3)))) = KeyWord(3) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Do While
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(19)))) = KeyWord(19) Then
            MB(tmpL + forCount) = Block
            Block = Block + 1
        'Loop
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(21)))) = KeyWord(21) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'End (If)
        ElseIf UCase(Left(Trim(PL(plint + forCount).Line), Len(KeyWord(5)))) = KeyWord(5) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        'Next
        ElseIf UCase(Trim(PL(plint + forCount).Line)) = KeyWord(2) Then
            If Block = MasterBlock Then lstop = True
            MB(tmpL + forCount) = Block
            Block = Block - 1
        
        Else
            MB(tmpL + forCount) = Block
        End If
    Loop Until lstop
    
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    CurLine = CurLine + forCount
    tmpVal = Val(AllItems(1))
    
    'Check for Value of 1st Param
    If Not CheckVal(1) Then
        For i = 0 To MemLimit
            If VarMem(i) = AllItems(1) Then
                ltmpVal1 = ValMem(i) 'set param
                Exit For
            End If
        Next
    Else
        ltmpVal1 = Val(AllItems(1))
    End If
    
    'tmpVal = Val(AllItems(3))
    
    'Check for Value of 2nd Param
    'If Not CheckVal(3) Then
    '    For i = 0 To MemLimit
    '        If VarMem(i) = AllItems(3) Then
    '            ltmpVal2 = ValMem(i) 'set param
    '            Exit For
    '        End If
    '    Next
    'Else
    
        ltmpVal2 = LTrim(AllItems(5))
    'End If
    
    'Perform If...Then...End If Command
    If ltmpVal1 <= ltmpVal2 Then 'code passthrough
        tmpL = PL(plint).LineNumber + 1
        DoEvents
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
    End If
    
    MasterBlock = MasterBlock - 1
    
End Sub

