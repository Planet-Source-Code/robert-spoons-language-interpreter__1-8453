Attribute VB_Name = "ModLanguageDOLOOP"
Public Sub DoLoop(plint As Integer)
'this function executes the Do While keyword structure

Dim forCount As Integer
Dim tmpL As Integer
Dim x As Integer
Dim y As Integer
Dim SubSide As Integer
Dim FoundI As Integer
Dim i As Integer
Dim j As Integer
Dim MB(100) As Integer
Dim lb As Boolean

    'The required Initializations for Recursion
    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    'The Block Analyizer
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
    
    'The Required Assignments
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    
    FoundI = FindIndex(AllItems(2)) 'find the index of var
    x = ValMem(FoundI) 'get val from var
    y = ReturnValue(4) 'get arg value
    
    'The Recursive Bolck Interpreter Controler
    Do While x = y ' code passthrough
    DoEvents
        tmpL = PL(plint).LineNumber + 1 'current line
        For j = tmpL To tmpL + forCount 'current block
            If MB(j) = MasterBlock Then 'line is member of block
                Interpret j
            End If
        Next
        x = ValMem(FoundI) 'refresh val from var
        If StopProgram Then Exit Sub
    Loop
    
    MasterBlock = MasterBlock - 1
End Sub

Public Sub DoLoopU(plint As Integer)
'this function executes the DO UNTIL keyword structure

Dim forCount As Integer
Dim tmpL As Integer
Dim x As Integer
Dim y As Integer
Dim SubSide As Integer
Dim FoundI As Integer
Dim i As Integer
Dim j As Integer
Dim MB(100) As Integer
Dim lb As Boolean

    'The required Initializations for Recursion
    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    'The Block Analyizer
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
    
    'The Required Assignments
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    
    FoundI = FindIndex(AllItems(2))
    x = ValMem(FoundI)
    y = ReturnValue(4)
    
    'The Recursive Bolck Interpreter Controler
    Do While x <> y
    DoEvents
        tmpL = PL(plint).LineNumber + 1
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
        x = ValMem(FoundI)
        If StopProgram Then Exit Sub
    Loop
    
    MasterBlock = MasterBlock - 1
End Sub

Public Sub DoLoops(plint As Integer)
Dim forCount As Integer
Dim tmpL As Integer
Dim x As String
Dim y As String
Dim SubSide As Integer
Dim FoundI As Integer
Dim i As Integer
Dim j As Integer
Dim MB(100) As Integer
Dim lb As Boolean

    'The required Initializations for Recursion
    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    'The Block Analyizer
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
    
    'The Required Assignments
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    
    FoundI = FindIndex(AllItems(2))
    x = Right(ValMem(FoundI), Len(ValMem(FoundI)) - 1)
    y = AllItems(5)
    
    'The Recursive Bolck Interpreter Controler
    
    Do While x = y
    v = 5
    DoEvents
        tmpL = PL(plint).LineNumber + 1
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
        x = ValMem(FoundI)
        If StopProgram Then Exit Sub
    Loop
    
    MasterBlock = MasterBlock - 1
End Sub

Public Sub DoLoopUs(plint As Integer)
'this function executes the For Next keyword structure

Dim forCount As Integer
Dim tmpL As Integer
Dim x As String
Dim y As String
Dim SubSide As Integer
Dim FoundI As Integer
Dim i As Integer
Dim j As Integer
Dim MB(100) As Integer
Dim lb As Boolean

    'The required Initializations for Recursion
    bInBlock = True
    MasterBlock = MasterBlock + 1
    Block = MasterBlock
    tmpL = PL(plint).LineNumber
    forCount = 0
    
    'The Block Analyizer
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
    
    'The Required Assignments
    If MasterBlock = 1 Then
        LoopCountDown = forCount
    End If
    forCount = forCount - 1
    bInBlockCount = forCount
    
    FoundI = FindIndex(AllItems(2))
    x = Right(ValMem(FoundI), Len(ValMem(FoundI)) - 1)
    y = AllItems(5)
    
    'The Recursive Bolck Interpreter Controler
    
    Do While x <> y
    v = 5
    DoEvents
        tmpL = PL(plint).LineNumber + 1
        For j = tmpL To tmpL + forCount
            If MB(j) = MasterBlock Then
                Interpret j
            End If
        Next
        x = ValMem(FoundI)
        If StopProgram Then Exit Sub
    Loop
    
    MasterBlock = MasterBlock - 1
End Sub

Private Function FindIndex(lStr As String) As Integer
Dim lReturn As Integer
Dim lvFound As Boolean
    
a = VarMem(0)
b = VarMem(1)
cc = VarMem(2)
d = VarMem(3)
    For i = 0 To MemLimit
        If VarMem(i) = lStr Then
            lReturn = i
            lvFound = True
            Exit For
        End If
    Next
    
    If Not lvFound Then
        lvFound = False
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = lStr
        ValMem(MemLimit) = 0
        TypMem(memlimt) = "Integer"
        lReturn = MemLimit
    End If
    
    FindIndex = lReturn

End Function


