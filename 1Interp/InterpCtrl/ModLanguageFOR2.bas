Attribute VB_Name = "ModLanguageFOR"
Public Function LngFOR(plint As Integer)
'this function executes the For Next keyword structure

Dim forCount As Integer
Dim tmpL As Integer
'Dim bInBlock As Boolean
Dim x As Integer
Dim y As Integer
Dim SubSide As Integer
Dim FoundI As Integer
Dim i As Integer
Dim j As Integer
Dim MB(100) As Integer

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
    
    x = ReturnValue(3) 'assign lower limit
    y = ReturnValue(5) 'assign upper limit
    
    FoundI = FindIndex(AllItems(1))
    
    'The Recursive Bolck Controler
    
    For i = x To y 'code passthrough
    DoEvents 'keep the enviorn alive
        ValMem(FoundI) = i
        tmpL = PL(plint).LineNumber + 1 'cur code line
        For j = tmpL To tmpL + forCount 'cur block
            If MB(j) = MasterBlock Then 'line in immediate block
                Interpret j
            End If
        Next
        If StopProgram Then Exit Function
    Next
    
    MasterBlock = MasterBlock - 1
    
End Function

Private Function FindIndex(lStr As String) As Integer
Dim lReturn As Integer
Dim lvFound As Boolean
    
    For i = 0 To MemLimit - 1
        If VarMem(i) = AllItems(1) Then 'found var
            lReturn = i
            lvFound = True
            Exit For
        End If
    Next
    
    If Not lvFound Then
        lvFound = False
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(1) 'var set
        ValMem(MemLimit) = 0
        TypMem(memlimt) = "Integer"
        lReturn = MemLimit
    End If
    
    FindIndex = lReturn
    
End Function


