Attribute VB_Name = "ModLngArithmetic"
Dim Ops1 As MyOps
Dim reg As Integer

Public Function Artith()
Dim tAll As Integer
Dim curPres As Integer
Dim LastPres As Integer
Dim curItem As String
Dim LastSymb As String
Dim CurSymb As String
Dim UseSymb As String
Dim lVal1 As Integer
Dim lVal2 As Integer
Dim ret As Integer

    tAll = StructureLen
    For i = 2 To tAll
        curItem = AllItems(i)
        If CurIsSymb(curItem) Then
            CurSymb = curItem
            If CurSymb = "(" Then curPres = 1
            If CurSymb = ")" Then curPres = 1
            If CurSymb = "^" Then curPres = 2
            If CurSymb = "*" Then curPres = 3
            If CurSymb = "/" Then curPres = 3
            If CurSymb = "+" Then curPres = 4
            If CurSymb = "-" Then curPres = 4
            
            If i = 2 Then
                LastPres = curPres
                LastSymb = CurSymb
                Push1 CurSymb
            End If
            
            If CurSymb = ")" Then
                UseSymb = Pop1
                Do Until UseSymb = "("
                    lVal2 = Pop2
                    lVal1 = Pop2
                    ret = DoMath(lVal1, lVal2, UseSymb)
                    Push2 ret
                    UseSymb = Pop1
                Loop
            Else
                If LastPres > curPres Then
                    Push1 CurSymb
                    LastSymb = CurSymb
                    LastPres = curPres
                Else
                    UseSymb = Pop1
                    lVal2 = Pop2
                    lVal1 = Pop2
                    ret = DoMath(lVal1, lVal2, UseSymb)
                    Push2 ret
                End If
            End If
        End If
    Next
    
End Function

Private Function CurIsSymb(lStr As String)
Dim cReturn As Boolean
    If lStr = "(" Then cReturn = True
    If lStr = ")" Then cReturn = True
    If lStr = "^" Then cReturn = True
    If lStr = "*" Then cReturn = True
    If lStr = "/" Then cReturn = True
    If lStr = "+" Then cReturn = True
    If lStr = "-" Then cReturn = True
    
    CurIsSymb = cReturn

End Function


Public Function Arithmetic(plint As Integer) As Variant 'Integer
Dim i As Integer
Dim testSymb As String
Dim testSymbVal As Integer
Dim sVal1 As String
Dim sVal2 As String
Dim tVal1 As Integer
Dim tVal2 As Integer
Dim tOps As String
Dim CurSymb As String
Dim curWord As Integer
Dim CurSymbCount As Integer
Dim CurCount As Integer
Dim NextSymb As String
Dim ret As Integer
Dim MathSymb As String

If NumNotReseved < 2 Then       'Straight assignment
    
    If Not CheckForVal(1) Then  'If Not An integer
        tVal1 = MemStore(1)     'New Variable
    Else
        tVal1 = Val(NotReserved(1))
    End If
    
    ret = tVal1

ElseIf NumNotReseved < 3 Then   'Binary operation
    
    If Not CheckForVal(1) Then  'If Not an Integer
        tVal1 = MemStore(1)     'New Variable
    Else
        tVal1 = Val(NotReserved(1))
    End If
    
    If Not CheckForVal(2) Then  'If Not an Integer
        tVal2 = MemStore(2)     'New Variable
    Else
        tVal2 = Val(NotReserved(2))
    End If
    tOps = GetSymbol(1)         'Set The Operator
    ret = DoMath(tVal1, tVal2, tOps)

Else                             'The good stuff
    testSymb = GetSymbol(1)     'Get Operator
    testSymbVal = GetSymVal(testSymb) 'Get Priority
    Push1 testSymb 'the fist op is now on the stack
    
    curWord = 1
    
    If Not CheckForVal(curWord) Then 'Make sure arg is integer
        tVal1 = MemStore(curWord)
    Else
        tVal1 = Val(NotReserved(curWord))
    End If
    Push2 tVal1                 'the 1st arg is stacked
    curWord = curWord + 1
    
    If Not CheckForVal(curWord) Then
        tVal2 = MemStore(curWord)
    Else
        tVal2 = Val(NotReserved(curWord))
    End If
    Push2 tVal2                 'the 2nd arg is stacked
    curWord = curWord + 1
    vv = SymbolCount
    For i = 2 To SymbolCount - 1 '1st sym is already on stack
        CurSymb = GetSymbol(i)   'get next symbol
        If CurSymb = ")" Then
            tOps = Pop1
            Do Until tOps = "(" 'GetSymbol(CurCount) = "("
            DoEvents
                tVal2 = Pop2
                tVal1 = Pop2
                ret = DoMath(tVal1, tVal2, tOps)
                Push2 ret
                tOps = Pop1
                If StopProgram Then Exit Function
            Loop
            If Stack1TopItem <> "(" Then
                If StackTop1 > -1 Then
                    tOps = Pop1
                    tVal2 = Pop2
                    tVal1 = Pop2
                    ret = DoMath(tVal1, tVal2, tOps)
                    testSymbVal = 1
                    Push2 ret
                End If
            End If
        ElseIf CurSymb = "(" Then
            Push1 CurSymb
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem
            
        ElseIf testSymbVal > GetSymVal(CurSymb) Then
            Push1 CurSymb 'cursymb has higher pressidence
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem 'set cursymb as testsymb
            
            If Not CheckForVal(curWord) Then
                tVal1 = MemStore(curWord)
            Else
                tVal1 = Val(NotReserved(curWord))
            End If
            Push2 tVal1 'stack associated arg
            curWord = curWord + 1
            
        Else 'testsymb has higher or equal pressidence
        If testSymbVal <> 1 Then
            tVal2 = Pop2 'get arg 1
            tVal1 = Pop2 'get arg 2
            MathSymb = Pop1 'get operator
            ret = DoMath(tVal1, tVal2, MathSymb)
            Push2 ret 'stack new processed arg
            
            testSymbVal = GetSymVal(CurSymb)
            testSymb = CurSymb
            Push1 CurSymb 'stack symb
            If Not CheckForVal(curWord) Then
                tVal2 = MemStore(curWord)
            Else
                tVal2 = Val(NotReserved(curWord))
            End If
            Push2 tVal2 'stack associated arg
            curWord = curWord + 1
        Else
            Push1 CurSymb 'cursymb has higher pressidence
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem 'set cursymb as testsymb
            
            If Not CheckForVal(curWord) Then
                tVal1 = MemStore(curWord)
            Else
                tVal1 = Val(NotReserved(curWord))
            End If
            Push2 tVal1 'stack associated arg
            curWord = curWord + 1
        End If
        
        End If
        
    Next
    
    Do Until StackTop1 = 0
        tVal2 = Pop2
        tVal1 = Pop2
        tOps = Pop1
        ret = DoMath(tVal1, tVal2, tOps)
        Push2 ret
        If StackTop1 = 0 Then
        y = StackTop1
        End If
    Loop
    ret = Pop2
End If
ClearStacks
Arithmetic = ret

End Function

Private Function GetSymVal(lStr As String) As Integer
Dim ret As Integer
    If lStr = "+" Then ret = 4
    If lStr = "-" Then ret = 4
    If lStr = "*" Then ret = 3
    If lStr = "/" Then ret = 3
    If lStr = "^" Then ret = 2
    If lStr = "(" Then ret = 1
    
    GetSymVal = ret
    
End Function

Private Function DoMath(lInt1 As Integer, lInt2 As Integer, lStr As String) As Integer
Dim ret As Integer
    Select Case lStr
        Case "+"
            ret = lInt1 + lInt2
        Case "-"
            ret = lInt1 - lInt2
        Case "*"
            ret = lInt1 * lInt2
        Case "/"
            If lInt2 <> 0 Then
                ret = lInt1 / lInt2
            End If
        Case "^"
            ret = lInt1 ^ lInt2
    End Select
    
    DoMath = ret
    
End Function

Private Function MemStore(lInt As Integer) As Integer
Dim ret As Integer

    For i = 0 To MemLimit '- 1
        If VarMem(i) = NotReserved(lInt) Then
            ret = ValMem(i)
            lvFound = True
            itm = i
            Exit For
        End If
    Next
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = NotReserved(lInt)
        ValMem(MemLimit) = 0
        TypMem(MemLimit) = "Integer"
        ret = 0
    End If
    
    MemStore = ret
    
End Function

Private Function CheckForVal(lInt As Integer) As Boolean
Dim TmpV As String
Dim bReturn As Boolean

    TmpV = NotReserved(lInt)
    If Val(TmpV) = 0 Then
        If TmpV = "0" Then
            bReturn = True
        End If
    ElseIf Len(CStr(Val(TmpV))) = Len(TmpV) Then
        bReturn = True
    End If
    
    CheckForVal = bReturn
End Function

Public Function ArithPar(lInt As Integer) As Integer
Dim i As Integer
Dim testSymb As String
Dim testSymbVal As Integer
Dim sVal1 As String
Dim sVal2 As String
Dim tVal1 As Integer
Dim tVal2 As Integer
Dim tOps As String
Dim CurSymb As String
Dim curWord As Integer
Dim CurSymbCount As Integer
Dim CurCount As Integer
Dim NextSymb As String
Dim ret As Integer
Dim MathSymb As String
Dim lst As String
Dim lst1 As String

    lst = PL(lInt).Line
    For i = 1 To Len(lst)
        If Mid(lst, i, 1) = "=" Then
            lst1 = Left(lst, i) + "1*" + Right(lst, Len(lst) - i)
            Parse lst1
        End If
    Next
    
                            'The good stuff
    testSymb = GetSymbol(1)     'Get Operator
    testSymbVal = GetSymVal(testSymb) 'Get Priority
    Push1 testSymb 'the fist op is now on the stack
    
    curWord = 1
    
    If Not CheckForVal(curWord) Then 'Make sure arg is integer
        tVal1 = MemStore(curWord)
    Else
        tVal1 = Val(NotReserved(curWord))
    End If
    Push2 tVal1                 'the 1st arg is stacked
    curWord = curWord + 1
    
    If Not CheckForVal(curWord) Then
        tVal2 = MemStore(curWord)
    Else
        tVal2 = Val(NotReserved(curWord))
    End If
    Push2 tVal2                 'the 2nd arg is stacked
    curWord = curWord + 1
    vv = SymbolCount
    For i = 2 To SymbolCount - 1 '1st sym is already on stack
        CurSymb = GetSymbol(i)   'get next symbol
        If CurSymb = ")" Then
            tOps = Pop1
            Do Until tOps = "(" 'GetSymbol(CurCount) = "("
            DoEvents
                tVal2 = Pop2
                tVal1 = Pop2
                ret = DoMath(tVal1, tVal2, tOps)
                Push2 ret
                tOps = Pop1
                If StopProgram Then Exit Function
            Loop
            If Stack1TopItem <> "(" Then
                If StackTop1 > -1 Then
                    tOps = Pop1
                    tVal2 = Pop2
                    tVal1 = Pop2
                    ret = DoMath(tVal1, tVal2, tOps)
                    testSymbVal = 1
                    Push2 ret
                End If
            End If
        ElseIf CurSymb = "(" Then
            Push1 CurSymb
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem
            
        ElseIf testSymbVal > GetSymVal(CurSymb) Then
            Push1 CurSymb 'cursymb has higher pressidence
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem 'set cursymb as testsymb
            
            If Not CheckForVal(curWord) Then
                tVal1 = MemStore(curWord)
            Else
                tVal1 = Val(NotReserved(curWord))
            End If
            Push2 tVal1 'stack associated arg
            curWord = curWord + 1
            
        Else 'testsymb has higher or equal pressidence
        If testSymbVal <> 1 Then
            tVal2 = Pop2 'get arg 1
            tVal1 = Pop2 'get arg 2
            MathSymb = Pop1 'get operator
            ret = DoMath(tVal1, tVal2, MathSymb)
            Push2 ret 'stack new processed arg
            
            testSymbVal = GetSymVal(CurSymb)
            testSymb = CurSymb
            Push1 CurSymb 'stack symb
            If Not CheckForVal(curWord) Then
                tVal2 = MemStore(curWord)
            Else
                tVal2 = Val(NotReserved(curWord))
            End If
            Push2 tVal2 'stack associated arg
            curWord = curWord + 1
        Else
            Push1 CurSymb 'cursymb has higher pressidence
            testSymbVal = GetSymVal(Stack1TopItem)
            testSymb = Stack1TopItem 'set cursymb as testsymb
            
            If Not CheckForVal(curWord) Then
                tVal1 = MemStore(curWord)
            Else
                tVal1 = Val(NotReserved(curWord))
            End If
            Push2 tVal1 'stack associated arg
            curWord = curWord + 1
        End If
        
        End If
        
    Next
    
    Do Until StackTop1 = 0
        tVal2 = Pop2
        tVal1 = Pop2
        tOps = Pop1
        ret = DoMath(tVal1, tVal2, tOps)
        Push2 ret
        If StackTop1 = 0 Then
        y = StackTop1
        End If
    Loop
    ret = Pop2

ClearStacks
ArithPar = ret
End Function
