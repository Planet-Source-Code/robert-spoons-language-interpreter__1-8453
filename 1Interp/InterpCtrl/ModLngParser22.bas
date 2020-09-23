Attribute VB_Name = "ModLngParser2"
Public NotReserved(100) As String
Public KeyWord(100) As String
Public Key(100) As String
Public Symbol(100) As String
Public Symb(100) As String
Public AllItems(100) As String
Public theStructure(100) As Integer
Public m_Symbols As Integer
Public fCount As Integer
Public sCount As Integer
Public tCount As Integer
Public wCount As Integer
Public wordFound As Boolean
Public lastValid As Boolean
Public symbFound As Boolean
Public m_Keys As Integer
Dim myTmplStr1 As String

Public Sub Parse(lStr As String)
'My Baby
'This sub is the heart of the beast.
'   A copy of the string lStr will be serialy
'   decomposed, examined, and re-assyembled into parts.
'   The parts will be catoagorized into one of
'   three classes: Keywords, Symbols, and
'   NotRecognized.
'   Along with the classification of the individual
'   parts, a structure (StructString) will be built
'   that holds all of the parts in the sequence
'   they were built, irregardless of the catagory.
'
'Note: All  lines remarked with *-*-*-*-* are required
'   to catagorize a complete string in quotes as a
'   single item.
'   You can remove all the lines marked with *-*-*-*-*,
'   and the parser will catagorize ALL string elements
'   (i.o.w. words inside quotes will also be
'   catagorized individually - the parser will become
'   a Full Itemizing Parser aka ModLngFullItemParse).
'
Dim lChar As String 'Used to decompose string
Dim lWord As String 'Used to build parts
Dim lReduce As Integer 'Used to indicate reduction scheme
Dim lInAString As Boolean '*-*-*-*-* flipFlop Flag triggered by the quote symbol
On Error GoTo ErrHere
wCount = 0 'Holds the number of words found
tCount = 0 'Hols the number of Keys found
fCount = 0 'Holds the number of all items found
sCount = 0 'Hold the number of symbols found

symbFound = False
wordFound = False
lastValid = False

'..................This is added for TAB operability....
lStr = Trim(lStr)
    myTmplStr1 = lStr
        For i = 1 To Len(myTmplStr1)
            If Left(lStr, 1) = Chr(vbKeyTab) Then
                lStr = Right(lStr, Len(lStr) - 1)
            End If
        Next
'.......................................................
lStr = Trim(lStr) 'Start with a clean string

LS.StructString = "" 'Init the structure
LS.StructLen = 0 'Init the stucture length

If Len(lStr) = 0 Then Exit Sub 'Blank line was entered

Do
    lChar = Left(lStr, 1) 'Dissassymbel string char by char
    For i = 0 To m_Symbols - 1 'Look for a symb
        If Symbol(i) = lChar Then 'Found a symb
            If Not lInAString Then '*-*-*-* FLopFlag designates lChar = chr(34) - a quote symb
                Symb(sCount) = lChar 'Assign symb
                sCount = sCount + 1 'Inc symb count
                symbFound = True 'Set the symb found flag
            
                lReduce = 1 'Designate String Reduction Routine
                If lastValid Then 'Last char completed word
                    NotReserved(wCount) = lWord 'Add word
                    wCount = wCount + 1 'Inc word count
                    AllItems(fCount) = lWord 'Add word
                    theStructure(fCount) = 0 'Assign struct
                    LS.StructString = LS.StructString + "0," 'Assign struct
                    LS.StructLen = LS.StructLen + 1 'Inc struct len
                    fCount = fCount + 1 'Inc items found
                    lWord = "" 'ReInit item stroage
                    lReduce = 3 'Designate String Reduction Routine
                End If
                lastValid = False 'this char is symb
                AllItems(fCount) = lChar 'Assign char
                theStructure(fCount) = i + m_Keys + 1 'Assign struct
                LS.StructString = LS.StructString + CStr(i + m_Keys + 1) + "," 'Assign struct
                LS.StructLen = LS.StructLen + 1 'Inc struct len
                fCount = fCount + 1 'inc item count
            Else '*-*-*-*-*
                lReduce = 4 '*-*-*-*-*
            End If '*-*-*-*-*
            If lChar = Chr(34) Then '*-*-*-*-*
                lInAString = Not lInAString '*-*-*-*-* The filpFlop
                If Not lInAString Then '*-*-*-*
                    Symb(sCount) = lChar '*-*-*-*Assign symb
                    sCount = sCount + 1 '*-*-*-*Inc symb count
                    symbFound = True '*-*-*-*Set the symb found flag
            
                    lReduce = 1 '*-*-*-*Designate String Reduction Routine
                    If lastValid Then '*-*-*-*Last char completed word
                        NotReserved(wCount) = lWord '*-*-*-*Add word
                        wCount = wCount + 1 '*-*-*-*Inc word count
                        AllItems(fCount) = lWord '*-*-*-*Add word
                        theStructure(fCount) = 0 '*-*-*-*Assign struct
                        LS.StructString = LS.StructString + "0," '*-*-*-*Assign struct
                        LS.StructLen = LS.StructLen + 1 '*-*-*-*Inc struct len
                        fCount = fCount + 1 '*-*-*-*Inc items found
                        lWord = "" '*-*-*-*ReInit item stroage
                        lReduce = 3 '*-*-*-*Designate String Reduction Routine
                    End If '*-*-*-*
                    lastValid = False '*-*-*-*this char is symb
                    AllItems(fCount) = lChar '*-*-*-*Assign char
                    theStructure(fCount) = i + m_Keys + 1 '*-*-*-*Assign struct
                    LS.StructString = LS.StructString + CStr(i + m_Keys + 1) + "," '*-*-*-*Assign struct
                    LS.StructLen = LS.StructLen + 1 '*-*-*-*Inc struct len
                    fCount = fCount + 1 '*-*-*-*inc item count
                End If '*-*-*-*
            End If '*-*-*-*-*
            
        End If
    Next
    If Not lInAString Then '*-*-*-*-*
        If Not symbFound Then 'No symb found
            If (lChar = Space(1)) Or (Asc(lChar) = 9) Then 'Current char is a space
                lastValid = False 'a space is not a valid char
                For i = 0 To m_Keys - 1 'Look for a Key
                    If KeyWord(i) = UCase(lWord) Then 'Found a Key
                        lastValid = False 'this char completes a Key
                        Key(tCount) = UCase(lWord) 'Assign Key
                        tCount = tCount + 1 'Inc Key count
                        AllItems(fCount) = Key(tCount - 1) 'Assign item
                        theStructure(fCount) = i + 1
                        LS.StructString = LS.StructString + CStr(i + 1) + ","
                        LS.StructLen = LS.StructLen + 1
                        fCount = fCount + 1
                        lWord = ""
                        wordFound = True 'Set word found flag
                        lReduce = 2
                    End If
                Next
                If wordFound Then 'A Key was found,
                    wordFound = False 'do not treat as non-Key
                Else
                    NotReserved(wCount) = lWord 'Assign a word
                    wCount = wCount + 1 'Inc non-key word count
                    AllItems(fCount) = lWord
                    theStructure(fCount) = 0
                    LS.StructString = LS.StructString + "0,"
                    LS.StructLen = LS.StructLen + 1
                    fCount = fCount + 1
                    lWord = ""
                    lReduce = 1
                End If
            Else
                lastValid = True ' Current char is a char
                lWord = lWord + lChar 'Add it to word storage
            End If
        Else
            symbFound = False 'symb was found, reset flag
        End If
    Else '*-*-*-*-*
        lastValid = True '*-*-*-*-*
        If lChar <> Chr(34) Then '*-*-*-*-*
            lWord = lWord + lChar '*-*-*-*-*
        End If '*-*-*-*-*
        lReduce = 4
    End If '*-*-*-*-*
    
    'Reduce string by appropriate method
    If lReduce = 1 Then
        lReduce = 0
        lStr = Right(lStr, Len(lStr) - 1)
        '..................This is added for TAB operability....
        lStr = Trim(lStr)
        myTmplStr1 = lStr
        For i = 1 To Len(myTmplStr1)
            If Left(lStr, 1) = Chr(vbKeyTab) Then
                lStr = Right(lStr, Len(lStr) - 1)
            End If
        Next
        '.......................................................
        lStr = LTrim(lStr)
    ElseIf lReduce = 2 Then
        lReduce = 0
        lStr = Right(lStr, Len(lStr) - (Len(lWord) + 1))
        '..................This is added for TAB operability....
        lStr = Trim(lStr)
        myTmplStr1 = lStr
        For i = 1 To Len(myTmplStr1)
            If Left(lStr, 1) = Chr(vbKeyTab) Then
                lStr = Right(lStr, Len(lStr) - 1)
            End If
        Next
        '.......................................................
        lStr = LTrim(lStr)
    ElseIf lReduce = 3 Then
        lReduce = 0
        lStr = Right(lStr, Len(lStr) - (Len(lWord) + 1))
        '..................This is added for TAB operability....
        lStr = Trim(lStr)
        myTmplStr1 = lStr
        For i = 1 To Len(myTmplStr1)
            If Left(lStr, 1) = Chr(vbKeyTab) Then
                lStr = Right(lStr, Len(lStr) - 1)
            End If
        Next
        '.......................................................
        lStr = LTrim(lStr)
    Else
        lStr = Right(lStr, Len(lStr) - 1)
    End If
    
Loop Until Len(lStr) < 1 'Keep going untill done

'Clean up for last word
For i = 0 To m_Keys - 1
    If KeyWord(i) = UCase(lWord) Then
        lastValid = False
        Key(tCount) = UCase(lWord)
        tCount = tCount + 1
        AllItems(fCount) = Key(tCount - 1)
        theStructure(fCount) = i + 1
        LS.StructString = LS.StructString + CStr(i + 1) + ","
        LS.StructLen = LS.StructLen + 1
        fCount = fCount + 1
        lWord = ""
        wordFound = True
        lReduce = 2
    End If
Next

If Not wordFound Then
    NotReserved(wCount) = lWord
    AllItems(fCount) = lWord
    theStructure(fCount) = 0
    LS.StructString = LS.StructString + "0,"
    LS.StructLen = LS.StructLen + 1
    x = 4
End If
Exit Sub
ErrHere:
ErrString = "Err #300   ERROR IN PARSE (ModLngParser2)"
End Sub

Public Function NumNotReseved() As Integer
On Error GoTo ErrHere
    NumNotReseved = wCount
Exit Function
ErrHere:
ErrString = "Err #301   ERROR IN NUMNOTRESERVED (ModLngParser2)"
End Function

Public Function Tokens() As Integer
On Error GoTo ErrHere
    Tokens = tCount
Exit Function
ErrHere:
ErrString = "Err #302   ERROR IN TOKENS (ModLngParser2)"
End Function

Public Function Symbols() As Integer
On Error GoTo ErrHere
    Symbols = m_Symbols
Exit Function
ErrHere:
ErrString = "Err #303   ERROR IN SYMBOLS (ModLngParser2)"
End Function

Public Function SymbolCount() As Integer
On Error GoTo ErrHere
    SymbolCount = sCount
Exit Function
ErrHere:
ErrString = "Err #304   ERROR IN SYMBOLCOUNT (ModLngParser2)"
End Function

Public Function StructureLen() As Integer
On Error GoTo ErrHere
    StructureLen = fCount
Exit Function
ErrHere:
ErrString = "Err #305   ERROR IN STRUCTURELEN (ModLngParser2)"
End Function

Public Function Keys() As Integer
On Error GoTo ErrHere
    Keys = m_Keys
Exit Function
ErrHere:
ErrString = "Err #306   ERROR IN KEYS (ModLngParser2)"
End Function

Public Function GetNotReserved(lInt As Integer) As String
On Error GoTo ErrHere
    GetNotReserved = NotReserved(lInt)
Exit Function
ErrHere:
ErrString = "Err #307   ERROR IN GETNOTRESERVED (ModLngParser2)"
End Function

Public Function GetKeyWord(lInt As Integer) As String
On Error GoTo ErrHere
    GetKeyWord = Key(lInt)
Exit Function
ErrHere:
ErrString = "Err #308   ERROR IN GETKEYWORD (ModLngParser2)"
End Function

Public Function GetSymbol(lInt As Integer) As String
On Error GoTo ErrHere
    GetSymbol = Symb(lInt)
Exit Function
ErrHere:
ErrString = "Err #309   ERROR IN GETSYMBOL (ModLngParser2)"
End Function

Public Function GetAllItems(lInt As Integer) As String
On Error GoTo ErrHere
    GetAllItems = AllItems(lInt)
Exit Function
ErrHere:
ErrString = "Err #310   ERROR IN GETALLITEMS (ModLngParser2)"
End Function
Public Function GetAllKeysAndSymbols() As String
Dim i As Integer
Dim lStr As String
    lStr = "Key Wiords:" + vbCrLf
    For i = 0 To m_Keys - 1
        If KeyWord(i) <> "XXXXXXXX" Then
            lStr = lStr + KeyWord(i) + vbCrLf
        End If
    Next
    lStr = lStr + vbCrLf + "Symbols:" + vbCrLf
    For i = 0 To m_Symbols - 1
        lStr = lStr + Symbol(i) + vbCrLf
    Next
    GetAllKeysAndSymbols = lStr
End Function
Public Sub SetKeyWord(lStr As String)
On Error GoTo ErrHere
    KeyWord(m_Keys) = lStr
    m_Keys = m_Keys + 1
Exit Sub
ErrHere:
ErrString = "Err #311   ERROR IN SETKEYWORD (ModLngParser2)"
End Sub

Public Sub SetSymbol(lStr As String)
On Error GoTo ErrHere
    Symbol(m_Symbols) = lStr
    m_Symbols = m_Symbols + 1
Exit Sub
ErrHere:
ErrString = "Err #312   ERROR IN SETSYMBOL (ModLngParser2)"
End Sub

Public Function GetTheStructure(lInt) As Integer
On Error GoTo ErrHere
    GetTheStructure = theStructure(lInt)
Exit Function
ErrHere:
ErrString = "Err #313   ERROR IN GETTHESTRUCTURE (ModLngParser2)"
End Function



