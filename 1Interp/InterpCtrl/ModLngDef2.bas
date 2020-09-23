Attribute VB_Name = "ModLngDef"
'Could use if desired
Public Const NUMBEROFKEYS = 36
Public Const SYMBEQUAL = NUMBEROFKEYS + 1
Public Const SYMBPLUS = NUMBEROFKEYS + 2
Public Const SYMBMINUS = NUMBEROFKEYS + 3
Public Const SYMBDIVIDE = NUMBEROFKEYS + 4
Public Const SYMBSTAR = NUMBEROFKEYS + 5
Public Const SYMBRPAREN = NUMBEROFKEYS + 6
Public Const SYMBLPAREN = NUMBEROFKEYS + 7
Public Const SYMBLTHAN = NUMBEROFKEYS + 8
Public Const SYMBGTHAN = NUMBEROFKEYS + 9
Public Const SYMBQUOTE = NUMBEROFKEYS + 10
Public Const SYMBDOLLAR = NUMBEROFKEYS + 11
Public Const SYMBCARROT = NUMBEROFKEYS + 12
Public Const SYMBCOMMA = NUMBEROFKEYS + 13
Private WordsDefined

Public Sub Define(Optional lStr As Variant)
'This sets the language

If IsMissing(lStr) Then
    If Not WordsDefined Then Default
Else
    AlternateLang1
    WordsDefined = 1
End If

End Sub

Private Sub Default()
'This is the sub that defines the legal keywords
'   and symbols the language will use.
'   must be all upper case.
'   vvvv - add you keywords here - vvvv

SetKeyWord "FOR" '1 indicates the ordinal
SetKeyWord "TO" '2
SetKeyWord "NEXT" '3
SetKeyWord "IF" '4
SetKeyWord "THEN" '5
SetKeyWord "END" '6
SetKeyWord "PRINT" '7
SetKeyWord "INKEY" '8
SetKeyWord "INPUT" '9
SetKeyWord "DATE" '10
SetKeyWord "TIME" '11
SetKeyWord "RANDOM" '12
SetKeyWord "CONCAT" '13
SetKeyWord "INSTRING (Future Implementation)" '14
SetKeyWord "SPACE" '15
SetKeyWord "CLS" '16
SetKeyWord "SCREEN (Future Implementation)" '17
SetKeyWord "PLOT (Future Implementation)" '18
SetKeyWord "PRINTTO (Future Implementation)" '19
SetKeyWord "DO" '20
SetKeyWord "WHILE" '21
SetKeyWord "LOOP" '22
SetKeyWord "UNTIL" '23
SetKeyWord "DIR" '24 add to it
SetKeyWord "DISPLAYFILE" '25 don't forget to add UDT
SetKeyWord "CHANGEDIR" '26
SetKeyWord "PATH" '27
SetKeyWord "SUB" '28
SetKeyWord "OPEN" '29
SetKeyWord "CLOSE" '30
SetKeyWord "SHELL" '31
SetKeyWord "SENDKEYS" '32
SetKeyWord "More Keywords Will Be Added" '33
SetKeyWord "XXXXXXXX" '34
SetKeyWord "XXXXXXXX" '35
SetKeyWord "XXXXXXXX" '36

SetSymbol "=" '37 if you go beyond 36 keywords
SetSymbol "+" '38 you have to adjust the UDT const's
SetSymbol "-" '39
SetSymbol "/" '40
SetSymbol "*" '41
SetSymbol "(" '42
SetSymbol ")" '43
SetSymbol "<" '44
SetSymbol ">" '45
SetSymbol Chr(34) '46
SetSymbol "$" '47
SetSymbol "^" '48
SetSymbol "," '49
End Sub

Private Sub AlternateLang1()
'This is the sub that defines the legal keywords
'   and symbols the language will use.
'
SetKeyWord "XXXXXX" '1
SetKeyWord "XXXXXX" '2
SetKeyWord "XXXXXX" '3
SetKeyWord "XXXXXX" '4
SetKeyWord "XXXXXX" '5
SetKeyWord "XXXXXX" '6
SetKeyWord "XXXXXX" '7
SetKeyWord "XXXXXX" '8
SetKeyWord "XXXXXX" '9
SetKeyWord "XXXXXX" '10
SetKeyWord "XXXXXX" '11
SetKeyWord "XXXXXX" '12
SetKeyWord "XXXXXX" '13
SetKeyWord "XXXXXX" '14
SetKeyWord "XXXXXX" '15
SetKeyWord "XXXXXX" '16

SetSymbol "@" '17
SetSymbol "@" '18
SetSymbol "@" '19
SetSymbol "@" '20
SetSymbol "@" '21
SetSymbol "@" '22
SetSymbol "@" '23
SetSymbol "@" '24
SetSymbol "@" '25
SetSymbol "@" '26
SetSymbol "@" '27

End Sub

