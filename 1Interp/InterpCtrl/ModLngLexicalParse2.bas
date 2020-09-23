Attribute VB_Name = "ModLngLexicalParse"
Dim IsBlock As Boolean
Dim bkCount As Integer

Public Function Valid(lInt As Integer) As Boolean
'This is the lexical parser
'This is all there is to it
On Error GoTo ErrHere
Dim Tmp As Integer
Dim bReturn As Boolean
Dim ilLen As Integer

For i = 1 To UBound(LF)
    ilLen = Len(LF(i))
    If LF(i) = Left(PL(lInt).StructString, ilLen) Then
        PL(lInt).LineType = i
        bReturn = True
        Exit For 'XXXXXXXXXXXXXXXXXXXXX
    End If
Next

Valid = bReturn
Exit Function
ErrHere:
ErrString = "Err #200   ERROR IN VALID (ModLngLexicalParse)"
End Function
