Attribute VB_Name = "ModLanguageSHELL"
Dim Program As String
Dim ReturnValue() As Long
Dim Programs As Integer

Public Sub UseAppL()
Dim i As Integer

Program = AllItems(2)
Programs = Programs + 1
ReDim Preserve ReturnValue(Programs)
ReturnValue(Programs) = Shell(Program, 1)  ' Run Calculator.
AppActivate ReturnValue(Programs)     ' Activate the Calculator.
If AllItems(4) <> "" Then
    SendKeys AllItems(4), True
End If
'SendKeys "%{F4}", True  ' Send ALT+F4 to close Calculator.

End Sub

Public Sub UseAppV()
'Dim ReturnValue As Long
Dim i As Integer
Dim Program As String
Dim lFound As Boolean

Programs = Programs + 1
ReDim Preserve ReturnValue(Programs)

On Error GoTo ErrHere

For i = 0 To MemLimit
    If VarMem(i) = AllItems(1) Then
        If TypMem(i) = "String" Then
            Program = ValMem(i)
            lFound = True
            Exit For
        End If
    End If
Next

If Not lFound Then GoTo ErrHere
        
ReturnValue(Programs) = Shell(Program, 1)  ' Run Program.
AppActivate ReturnValue(Programs)     ' Activate the Calculator.

Exit Sub
ErrHere:

End Sub

Public Sub UseAppKeys()
On Error GoTo ErrHere

SendKeys AllItems(2)
    
ErrHere:
End Sub

Public Sub UseAppKeysLate()
Dim curProg As Integer

On Error GoTo ErrHere

For i = 0 To MemLimit
    If VarMem(i) = AllItems(1) Then
        If TypMem(i) = "Integer" Then
            curProg = ValMem(i)
            lFound = True
            Exit For
        End If
    End If
Next

If Not lFound Then GoTo ErrHere

AppActivate ReturnValue(curProg)
SendKeys AllItems(4)
    
ErrHere:
End Sub
