Attribute VB_Name = "ModLanguageINPUT"
'This module requires the following code to be
'   entered into the Code Windows Key(Down|Up) Event
'
'   If bInInput=True Then
'       If KeyCode = vbKeyReturn Then
'           InputBuffer = Buffer
'           Buffer = ""
'           bInInput=False
'           CompleteInput
'       Else
'           Buffer = Buffer + CStr(KeyCode)
'       End If
'   End If
'
Public Buffer As String
Public InputBuffer As String
Public bInInput As Boolean
Public InputType As Integer
Public bKeyInput As Boolean
Public KeyBuffer As String

Public Function InputVal(tbox As TextBox) As Integer

    tbox.Text = tbox.Text + "?: "
    lStart = Len(tbox.Text)
    tbox.SelStart = lStart
    tbox.SetFocus
    bInInput = True
    InputType = 1
    Do
        DoEvents
        If StopProgram Then Exit Function
    Loop Until bInInput = False
    
End Function

Public Function CompleteInput()
Dim lVal As Integer
Dim lvFound As Boolean
Dim lLoc As Integer
Dim lStart As Integer
Dim lStr As String

    For i = 0 To MemLimit '- 1
        If InputType = 1 Then
            If VarMem(i) = AllItems(1) Then
                lLoc = i
                lvFound = True
                Exit For
            End If
        ElseIf InputType = 2 Then
            If VarMem(i) = AllItems(2) Then
                lLoc = i
                lvFound = True
                Exit For
            End If
        End If
    Next
    
    If Not lvFound Then
        lLoc = MemLimit
    End If
    
    'lVal = InputBox("Value", "Numeric Value Input", "0")
    If InputType = 1 Then
        lVal = Val(InputBuffer)
    ElseIf InputType = 2 Then
        lStr = InputBuffer
    End If
    
    If Not lvFound Then
        MemLimit = MemLimit + 1
        
        If InputType = 1 Then
            VarMem(MemLimit) = AllItems(1)
            ValMem(MemLimit) = lVal
            TypMem(MemLimit) = "Integer"
        ElseIf InputType = 2 Then
            VarMem(MemLimit) = AllItems(2)
            ValMem(MemLimit) = lStr
            TypMem(MemLimit) = "String"
        End If
    Else
        If InputType = 1 Then
            ValMem(lLoc) = lVal
        ElseIf InputType = 2 Then
            ValMem(lLoc) = lStr
        End If
    End If

End Function

Public Function InputString(tbox As TextBox) As String

    tbox.Text = tbox.Text + "?: "
    lStart = Len(tbox.Text)
    tbox.SelStart = lStart
    tbox.SetFocus
    bInInput = True
    InputType = 2
    Do
        DoEvents
        If StopProgram Then Exit Function
    Loop Until bInInput = False
    
End Function

Public Sub GetStringFromKey(tbox As TextBox)
    
    lStart = Len(tbox.Text)
    tbox.SelStart = lStart
    tbox.SetFocus
    bKeyInput = True
    InputType = 2
    Do
        DoEvents
        If StopProgram Then Exit Sub
    Loop Until bKeyInput = False
End Sub

Public Function CompleteKeyInput()
Dim lVal As Integer
Dim lvFound As Boolean
Dim lLoc As Integer
Dim lStart As Integer
Dim lStr As String

    For i = 0 To MemLimit '- 1
        If VarMem(i) = AllItems(0) Then
            lLoc = i
            lvFound = True
            Exit For
        End If
    Next
    
    If Not lvFound Then
        lLoc = MemLimit
    End If
    
    lStr = KeyBuffer
        
    If Not lvFound Then
        MemLimit = MemLimit + 1
        VarMem(MemLimit) = AllItems(0)
        ValMem(MemLimit) = lStr
        TypMem(MemLimit) = "String"
    Else
        ValMem(lLoc) = lStr
    End If

End Function
