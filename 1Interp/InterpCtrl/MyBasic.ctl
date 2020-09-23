VERSION 5.00
Begin VB.UserControl MyBasic 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "MyBasic.ctx":0000
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "MyBasic.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "MyBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function EnterCodeString(lStr As String, txtBox As Object) As String
Attribute EnterCodeString.VB_Description = "Accepts VBKEYRETURN deliminated string"
    On Error GoTo ErrHere
    StopProgram = False
    ErrString = ""
    Set TargetBox = txtBox
    ParseCode lStr
    EnterCodeString = ErrString
    Exit Function
ErrHere:
    EnterCodeString = "Err #100   ERROR IN ENTERCODESTRING (BSSOKBasic)"
End Function

Public Function WaitForInput(lKey As Integer) As String
Attribute WaitForInput.VB_Description = "Use In Textbox Key Event for User Input"
' bInInput, InputBuffer, Buffer, and bKeyInput
' are globals defined in ModLanguageInput
On Error GoTo ErrHere
ErrString = ""
    If bInInput = True Then
       If lKey = vbKeyReturn Then
           InputBuffer = Buffer
           Buffer = ""
           bInInput = False
           CompleteInput
       Else
           Buffer = Buffer + Chr(lKey)
       End If
    End If
    If bKeyInput = True Then
        DoEvents
        KeyBuffer = Chr(lKey)
        bKeyInput = False
        CompleteKeyInput
    End If
    WaitForInput = ErrString
Exit Function
ErrHere:
    WaitForInput = "Err #101   ERROR IN WAITFORINPUT (BSSOKBasic)"
End Function

Public Function KeyHelp() As String
Dim lRetKey As String
    lRetKey = GetAllKeysAndSymbols
    lRetKey = lRetKey + vbCrLf
    KeyHelp = lRetKey
End Function
Public Function Init() As Boolean
Attribute Init.VB_Description = "Initializes Language"
    Define
End Function

Private Sub UserControl_Resize()
UserControl.Width = Image1.Width
UserControl.Height = Image1.Height
End Sub

Private Sub UserControl_Terminate()
Set TargetBox = Nothing
End Sub
Public Sub TerminateMe()
    StopProgram = True
End Sub
