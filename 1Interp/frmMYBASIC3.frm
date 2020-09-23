VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "*\AInterpCtrl\MyBasicInterpreter.vbp"
Begin VB.Form frmMYBASIC3 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9585
   Icon            =   "frmMYBASIC3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin MyBasicInterpeter.MyBasic MyBasic1 
      Left            =   2880
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin ComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   4980
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "STOPPED"
            TextSave        =   "STOPPED"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   11271
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:09 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMYBASIC3.frx":27A2
      Left            =   3480
      List            =   "frmMYBASIC3.frx":27A4
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "KeyWords"
      Top             =   0
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0FF&
      Height          =   1620
      ItemData        =   "frmMYBASIC3.frx":27A6
      Left            =   0
      List            =   "frmMYBASIC3.frx":27C2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"frmMYBASIC3.frx":2811
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSepF2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut    (Use The Mouse or <CTRL + X>)"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy  (Use The Mouse or <CTRL + C>)"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste (Use The Mouse or <CTRL + V>)"
      End
      Begin VB.Menu mnuSepE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearCode 
         Caption         =   "ClearCo&de Window"
      End
      Begin VB.Menu mnuClearProg 
         Caption         =   "C&learProgram Window"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOpCode 
         Caption         =   "&Code Window"
         Begin VB.Menu mnuProgWin 
            Caption         =   "&Window Color"
         End
         Begin VB.Menu mnuProgColor 
            Caption         =   "&Text Color"
         End
         Begin VB.Menu mnuCodeFont 
            Caption         =   "&Font Size"
            Begin VB.Menu mnuC8 
               Caption         =   "8"
            End
            Begin VB.Menu mnuC10 
               Caption         =   "10"
            End
            Begin VB.Menu mnuC12 
               Caption         =   "12"
            End
            Begin VB.Menu mnuC14 
               Caption         =   "14"
            End
         End
         Begin VB.Menu mnuCBold 
            Caption         =   "&BOLD"
         End
      End
      Begin VB.Menu mnuOpProg 
         Caption         =   "&Program Window"
         Begin VB.Menu mnuCodeWin 
            Caption         =   "&Window Color"
         End
         Begin VB.Menu mnuCodeColor 
            Caption         =   "&Text Color"
         End
         Begin VB.Menu mnuProgFont 
            Caption         =   "&Font Size"
            Begin VB.Menu mnuP8 
               Caption         =   "8"
            End
            Begin VB.Menu mnuP10 
               Caption         =   "10"
            End
            Begin VB.Menu mnuP12 
               Caption         =   "12"
            End
            Begin VB.Menu mnuP14 
               Caption         =   "14"
            End
         End
         Begin VB.Menu mnuPBOLD 
            Caption         =   "&BOLD"
         End
      End
      Begin VB.Menu mnuKILL 
         Caption         =   "Show Example"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSingle 
         Caption         =   "&Single"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "&TileVerticle"
      End
      Begin VB.Menu mnuTileHoriz 
         Caption         =   "T&ile Horizontal"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&RUN"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "&STOP"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuCode 
      Caption         =   "&Code"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Program"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMYBASIC3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim View As Integer
Dim sWidth As Single
Dim hWidth As Single
Dim vWidth As Single
Dim sHeight As Single
Dim hHeight As Single
Dim vHeight As Single
Dim RunCount As Integer
Dim WindowPane As String
Dim sCap As String
Dim sProgram As String
Dim gFocus As Integer
Const WCODE = "CODE WINDOW "
Const WPROG = "PROGRAM WINDOW "
Dim HelpItems As Integer
Dim ProgRunning As Boolean
Dim KeepRunning As Boolean
Dim Colors(10) As Long
Dim ColorSelect As Integer
Dim Code1 As Long
Dim Code2 As Long
Dim Code3 As Integer
Dim Code4 As Boolean
Dim Prog1 As Long
Dim Prog2 As Long
Dim Prog3 As Integer
Dim Prog4 As Boolean

Private Sub Combo1_Click()
Dim i As Integer
    i = Combo1.ListIndex
    If i = 0 Then
        Text3 = MyBasic1.KeyHelp
    Else
        Text3 = ShowHelpMe(i)
    End If
End Sub

Private Sub Combo1_GotFocus()
gFocus = 1
End Sub

Private Sub Form_Load()
Dim lStr As String
Dim i As Integer

Colors(0) = &H0&      'Black
Colors(1) = &HFFFFFF  'White
Colors(2) = &HC0E0FF  'Light Brown
Colors(3) = &H4080&   'Dark Brown
Colors(4) = &HC0FFFF  'Light Yelow
Colors(5) = &H8080&   'Dark Yellow
Colors(6) = &HC000&   'Light Green
Colors(7) = &H8000&   'Dark Green
Colors(8) = &HFFFF80  'Light Blue
Colors(9) = &HC00000  'Dark Blue

Code1 = GetSetting("BSSOKBASIC", "Startup", "CODECOLOR", &HFFFFFF)
Code2 = GetSetting("BSSOKBASIC", "Startup", "CODEWIN", &H0&)
Code3 = GetSetting("BSSOKBASIC", "Startup", "CODESIZE", 10)
Code4 = GetSetting("BSSOKBASIC", "Startup", "CODEBOLD", False)
Prog1 = GetSetting("BSSOKBASIC", "Startup", "PROGCOLOR", &HFFFFFF)
Prog2 = GetSetting("BSSOKBASIC", "Startup", "PROGWIN", &H0&)
Prog3 = GetSetting("BSSOKBASIC", "Startup", "PROGSIZE", 10)

KillExample = GetSetting("BSSOKBASIC", "Startup", "EXAMPLE", False)

Text1.BackColor = Code1
Text1.ForeColor = Code2
Text1.FontSize = Code3
Text1.FontBold = Code4
Text2.BackColor = Prog1
Text2.ForeColor = Prog2
Text2.FontSize = Prog3
Text2.FontBold = Prog4

If Code4 = True Then mnuCBold.Caption = "&NORMAL"
If Prog4 = True Then mnuPBOLD.Caption = "&NORMAL"

HelpItems = InitHelp
sCap = "BSSOK BASIC "
MyBasic1.Init

If Command$ <> "" Then
    If Command$ <> "%1" Then
        WindowPane = WPROG
        Me.Caption = sCap + WindowPane + "RUNNING " + Command$
        Text2.ZOrder
        Me.Show
        RichTextBox1.LoadFile Command$, rtfText
        lStr = RichTextBox1.Text
        Text1.Text = lStr
        lStr = MyBasic1.EnterCodeString(lStr, Text2)
        If lStr <> "" Then
            MsgBox lStr, vbCritical, "ERROR"
        End If
    End If
Else
    WindowPane = WCODE
    Me.Caption = sCap + WindowPane + "WELCOME"
    Text1.ZOrder
End If
Text3.Top = Me.Top + 305
For i = 0 To HelpItems - 1
    Combo1.AddItem HelpName(i)
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lRet As Integer
'If UnloadMode = vbFormControlMenu Then
    If ProgRunning Then
        Cancel = True
        KeepRunning = True
        MsgBox "A Program is still running" + vbCrLf + "Click STOP in the Menu Bar", vbOKOnly, Me.Caption
    Else
        Cancel = False
        KeepRunning = False
    End If
    Unload frmOpenAndSave
End Sub

Private Sub Form_Resize()

If Me.Width > 2400 Then
    If Me.Height < 4000 Then Me.Height = 3000
    If Me.Width < 6645 Then
        Me.Width = 6645
    End If
    sWidth = Me.Width - 100
    sHeight = Me.Height - 1040 '670
    vWidth = sWidth
    vHeight = Int(sHeight / 2)
    hWidth = Int(sWidth / 2)
    hHeight = sHeight

    Select Case View
        Case 0
            Text1.Top = 0
            Text1.Left = 0
            Text2.Top = 0
            Text2.Left = 0
            Text1.Width = sWidth
            Text1.Height = sHeight
            Text2.Width = sWidth
            Text2.Height = sHeight
        Case 1
            Text1.Top = 0
            Text1.Left = 0
            Text2.Top = vHeight
            Text2.Left = 0
            Text1.Width = vWidth
            Text1.Height = vHeight
            Text2.Width = vWidth
            Text2.Height = vHeight
        Case 2
            Text1.Top = 0
            Text1.Left = 0
            Text2.Top = 0
            Text2.Left = hWidth
            Text1.Width = hWidth
            Text1.Height = hHeight
            Text2.Width = hWidth
            Text2.Height = hHeight
    End Select
    Text3.Width = Me.Width - 100
    Text3.Height = Me.Height - 1305 '1005
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If KeepRunning Then
    Cancel = True
    KeepRunning = False
End If
SaveSetting "BSSOKBASIC", "Startup", "CODECOLOR", Code1
SaveSetting "BSSOKBASIC", "Startup", "CODEWIN", Code2
SaveSetting "BSSOKBASIC", "Startup", "CODESIZE", Code3
SaveSetting "BSSOKBASIC", "Startup", "CODEBOLD", Code4
SaveSetting "BSSOKBASIC", "Startup", "PROGCOLOR", Prog1
SaveSetting "BSSOKBASIC", "Startup", "PROGWIN", Prog2
SaveSetting "BSSOKBASIC", "Startup", "PROGSIZE", Prog3
SaveSetting "BSSOKBASIC", "Startup", "PROGBOLD", Prog4
End Sub

Private Sub List1_Click()
Dim i As Integer
    i = List1.ListIndex
    Select Case ColorSelect
        Case 1
            Text1.BackColor = Colors(i)
            Code1 = Colors(i)
        Case 2
            Text1.ForeColor = Colors(i)
            Code2 = Colors(i)
        Case 3
            Text2.BackColor = Colors(i)
            Prog1 = Colors(i)
        Case 4
            Text2.ForeColor = Colors(i)
            Prog2 = Colors(i)
    End Select
    
    List1.ZOrder 1
    
End Sub

Private Sub mnuC10_Click()
Text1.FontSize = 10
Code3 = 10
End Sub

Private Sub mnuC12_Click()
Text1.FontSize = 12
Code3 = 12
End Sub

Private Sub mnuC14_Click()
Text1.FontSize = 14
Code3 = 14
End Sub

Private Sub mnuC8_Click()
Text1.FontSize = 8
Code3 = 8
End Sub

Private Sub mnuCBold_Click()
If mnuCBold.Caption = "&BOLD" Then
    mnuCBold.Caption = "&NORMAL"
    Text1.FontBold = True
    Code4 = True
Else
    mnuCBold.Caption = "&BOLD"
    Text1.FontBold = False
    Code4 = False
End If
End Sub

Private Sub mnuClearCode_Click()
    Text1.Text = ""
    Text1.ZOrder
    Text1.SetFocus
End Sub

Private Sub mnuClearProg_Click()
    Text2.Text = ""
    Text2.ZOrder
    Text2.SetFocus
End Sub

Private Sub mnuCode_Click()
    mnuCode.Enabled = False
    mnuProgram.Enabled = True
    WindowPane = WCODE
    Me.Caption = sCap + WindowPane
    Text1.ZOrder
    Text1.SetFocus
End Sub

Private Sub mnuCodeColor_Click()
ColorSelect = 4
List1.ZOrder
List1.SetFocus
End Sub

Private Sub mnuCodeWin_Click()
ColorSelect = 3
List1.ZOrder
List1.SetFocus
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuHelp_Click()
    Combo1.Visible = True
    Combo1.ZOrder
    Text3.Width = Me.Width - 100
    Text3.Height = Me.Height - 1305 '1005
    Text3.Visible = True
    Text3.ZOrder
    Combo1.SetFocus
End Sub

Private Sub mnuKILL_Click()
mnuCode.Enabled = False
mnuProgram.Enabled = True
Dim i As Integer
    For i = 0 To 2
        Text1.Text = Text1.Text + Example(i)
        Text1.ZOrder
        Text1.SetFocus
    Next
End Sub

Private Sub mnuNew_Click()
    Text2.Text = ""
    Text1.Text = ""
    Text1.ZOrder
    Text1.SetFocus
End Sub

Private Sub mnuOpen_Click()
OpenSave = 0
frmOpenAndSave.Caption = "Open File"
frmOpenAndSave.Show
End Sub

Private Sub mnuP10_Click()
Text2.FontSize = 10
Prog3 = 10
End Sub

Private Sub mnuP12_Click()
Text2.FontSize = 12
Prog3 = 12
End Sub

Private Sub mnuP14_Click()
Text2.FontSize = 14
Prog3 = 8
End Sub

Private Sub mnuP8_Click()
Text2.FontSize = 8
End Sub

Private Sub mnuPBOLD_Click()
If mnuPBOLD.Caption = "&BOLD" Then
    mnuPBOLD.Caption = "&NORMAL"
    Text2.FontBold = True
    Prog4 = True
Else
    mnuPBOLD.Caption = "&BOLD"
    Text2.FontBold = False
    Prog4 = False
End If
End Sub

Private Sub mnuProgColor_Click()
ColorSelect = 2
List1.ZOrder
List1.SetFocus
End Sub

Private Sub mnuProgram_Click()
    mnuProgram.Enabled = False
    mnuCode = True
    WindowPane = WPROG
    Me.Caption = sCap + WindowPane
    Text2.ZOrder
    Text2.SetFocus
End Sub

Private Sub mnuProgWin_Click()
ColorSelect = 1
List1.ZOrder
List1.SetFocus
End Sub

Private Sub mnuRun_Click()
Dim lStr As String
    If ProgRunning Then
        MsgBox "There is already a program running" + vbCrLf + "Use STOP in the Menu Bar to terminate it", vbOKOnly, Me.Caption
        Exit Sub
    End If
    SB1.Panels(1).Text = "RUNNING"
    mnuRun.Enabled = False
    mnuStop.Enabled = True
    mnuCode.Enabled = True
    mnuProgram.Enabled = False
    ProgRunning = True
    WindowPane = WPROG
    Me.Caption = sCap + WindowPane
    Text2.ZOrder
    Text2.SetFocus
    lStr = Me.Text1.Text
    lStr = MyBasic1.EnterCodeString(lStr, Me.Text2)
    If lStr <> "" Then
        MsgBox lStr, vbCritical, "ERROR"
    End If
    ProgRunning = False
    SB1.Panels(1).Text = "STOPPED"
    mnuRun.Enabled = True
    mnuStop.Enabled = False
    RunCount = RunCount + 1
    Text2 = Text2 + "Run # " + CStr(RunCount) + vbCrLf + vbCrLf
    Text2.SelStart = Len(Text2)
End Sub

Private Sub mnuSave_Click()
OpenSave = 1
frmOpenAndSave.Caption = "Save File"
frmOpenAndSave.Show
End Sub

Private Sub mnuSingle_Click()
    View = 0
    Form_Resize
End Sub

Private Sub mnuStop_Click()
    MyBasic1.TerminateMe
End Sub

Private Sub mnuTileHoriz_Click()
    View = 2
    Form_Resize
End Sub

Private Sub mnuTileVert_Click()
    View = 1
    Form_Resize
End Sub

Private Sub Text1_Change()
RunCount = 0
End Sub

Private Sub Text1_Click()
mnuCode.Enabled = False
mnuProgram.Enabled = True
End Sub

Private Sub Text1_GotFocus()
If gFocus = 1 Then
    gFocus = 0
    Text3.Text = ""
    Text3.Visible = False
    Combo1.Visible = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        mnuProgram.Enabled = False
        mnuCode.Enabled = True
        WindowPane = WPROG
        Me.Caption = sCap + WindowPane
        Text2.ZOrder
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_Click()
mnuProgram.Enabled = False
mnuCode.Enabled = True
End Sub

Private Sub Text2_GotFocus()
If gFocus = 1 Then
    gFocus = 0
    Text3.Text = ""
    Text3.Visible = False
    Combo1.Visible = False
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim lStr As String
    If KeyAscii = vbKeyEscape Then
        mnuProgram.Enabled = True
        mnuCode.Enabled = False
        WindowPane = WCODE
        Me.Caption = sCap + WindowPane
        Text1.ZOrder
        Text1.SetFocus
    Else
        lStr = MyBasic1.WaitForInput(KeyAscii)
        If lStr <> "" Then
            MsgBox lStr, vbCritical, "ERROR"
        End If
    End If
End Sub

Private Sub Text3_GotFocus()
gFocus = 1
End Sub


