VERSION 5.00
Begin VB.Form frmOpenAndSave 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   3675
   ClientTop       =   2640
   ClientWidth     =   6330
   Icon            =   "frmOpenAndSave.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6330
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Basic1"
      Top             =   3360
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmOpenAndSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lStrLine(100)
Dim lStrLineNum As Integer
Dim lFile As String

Private Sub Command1_Click()
Dim FileNumber As Long
Dim TextLine
Dim lStr As String
Dim theNumber As Integer
Dim tLen As Integer
Dim PathFile As String

On Error GoTo ErrHandler

If lFile = "" Then
    lPath = File1.Path
    If Right(lPath, 1) <> "\" Then
        lPath = lPath + "\"
    End If
    lFile = lPath + Text1.Text
    Label1 = lFile
End If

If OpenSave = 0 Then
    FileNumber = FreeFile
    Open lFile For Input As #FileNumber ' Open file.
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, TextLine ' Read line into variable.
        lStr = lStr + TextLine + vbCrLf ' Print to Debug window.
    Loop
    Close FileNumber
    frmMYBASIC3.Text1.Text = lStr 'TextObject.Text = lStr
Else

    tLen = Len(frmMYBASIC3.Text1.Text) 'TextObject.Text)
    If tLen < 1 Then Exit Sub
    lStr = frmMYBASIC3.Text1.Text 'TextObject.Text
    
    j = 1
    For i = 1 To tLen - 1
        If Not i < j Then
        If Mid(lStr, i, 2) = vbCrLf Then
            lStrLine(lStrLineNum) = Mid(lStr, j, i - j)
            lStrLineNum = lStrLineNum + 1
            j = i + 2
            a = 2
        End If
        End If
    Next

    FileNumber = FreeFile
    PathFile = Dir1.Path
    If Right(PathFile, 1) <> "\" Then
        PathFile = PathFile + "\"
    End If
    
    FullFile = Text1
    If Mid(FullFile, 2, 1) <> ":" Then
        FullFile = PathFile + FullFile
    End If
    If Mid(FullFile, Len(FullFile) - 3, 1) <> "." Then
        FullFile = PathFile + FullFile + ".RSB"
    End If
    
    Open FullFile For Output As #FileNumber   ' Open file for output.
    Do
        Print #FileNumber, lStrLine(theNumber)  ' Print text to file.
        theNumber = theNumber + 1
    Loop Until theNumber = lStrLineNum
    Close FileNumber
    
End If
frmOpenAndSave.Hide

Exit Sub
ErrHandler:
 FullFile = "Basic1"
Resume Next

End Sub

Private Sub Command2_Click()
frmOpenAndSave.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim lPath As String

lPath = File1.Path
If Right(lPath, 1) <> "\" Then
    lPath = lPath + "\"
End If
lFile = lPath + File1.FileName
Label1 = lFile

End Sub

Private Sub Form_Load()
Drive1.Drive = "C:"
File1.Pattern = "*.RSB;*.txt"
End Sub
