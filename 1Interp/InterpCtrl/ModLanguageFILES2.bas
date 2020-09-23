Attribute VB_Name = "ModLanguageFILES"
Public Type DirReturn
    DirString As String
    DirCount As Currency
    FileCount As Currency
    TotalFileLen As Currency
End Type
Public gPath As String

Public Function Directories(lInt As Integer) As DirReturn
Dim lStr As String
Dim lret As String
Dim i As Integer
Dim j As Integer
Dim MyName As String
Dim lDirRet As DirReturn
Dim aFile As Boolean
Dim lc9 As Integer

On Error GoTo ErrHere

lStr = AllItems(2)
For i = 1 To Len(lStr)
    If Mid(lStr, i, 1) = "\" Then j = i
    If Mid(lStr, i, 1) = "." Then
        aFile = True
        Exit For
    End If
Next
lret = lStr + vbCrLf
If aFile Then
    lDirRet.DirCount = 0
    MyPath = lStr
    Path = MyPath
    MyName = Dir(MyPath, vbHidden)
    Do While MyName <> ""
    DoEvents
        If MyName <> "." And MyName <> ".." Then
            lDirRet.FileCount = lDirRet.FileCount + 1
            lret = lret + vbCrLf + MyName
        End If
        MyName = Dir
        If StopProgram Then Exit Function
    Loop
    lDirRet.DirString = lret
Else
    MyPath = Left(lStr, j)
    Path = MyPath
    MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While MyName <> ""
    DoEvents
    ' Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                lDirRet.DirCount = lDirRet.DirCount + 1
                lret = lret + vbCrLf + "<" + MyName + ">"
            End If
        End If
        MyName = Dir
        If StopProgram Then Exit Function
    Loop

    MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While MyName <> ""
    DoEvents
    ' Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
                If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
                    lDirRet.FileCount = lDirRet.FileCount + 1
                    lret = lret + vbCrLf + MyName + "     " + CStr(FileLen(MyPath + MyName))
                End If
         End If
        MyName = Dir
        If StopProgram Then Exit Function
    Loop
    lDirRet.DirString = lret
End If
Directories = lDirRet
ErrHere:

End Function

Public Function DisplayAFile(plint) As String
Dim lStr As String
Dim testlStr As String
Dim FileNum As Integer
Dim lLine As String

FileNum = FreeFile
lStr = AllItems(2)
testlStr = Dir(lStr)
If testlStr <> "" Then
    If testlStr <> "." And testlStr <> ".." Then
        If FileLen(lStr) < 9000 Then
            Open lStr For Input As #FileNum ' Open file.
                'axxx = EOF(FreeFile)
                v = 5
                Do While Not EOF(FileNum) 'FreeFile) ' Loop until end of file.
                    Line Input #FileNum, lLine ' Read line into variable.
                    lStr = lStr + lLine + vbCrLf ' Print to Debug window.
                Loop
            Close #FileNum   ' Close file.
        Else
            lStr = "File is too large to open" + vbCrLf + _
            "FILESIZE = " + CStr(FileLen(lStr)) + " BYTES."
        End If
    End If
End If
DisplayAFile = lStr

End Function

Public Function DisplayBFile(plint) As String
Dim lStr As String
Dim testlStr As String
Dim FileNum As Integer
Dim lLine As String
Dim lStrFound As Boolean

FileNum = FreeFile
lStr = AllItems(1)
For i = 0 To MemLimit
    If VarMem(i) = lStr Then
        If TypMem(i) = "String" Then
            lStr = LTrim(ValMem(i))
            lStrFound = True
            Exit For
        End If
    End If
Next
If Not lStrFound Then
    lStr = gPath + lStr
End If

testlStr = Dir(lStr)
If testlStr <> "" Then
    If testlStr <> "." And testlStr <> ".." Then
        If FileLen(lStr) < 9000 Then
            Open lStr For Input As #FileNum ' Open file.
                'axxx = EOF(FreeFile)
                v = 5
                Do While Not EOF(FileNum) 'FreeFile) ' Loop until end of file.
                    Line Input #FileNum, lLine ' Read line into variable.
                    lStr = lStr + lLine + vbCrLf ' Print to Debug window.
                Loop
            Close #FileNum   ' Close file.
        Else
            lStr = "File is too large to open" + vbCrLf + _
            "FILESIZE = " + CStr(FileLen(lStr)) + " BYTES."
        End If
    End If
End If
DisplayBFile = lStr

End Function

Public Function ReturnPath() As DirReturn
Dim lStr As String
Dim lret As String
Dim i As Integer
Dim j As Integer
Dim MyName As String
Dim aFile As Boolean
Dim lDirRet As DirReturn
Dim lTmpLong As Currency
Dim lc9 As Currency
Dim strTmp As String

lDirRet.TotalFileLen = 0
masKstr = String(20, "_")
On Error GoTo ErrHere
If gPath = "" Then
    gPath = "C:\"
End If
If Right(gPath, 1) <> "\" Then gPath = gPath + "\"


MyPath = Left(gPath, j)
Path = MyPath
MyName = Dir(gPath, vbDirectory)   ' Retrieve the first entry.
Do While MyName <> ""
DoEvents
' Ignore the current directory and the encompassing directory.
    If MyName <> "." And MyName <> ".." Then
    ' Use bitwise comparison to make sure MyName is a directory.
        If (GetAttr(gPath & MyName) And vbDirectory) = vbDirectory Then
            lDirRet.DirCount = lDirRet.DirCount + 1
            lret = lret + vbCrLf + "<" + MyName + ">"
        End If
    End If
    MyName = Dir
    If StopProgram Then Exit Function
Loop

MyName = Dir(gPath, vbDirectory)   ' Retrieve the first entry.
Do While MyName <> ""
DoEvents
' Ignore the current directory and the encompassing directory.
    
    If MyName <> "." And MyName <> ".." Then
    ' Use bitwise comparison to make sure MyName is a directory.
        If (GetAttr(gPath & MyName) And vbDirectory) <> vbDirectory Then
                lDirRet.FileCount = lDirRet.FileCount + 1
                tmpl1 = Len(MyName)
                lTmpLong = lTmpLong + FileLen(gPath + MyName)
                strTmp = Format(MyName, "!>&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
                ltmpl2 = CStr(FileLen(gPath + MyName))
                ltmpl2 = Format(ltmpl2, "###,###,###,###,###")
                If tmpl1 < 40 Then
                    strTmp = strTmp + Mid(masKstr, 1, (40 - tmpl1)) + ltmpl2
                Else
                    strTmp = strTmp + ltmpl2
                End If
                lret = lret + vbCrLf + strTmp
        End If
    End If
    MyName = Dir
    If StopProgram Then Exit Function
Loop
lDirRet.TotalFileLen = lTmpLong
lDirRet.DirString = lret
ReturnPath = lDirRet
Exit Function
ErrHere:
End Function

Public Function ChangePath(lInt As Integer)
Dim lStr As String
Dim testlStr As String
Dim i As Integer
Dim j As Integer
Dim aFile As Integer
Dim IsARoot As Boolean

lStr = AllItems(2)
If Len(lStr) = 3 Then
    If Right(lStr, 2) = ":\" Then
        IsARoot = True
    End If
End If
If Not IsARoot Then
    For i = 1 To Len(lStr)
        If Mid(lStr, i, 1) = "\" Then j = i
        If Mid(lStr, i, 1) = "." Then
            aFile = True
            Exit For
        End If
    Next
    If aFile Then
        lStr = Left(lStr, j)
    End If
    testlStr = Dir(lStr, vbDirectory)
    If testlStr <> "" Then
        gPath = lStr
    End If
Else
    testlStr = Dir(lStr, vbDirectory)
    If testlStr <> "" Then
        gPath = lStr
    End If
End If

End Function

Public Function ShowPath() As String
    If gPath = "" Then
        gPath = "C:\"
    End If
    If Right(gPath, 1) <> "\" Then gPath = gPath + "\"
    ShowPath = gPath
End Function
