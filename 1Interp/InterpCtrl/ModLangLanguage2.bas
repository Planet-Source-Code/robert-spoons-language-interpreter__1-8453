Attribute VB_Name = "ModLngLanguage"
'Declaration of Major Variables
Public Lines(100) As String
Public VarMem(100) As String
Public ValMem(100) As Variant 'Integer
'Public StrMem(100) As String
Public TypMem(100) As String
Public LineCode(100) As Integer
Public MemLimit As Integer
Public TheLine As Integer
Public CurLine As Integer
Public bInBlock As Boolean
Public bInBlockCount As Integer
Public SubSide As Integer
Public gWindow As Integer
Public tbox2 As TextBox
Dim thisbox As TextBox
Public MasterBlock As Integer
Public TmpForVar As Integer
Public LoopCountDown As Integer
Type Funcs
    LineNum() As Integer
    FuncLen As Integer
    Name As String
End Type
Dim Func() As Funcs
Dim inFunction As Boolean
Public Functions As Integer

Public Sub ParseCode(lStr As String) '(tbox As TextBox)
'This is the Code Line loop.
'   This function is resposible for the seperation
'   of the code text into program lines.
'   It also hold the main processing loop for the
'   interpretation of each program line
'
Dim lCount As Integer
Dim lLen As Integer
Dim lMore As Boolean
Dim C As String
Dim Code As String
Dim i As Integer
Dim FunctionsLine As Integer
'Dim lstr As String
Dim crFound As Boolean

On Error GoTo ErrHere

Randomize

Static lInitME

For i = 0 To MemLimit
    VarMem(i) = ""
    ValMem(i) = 0
Next
ReDim Func(0)
Dim ReturnValue(0)

MasterBlock = 0
CurLine = 0
MemLimit = 0
bInBlockCount = 0
bInBlock = False

If Not lInitME Then
    ReDim LF(LFNUM + 1)
    LF(0) = "" 'Not Recognised
    LF(1) = formAssign
    LF(2) = formFOR
    LF(3) = formIFTHEN
    LF(4) = formIFTHENELSE
    LF(5) = formIFTHENbin
    LF(6) = formPRINTVAL
    LF(7) = formPRINTSTRING
    LF(8) = formINPUTVAL
    LF(9) = formINPUTSTRING
    LF(10) = formStrAssign
    LF(11) = formDATE
    LF(12) = formTIME
    LF(13) = formRANDOM
    LF(14) = formCONCAT1
    LF(15) = formCONCAT2
    LF(16) = formINSTRING
    LF(17) = formSPACE
    LF(18) = formCLS
    LF(19) = formSCREEN
    LF(20) = formPLOT
    LF(21) = formPRINTTO
    LF(22) = formINKEY
    LF(23) = formDOLOOP
    LF(24) = formDOLOOPS
    LF(25) = formDOLOOPU
    LF(26) = formDOLOOPUS
    LF(27) = formIFTHENS
    LF(28) = formIFTHENNE
    LF(29) = formIFTHENG
    LF(30) = formIFTHENL
    LF(31) = formIFTHENGE
    LF(32) = formIFTHENEG
    LF(33) = formIFTHENLE
    LF(34) = formIFTHENEL
    LF(35) = formIFTHENNES
    LF(36) = formIFTHENGS
    LF(37) = formIFTHENLS
    LF(38) = formIFTHENGES
    LF(39) = formIFTHENEGS
    LF(40) = formIFTHENLES
    LF(41) = formIFTHENELS
    LF(42) = formCHANGEDIR
    LF(43) = formDISPLAYFILE
    LF(44) = formDIRPATH
    LF(45) = formDIR
    LF(46) = formPATH
    LF(47) = formDISPLAYFILEVAR
    LF(48) = formFUNCTION
    LF(49) = formOPEN
    LF(50) = formCLOSE
    LF(51) = formPRINTF
    LF(52) = formINPUTF
    LF(53) = formASSIGNPAR
    LF(54) = formSHELLV
    LF(55) = formSHELLL
    LF(56) = formSENDKEYS
    LF(57) = formSENDKEYSLATE
End If
    
    'If Len(tbox.Text) < 1 Then Exit Function
    If Len(lStr) < 1 Then Exit Sub
    
    lMore = True
    lLen = Len(lStr)
    
    For i = 1 To lLen - 1
        If Mid(lStr, i, 2) = vbCrLf Then
            crFound = True
            Lines(lCount) = Code
            ReDim Preserve PL(lCount)
            PL(lCount).Line = Code
            PL(lCount).LineNumber = lCount
            Code = ""
            lCount = lCount + 1
        Else
            If Not crFound Then
                Code = Code + Mid(lStr, i, 1)
            End If
            crFound = False
        End If
    Next
    
    TheLine = lCount - 1

'Below is the main processing loop.
'   It will iterate through each line of the
'   code text, in sequential order.
'
    For i = 0 To TheLine
        If UCase(Left(PL(i).Line, Len("sub"))) = "SUB" Then
            inFunction = True
                Functions = Functions + 1
            ReDim Preserve Func(Functions)
            Func(Functions).Name = Trim(Right(PL(i).Line, Len(PL(i).Line) - Len("sub")))
        End If
        If inFunction Then
            If UCase(PL(i).Line) = "END SUB" Then
                Func(Functions).FuncLen = FunctionsLine - 1
                FunctionsLine = 0
                inFunction = False
            Else
                FunctionsLine = FunctionsLine + 1
                ReDim Preserve Func(Functions).LineNum(FunctionsLine)
                Func(Functions).LineNum(FunctionsLine) = i + 1
           End If
        Else
            If bInBlock Then
                LoopCountDown = LoopCountDown - 1
            
                If LoopCountDown = 0 Then
            
                    bInBlockCount = 0
                    bInBlock = False
                End If
            End If
        
            If Not bInBlock Then
                Interpret i
            End If
        End If
    Next
    
    TargetBox.Text = TargetBox.Text + _
        vbCrLf + "END PROGRAM" + vbCrLf
    TargetBox.SelStart = Len(TargetBox.Text)
    
    Exit Sub
ErrHere:
    ErrString = "Err #101   ERROR IN PARSECODE (ModLanLanguage)"
End Sub

Public Sub Interpret(ByVal lInt As Integer) 'As Boolean
'This is the Code Execution controller.
'   First, it sends the code line lInt to
'       ModLangParser's Parser function
'       (the general parser).
'   Second, it sends the line to ModLaexicalParse's
'       Valid function (the lexical parser) for
'       verifaction of the lines validity.
'   Third, If the line is a valid line as defined
'       in the language defs, it sends the line to
'       DoIT to be executed.
'
On Error GoTo ErrHere
Dim lStr As String
    
    lStr = PL(lInt).Line
    Parse lStr
    PL(lInt).StructString = LS.StructString
    PL(lInt).StructLen = LS.StructLen
    Response = Valid(lInt) 'Method of LEX
    
        Set thisbox = TargetBox
    
    If Response Then
        DoIt thisbox, lInt
    ElseIf (UCase(PL(lInt).Line) = "") Or _
            (UCase(PL(lInt).Line) = "LOOP") Or _
            (UCase(PL(lInt).Line) = "NEXT") Or _
            (UCase(PL(lInt).Line) = "END SUB") Or _
            (UCase(PL(lInt).Line) = "END IF") Then
    Else
        For j = 0 To Functions
            If PL(lInt).Line = Func(j).Name Then
                l = Func(j).FuncLen
                For k = 1 To l
                    lStr = PL(Func(j).LineNum(k)).Line
                    Parse lStr
                    PL(Func(j).LineNum(k)).StructString = LS.StructString
                    PL(Func(j).LineNum(k)).StructLen = LS.StructLen
                    Response = Valid(Func(j).LineNum(k)) 'Method of LEX
                    If Response Then
                        DoIt thisbox, (Func(j).LineNum(k))
                    End If
                Next
            End If
        Next
        If l = 0 Then
            ErrString = "Error #1   SYNTAX ERROR IN LINE# " + CStr(PL(lInt).LineNumber + 1) + _
            vbCrLf + CStr(PL(lInt).LineNumber + 1) + ": " + PL(lInt).Line + "  <--- ERROR."
            Exit Sub
        End If
    End If
    
    Set thisbox = Nothing
    Exit Sub
    
ErrHere:
    ErrString = "Err #102   ERROR IN INTRPRET (ModLanLanguage)"
End Sub

Public Sub DoIt(tbox As TextBox, plint As Integer)
'This sub uses the LineType of the current line
'   to jump to the corect precdure associated with
'   the keyword (if a keyword is present).
'
Dim i As Integer
Dim j As Integer
Dim vFound As Boolean
Dim tmpVal As Variant
Dim itemNum As Integer
Dim tmpLine As Integer
Dim forCount As Integer
Dim lInt As Integer
Dim FoundI As Integer
Dim lStr As String
Dim lDRet As DirReturn
Dim llTot As String

On Error GoTo ErrHere
If StopProgram Then Exit Sub
lInt = PL(plint).LineType
itemNum = -1
'-------------------------------
Select Case lInt
        '1** x = VALUE ************
    Case ValidInput.ASSIGN
        AssignTo plint
        
        '2** FOR i = x TO y ***
    Case ValidInput.FORNEXT
        LngFOR plint
        
        '3** IF x = y THEN z ****
    Case ValidInput.IFTHEN
        Conditional plint
        
        '4** PRINT Value **********
    Case ValidInput.LPRINT
        lStr = PrintAValue
        If Right(PL(plint).Line, 1) = ";" Then
            'lStr = Left(lStr, Len(lStr) - 1)
            tbox.Text = tbox.Text + lStr
            'tbox.SelStart = Len(tbox)
        Else
            tbox.Text = tbox.Text + lStr + vbCrLf
            'tbox.SelStart = Len(tbox)
        End If
        
        '5** PRINT String *******
    Case ValidInput.SPRINT
        lStr = PrintAString
        If Right(PL(plint).Line, 1) = ";" Then
            'lStr = Left(lStr, Len(lStr) - 1)
            tbox.Text = tbox.Text + lStr
            'tbox.SelStart = Len(tbox)
        Else
            tbox.Text = tbox.Text + lStr + vbCrLf
            'tbox.SelStart = Len(tbox)
        End If
    
        '6** INPUT VALUE x = 5 **********
    Case ValidInput.LINPUT
        InputVal tbox
        
        '7** INPUT STRING x = "cool" *******
    Case ValidInput.SINPUT
        InputString tbox
        
        '8** x = STRING *******
    Case ValidInput.SASSIGN
        AssignStrTo plint
        
        '9** DATE ***
    Case ValidInput.THEDATE
        lStr = ReturnDate
        If Right(PL(plint).Line, 1) = ";" Then
            lStr = Left(lStr, Len(lStr) - 1)
            tbox.Text = tbox.Text + lStr
        Else
            tbox.Text = tbox.Text + lStr + vbCrLf
        End If
        
        '10** TIME ***
    Case ValidInput.THETIME
        lStr = ReturnTime
        If Right(PL(plint).Line, 1) = ";" Then
            lStr = Left(lStr, Len(lStr) - 1)
            tbox.Text = tbox.Text + lStr
        Else
            tbox.Text = tbox.Text + lStr + vbCrLf
        End If
    
        '11** x = RANDOM opx opy ***
    Case ValidInput.RANDOMNUMBER
        RandNumber plint
       
        '12** x = CONCAT y + z ***
    Case ValidInput.CONCAT1
        ConcatTwoVars
        '13
    Case ValidInput.CONCAT2
    
        '14
    Case ValidInput.SINSTRING
    
        '15** SPACE (x) ***
    Case ValidInput.SSPACE
        lStr = PrintSpaces(plint)
        tbox.Text = tbox.Text + lStr
        
        '16** CLS ***
    Case ValidInput.SCLS
        tbox.Text = ""
        
        '17
    Case ValidInput.SSCREEN
        SetUpScreen
        
        '18
    Case ValidInput.SPLOT
        PlotToScreen plint
        
        '19
    Case ValidInput.SPRINTTO
    
        '20** x = INKEY ***
    Case ValidInput.SINKEY
        GetStringFromKey tbox
         
        '21** DO WHILE x = y ***
    Case ValidInput.SDOLOOP
        DoLoop plint
        
        '22** DO WHILE x = "cool" ***
    Case ValidInput.SDOLOOPS
        DoLoops plint
        
        '23** DO UNTIL x = y ***
    Case ValidInput.SDOLOOPU
        DoLoopU plint
        
        '24** DO UNTIL x = "cool" ***
    Case ValidInput.SDOLOOPUS
        DoLoopUs plint
        
        '25** IF x = "cool" THEN ***
    Case ValidInput.SIFTHENS
        ConditionalS plint
    
        '26** IF x <> y THEN ***
    Case ValidInput.SIFTHENNE
        ConditionalNE plint
    
        '27** IF x <> "cool" THEN ***
    Case ValidInput.SIFTHENNES
        ConditionalNES plint
    
        '28** IF x > y THEN ***
    Case ValidInput.SIFTHENG
        ConditionalG plint
    
        '29** IF x > "cool" THEN ***
    Case ValidInput.SIFTHENGS
        ConditionalGS plint
        
        '30** IF x < y THEN ***
    Case ValidInput.SIFTHENL
        ConditionalL plint
    
        '31** IF x < "cool" THEN  ***
    Case ValidInput.SIFTHENLS
        ConditionalLS plint
        
        '32** IF x > = y THEN ***
    Case ValidInput.SIFTHENGE
        ConditionalGE plint
    
        '33** IF x = > "cool" THEN ***
    Case ValidInput.SIFTHENGES
        ConditionalGES plint
        
        '34** IF x < = y THEN ***
    Case ValidInput.SIFTHENLE
        ConditionalLE plint
    
        '35** IF x < = "cool" THEN ***
    Case ValidInput.SIFTHENLES
        ConditionalLES plint
        
        '36** IF x = > y THEN ***
    Case ValidInput.SIFTHENEG
        ConditionalGE plint
    
        '37** IF x = > "cool" THEN ***
    Case ValidInput.SIFTHENEGS
        ConditionalGES plint
        
        '38** IF x = < y THEN ***
    Case ValidInput.SIFTHENEL
        ConditionalLE plint
    
        '39** IF x = < "cool" THEN ***
    Case ValidInput.SIFTHENELS
        ConditionalLES plint
    
        '40** DIR string ***
    Case ValidInput.SDIR
        lDRet = Directories(plint)
        tbox.Text = tbox.Text + lDRet.DirString + vbCrLf + _
        vbCrLf + "END LISTING" + vbCrLf + vbCrLf + _
        "There are " + CStr(lDRet.DirCount) + _
        " Directories " + " and " + CStr(lDRet.FileCount) + _
        " Files Listed." + vbCrLf
    
        '41** DISPLAYFILE string ***
    Case ValidInput.SDISPLAYFILE
        lStr = DisplayAFile(plint)
        tbox.Text = tbox.Text + lStr + vbCrLf
        
        '42** DIR ***
    Case ValidInput.SDIRPATH
        lDRet = ReturnPath()
        llTot = CStr(lDRet.TotalFileLen)
        llTot = Format(llTot, "###,###,###,###")
        tbox.Text = tbox.Text + lDRet.DirString + vbCrLf + _
        vbCrLf + "END LISTING" + vbCrLf + vbCrLf + _
        "Directory " + gPath + vbCrLf + _
        "There are " + CStr(lDRet.DirCount) + _
        " Directories " + " and " + CStr(lDRet.FileCount) + _
        " Files Listed." + vbCrLf + _
        "There are " + llTot + " Total Bytes " + _
        "in this directory" + vbCrLf
        
        '43** CHANGEDIR string ***
    Case ValidInput.SCHANGEDIR
        ChangePath plint
        
        '44** PATH ***
    Case ValidInput.SPATH
        lStr = ShowPath
        tbox.Text = tbox.Text + lStr + vbCrLf
    
        '45** DISPLAYFILE x ***
    Case ValidInput.SDISPLAYFILEVAR
        lStr = DisplayBFile(plint)
        tbox.Text = tbox.Text + lStr + vbCrLf
    
        '46** FUNCTION x ( y1 y2 ...) ***
    Case ValidInput.SFUNCTION
        
        '47** OPEN #1, "C:/MYFILE" ***
    Case ValidInput.SOPEN
    
        '48** CLOSE #1 ***
    Case ValidInput.SCLOSE
    
        '49** PRINT #1, x ***
    Case ValidInput.SPRINTF
    
        '50** INPUT #1, x ***
    Case ValidInput.SINPUTF
    
    Case ValidInput.ASSIGNPAR
        AssignTo plint
        
        
    Case ValidInput.ESHELLV
        UseAppV
    
    Case ValidInput.ESHELLL
        UseAppL
        
    Case ValidInput.ESENDKEYS
        UseAppKeys
        
    Case ValidInput.ESENDKEYSLATE
        UseAppKeysLate
        
End Select
Exit Sub
ErrHere:
ErrString = "Err #103   ERROR IN DOIT (ModLanLanguage)"
End Sub



