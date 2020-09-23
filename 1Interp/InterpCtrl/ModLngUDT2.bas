Attribute VB_Name = "ModLngUDT"

Public Type PorgramLines
    Line As String
    LineNumber As Integer
    LineType As Integer
    IsBlockHead As Boolean
    InBlock As Integer
    Block As Integer
    BlockLines As Integer
    Arg() As String
    Args As Integer
    TrueBlock As Boolean
    FalseBlock As Boolean
    StructString As String
    StructLen As Integer
End Type
Public PL() As PorgramLines 'ProgramLine

'Valid input is used by the lexical parser and
'   by the execution controller
'
Public Enum ValidInput
    ASSIGN = 1
    FORNEXT = 2
    IFTHENELSEbin = 3
    IFTHEN = 4
    IFTHENELSE = 5
    LPRINT = 6
    SPRINT = 7
    LINPUT = 8
    SINPUT = 9
    SASSIGN = 10
    THEDATE = 11
    THETIME = 12
    RANDOMNUMBER = 13
    CONCAT1 = 14
    CONCAT2 = 15
    SINSTRING = 16
    SSPACE = 17
    SCLS = 18
    SSCREEN = 19
    SPLOT = 20
    SPRINTTO = 21
    SINKEY = 22
    SDOLOOP = 23
    SDOLOOPS = 24
    SDOLOOPU = 25
    SDOLOOPUS = 26
    SIFTHENS = 27
    SIFTHENNE = 28
    SIFTHENG = 29
    SIFTHENL = 30
    SIFTHENGE = 31
    SIFTHENEG = 32
    SIFTHENLE = 33
    SIFTHENEL = 34
    SIFTHENNES = 35
    SIFTHENGS = 36
    SIFTHENLS = 37
    SIFTHENGES = 38
    SIFTHENEGS = 39
    SIFTHENLES = 40
    SIFTHENELS = 41
    SCHANGEDIR = 42
    SDISPLAYFILE = 43
    SDIRPATH = 44
    SDIR = 45
    SPATH = 46
    SDISPLAYFILEVAR = 47
    SFUNCTION = 48
    SOPEN = 49
    SCLOSE = 50
    SPRINTF = 51
    SINPUTF = 52
    ASSIGNPAR = 53
    ESHELLV = 54
    ESHELLL = 55
    ESENDKEYS = 56
    ESENDKEYSLATE = 57
End Enum

'These public constant structures are used by
'   the lexical parser
'   Numbers correspond to keys and symbols
'   V = Variable, L = Lieral, | = OR.
'   Layout: Public Const NAME = STRUCTURE REM FORM
'
Public Const formAssign = "0,37,0," '1 V = V|L                       Basic assignment, handles numerics
Public Const formFOR = "1,0,37,0,2,0," '2 FOR V = V|L                Basic loop
Public Const formIFTHEN = "4,0,37,0,5,0," '3IF V = V|L THEN          Basic conditional
Public Const formIFTHENELSE = "4,0,37,0,5," '4
Public Const formIFTHENbin = "4,0,5,0" '5
Public Const formPRINTVAL = "7,0," '6 PRINT V|N                      Basic print, handles variables and numeric literals
Public Const formPRINTSTRING = "7,46,0" '7 PRINT "L"                  Basic print, handles string literals
Public Const formINPUTVAL = "9,0" '8 INPUTVAL V                      Basic input, handles numeric input
Public Const formINPUTSTRING = "9,47,0" '9 INPUTVAL $V               Basic input, handles string input
Public Const formStrAssign = "0,37,46" '10 V = "L"                    Basic assignment , handles strings
Public Const formDATE = "10," '11 DATE                                Date function exposed
Public Const formTIME = "11," '12 TIME                                Time function exposed
Public Const formRANDOM = "0,37,12," '13 V = RANDOM                   Basic assignment, hadles exposed random function
Public Const formCONCAT1 = "0,37,13,0,38,0," '14 V = CONCAT V + V
Public Const formCONCAT2 = "0,37,13,0,38,46,0,46," '15 V = CONCAT V + "L"
Public Const formINSTRING = "0,37,14,0,0" '16 V = INSTRING V|L V|L
Public Const formSPACE = "15,42,0,43," '17 SPACE (V|L)                Space function exposed
Public Const formCLS = "16," '18 CLS                                  Basic clear screen
Public Const formSCREEN = "17,0,0," '19 SCREEN V|L V|L                 Used to define a screen area Widtn,Height
Public Const formPLOT = "18,0,0," '20 PLOT V|L V|L                    Used to plot a point to a defined screen
Public Const formPRINTTO = "0,37,19,0," '21 V = PRINTTO V|L           Basic print format control
Public Const formINKEY = "0,37,8," '22 V = INKEY
Public Const formDOLOOP = "20,21,0,37,0" '23 DO WHILE Var = Var|Lit
Public Const formDOLOOPS = "20,21,0,37,46" '24 DO WHILE Var = Lit
Public Const formDOLOOPU = "20,23,0,37,0" '25 DO UNTIL Var = Var|Lit
Public Const formDOLOOPUS = "20,23,0,37,46" '26 DO UNTIL Var = Lit
Public Const formIFTHENS = "4,0,37,46,0,46,5" '27 IF V = "L" THEN          Basic conditional
Public Const formIFTHENNE = "4,0,44,45,0,5" '28 IF V <> V|L THEN
Public Const formIFTHENG = "4,0,45,0,5" '29 IF V > V|L THEN
Public Const formIFTHENL = "4,0,44,0,5" '30 IF V < V|L THEN
Public Const formIFTHENGE = "4,0,45,37,0,5" '31 IF V >= V|L THEN
Public Const formIFTHENEG = "4,0,37,45,0,5" '32 IF V => V|L THEN
Public Const formIFTHENLE = "4,0,44,37,0,5" '33 IF V <= V|L THEN
Public Const formIFTHENEL = "4,0,37,44,0,5" '34 IF V =< V|L THEN
Public Const formIFTHENNES = "4,0,44,45,46,0,46,5" '35 IF V <> "L" THEN
Public Const formIFTHENGS = "4,0,45,46,0,46,5" '36 IF V > "L" THEN
Public Const formIFTHENLS = "4,0,44,46,0,46,5" '37 IF V < "L" THEN
Public Const formIFTHENGES = "4,0,45,37,46,0,46,5" '38 IF V >= "L" THEN
Public Const formIFTHENEGS = "4,0,37,45,46,0,46,5" '39 IF V => "L" THEN
Public Const formIFTHENLES = "4,0,44,37,46,0,46,5" '40 IF V <= "L" THEN
Public Const formIFTHENELS = "4,0,37,44,46,0,46,5" '41 IF V =< "L" THEN
Public Const formCHANGEDIR = "26,46,0,46" '42 CHANGEDIR "L"
Public Const formDISPLAYFILE = "25,46,0,46" '43 DISPLAYFILE "L"
Public Const formDIRPATH = "24," '44 DIR
Public Const formDIR = "24,46,0,46" '45 DIR "L"
Public Const formPATH = "27," '46 PATH
Public Const formDISPLAYFILEVAR = "25,0," ' 47 DISPLAYFILE V
Public Const formFUNCTION = "28,0," '48 'FUNCTION V
Public Const formOPEN = "29,0,49,46,0,46," '49 OPEN L, L
Public Const formCLOSE = "30,0," '50 CLOSE L
Public Const formPRINTF = "7,51,0,49,0," '51 PRINT # L, V|L"
Public Const formINPUTF = "9,51,0,49,0," '52 INPUT # L, V
Public Const formASSIGNPAR = "0,37,42,0," '53 V = (V|L)
Public Const formSHELLV = "31,0," '54 SHELL V
Public Const formSHELLL = "31,46,0,46" '55 SHELL "L"
Public Const formSENDKEYS = "32,46,0,46" '56 EXEC "L"
Public Const formSENDKEYSLATE = "32,0,49,46,0,46,"
Public Const LFNUM = 56 'The number of structures - 1

Public LF() As String 'LF will hold the structures

Public BlockHead() As Integer


Public Type LineStructure
    StructString As String
    StructLen As Integer
End Type

Public LS As LineStructure
