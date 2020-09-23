Attribute VB_Name = "ModLngHelp"
Dim Demo() As String
Dim DName() As String
Dim Limit As Integer
Dim Examples(3) As String
Public OpenSave As Integer

Public Function ShowHelpMe(lInt As Integer) As String
Dim lReturn As String
    lReturn = ""
    If lInt > -1 And lInt < Limit Then
        lReturn = Demo(lInt)
    End If
    ShowHelpMe = lReturn
End Function

Public Function InitHelp() As Integer
    Limit = 21
    ReDim Demo(Limit)
    ReDim DName(Limit)
    
    Demo(1) = "ASSIGNMENTS" + _
        vbCrLf + "Allows a variable to take on a value" + vbCrLf + _
        vbCrLf + "Structure" + _
        vbCrLf + "    x = A" + _
        vbCrLf + "    x = A <OP> B <OP> C ..." + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "    x = 5                      'Assigns to x the value 5" + _
        vbCrLf + "    x = " + Chr(34) + "This is a Litteral String" + Chr(34) + " 'Assigns to x the Litteral String" + _
        vbCrLf + "    x = y                      'Assigns to x the value of y" + vbCrLf + _
        vbCrLf + "In structure 1, A can be either a Numeric or String Litteral," + _
        vbCrLf + "or a Numeric or String Varriable" + _
        vbCrLf + "In structure 2, A,B,and C can be either Literral Numeric Values," + _
        vbCrLf + "or Numeric Varriables" + _
        vbCrLf + "<OP> must be from the following set: {+,-,*,/,^}"
    Demo(2) = "IF THEN : END IF" + _
        vbCrLf + "Conditional structure that allows comparrisons" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "IF A <OP> B THEN" + _
        vbCrLf + "      <Your Code>" + _
        vbCrLf + "END IF" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "If x = 5 Then       'Only do the code if x is holding the value 5" + _
        vbCrLf + "      Print " + Chr(34) + "You Typed in the number 5" + Chr(34) + _
        vbCrLf + "END IF              'End the conditional code" + vbCrLf + _
        vbCrLf + "If x <> MyName Then" + _
        vbCrLf + "      Print " + Chr(34) + "You are not allowed access" + Chr(34) + _
        vbCrLf + "      NoPass = True" + _
        vbCrLf + "END IF" + vbCrLf + _
        vbCrLf + "In the Strucure, A must be either a Numeric or String Varriable," + _
        vbCrLf + "while B can be either a Numeric or String Varriable, or a Numeric" + _
        vbCrLf + "or String Litteral." + _
        vbCrLf + "<OP> must be from the following set: {=,<,<=,=<,=>,>=,<>}"
    Demo(3) = "DO WHILE : LOOP         DO UNTIL : LOOP " + _
        vbCrLf + "Recursive structure that allows reptition of a task, while, or until" + _
        vbCrLf + "a condition is met" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "DO WHILE A = B          DO UNTIL A = B" + _
        vbCrLf + "      <Your Code>                   <Your Code>" + _
        vbCrLf + "LOOP                             LOOP" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "DO WHILE x = " + Chr(34) + "N" + Chr(34) + _
        vbCrLf + "      Print " + Chr(34) + "Would you like to quit" + Chr(34) + " ;" + _
        vbCrLf + "      Input x" + _
        vbCrLf + "Loop" + vbCrLf + _
        vbCrLf + "y = 5" + _
        vbCrLf + "DO Until x = y      'Repeate the code until x and y have the same value" + _
        vbCrLf + "      t = t + x * 4" + _
        vbCrLf + "      print t" + _
        vbCrLf + "Loop                     'End the loop" + vbCrLf + _
        vbCrLf + "In the structure, A is a Numeric or String Varriable," + _
        vbCrLf + "while B can be either a Numeric or String Variable or Litteral" + _
        vbCrLf + "This structure can be used to control Program Termination"
    Demo(4) = "FOR TO : NEXT" + _
        vbCrLf + "Recursive structure that allows a specified number of repetitions " + _
        vbCrLf + "of a task" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "FOR A = B TO C" + _
        vbCrLf + "   <Your Code>" + _
        vbCrLf + "NEXT" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "For i = 1 To 25       'Repete the code 25 times" + _
        vbCrLf + "   x= x+ 1" + _
        vbCrLf + "   Print x" + _
        vbCrLf + "Next                        'End the loop" + vbCrLf + _
        vbCrLf + "For i = x To y" + _
        vbCrLf + "   z = z +(y + x^2)" + _
        vbCrLf + "Next" + vbCrLf + _
        vbCrLf + "In the structure, A (the index) is a Numeric Varriable," + _
        vbCrLf + "while B can be either a Numeric Variable or Numeric Litteral" + _
        vbCrLf + "This is a standard structure used to iterate through an Index"
    Demo(5) = "INPUT" + _
        vbCrLf + "This basic function allows the User to interact with your program" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "INPUT A         INPUT $A" + vbCrLf + _
        vbCrLf + "Exapmles:" + _
        vbCrLf + "Do Until x = " + Chr(34) + "quit" + Chr(34) + _
        vbCrLf + "   Print " + Chr(34) + "Do you want to exit the program" + Chr(34) + " ;" + _
        vbCrLf + "   Input $x       'x holds the string entered by the User" + _
        vbCrLf + "Loop" + vbCrLf + _
        vbCrLf + "Input x           'x holds the value enterd by the User" + _
        vbCrLf + "t = x*5/2" + _
        vbCrLf + "Print t" + vbCrLf + _
        vbCrLf + "There are two types of Input, Integer and String." + _
        vbCrLf + "To get a String Value from the User, prface the variable with" + _
        vbCrLf + "the Dollar Symbol ($) as in Example 1." + _
        vbCrLf + "To get an Integer Value from the User, do not preface the Variable."
    Demo(6) = "INKEY" + _
        vbCrLf + "This function returns a User Keystroke to a Variable" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "A = INKEY" + vbCrLf + _
        vbCrLf + "Exapmle:" + _
        vbCrLf + "x = Inkey     ' x will hold the character of the Users keystoke" + _
        vbCrLf + "Print " + Chr(34) + "You just typed in " + Chr(34) + " ;" + _
        vbCrLf + "Print x" + vbCrLf + _
        vbCrLf + "In the Structure, A will be a String Variable if it is new." + _
        vbCrLf + "To Make A an Integer, write the assignment A = 0 before using Inkey"
    Demo(7) = "RANDOM" + _
        vbCrLf + "Used to generate a random number" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "A = RANDOM" + _
        vbCrLf + "A = RANDOM B" + _
        vbCrLf + "A = RANDOM B C" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "x = Random            'Generates a Number between 0 and 100" + _
        vbCrLf + "x = Random 5         'Gererates a Number between 0 and 5" + _
        vbCrLf + "x = Random 3 36    'Generates a Number between 3 and 36" + _
        vbCrLf + "x = Random x y      'Generates a Number between x and y" + _
        vbCrLf + "NOTE: You cannot have a Comment on the same line as Structures 1 and 2"
    Demo(8) = "CONCAT" + _
        vbCrLf + "Used to add Strings together" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "A = CONCAT B + C" + vbCrLf + _
        vbCrLf + "Example:" + _
        vbCrLf + "y = " + Chr(34) + "That" + Chr(34) + _
        vbCrLf + "x = CONCAT " + Chr(34) + "This And" + Chr(34) + " + y" + vbCrLf + _
        vbCrLf + "In the structuer, A is a String Variable, B and C are either" + _
        vbCrLf + "String Literrals, or String Variables"
    Demo(9) = "PRINT" + _
        vbCrLf + "This function is used to let your program communcaite with" + _
        vbCrLf + "the User" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "PRINT A" + _
        vbCrLf + "PRINT A ;" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "Print " + Chr(34) + "Hello" + Chr(34) + _
        vbCrLf + "Print 5" + _
        vbCrLf + "Print x" + _
        vbCrLf + "Print x ;" + vbCrLf + _
        vbCrLf + "Use the Semicolon (;) to Append the next Print to the same Line"
    Demo(10) = "CLS" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "CLS" + vbCrLf + _
        vbCrLf + "Example:" + _
        vbCrLf + "Print " + Chr(34) + "Input 1 to Clear the screen" + Chr(34) + _
        vbCrLf + "Input a" + _
        vbCrLf + "If a = 1 Then" + _
        vbCrLf + "      Cls            ' Clears the Program Window" + _
        vbCrLf + "End If" + vbCrLf + _
        vbCrLf + "Use this function to Clear the Program Window"
    Demo(11) = "SPACE" + _
        vbCrLf + "This function Prints a series of spaces" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "SPACE (A)" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "Space (20)     'Prints 20 spaces" + _
        vbCrLf + "Space (x)      'Prints x Number of spaces" + vbCrLf + _
        vbCrLf + "Use this function to format your programs output"
    Demo(12) = "DIR" + _
        vbCrLf + "Returns a Listing of a Directories (Folder) contents to the Program Window" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "DIR A" + _
        vbCrLf + "DIR" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "Dir " + Chr(34) + "c:\MyDir" + Chr(34) + "       'Lists the contents of the Directory MyDir" + _
        vbCrLf + "Dir  d                     'd is a Variable that holds the Directories path" + _
        vbCrLf + "Dir                         'Lists the contents of the Current Directory" + vbCrLf + _
        vbCrLf + "Use Dir to get a listing of a directories contents"
    Demo(13) = "PATH" + _
        vbCrLf + "Returns the Current Directory to the Program Window" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "PATH" + vbCrLf + _
        vbCrLf + "Example:" + _
        vbCrLf + "Path        'Example Result = c:\MyDir" + vbCrLf + _
        vbCrLf + "Use this function to verify the Current Directory"
    Demo(14) = "CHANGEDIR" + _
        vbCrLf + "Changes the Current Directory" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "CHANGDIR A" + vbCrLf + _
        vbCrLf + "Example:" + _
        vbCrLf + "ChangeDir " + Chr(34) + "c:\ADir" + Chr(34) + vbCrLf + _
        vbCrLf + "Use this function to changr the Current Directory"
    Demo(15) = "DISPLAYFILE" + _
        vbCrLf + "Returns a files contents to the Program Window" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "DISPLAYFILE A" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "DisplayFile " + Chr(34) + "C:\MyText.txt" + _
        vbCrLf + "DisplayFile x" + vbCrLf + _
        vbCrLf + "In the structure, A can be either a String Litteral, or a String Vairable" + _
        vbCrLf + "Note: There is a 15KB Limit on the File Size"
    Demo(16) = "SENDKEYS" + _
        vbCrLf + "Sends A String To The Shelled Application" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "SENDKEYS A" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "SHELL " + Chr(34) + "C:/Windows/Calc.EXE" + Chr(34) + _
        vbCrLf + "SENDKEYS " + Chr(34) + "5 {+} 5 {=}" + Chr(34) + vbCrLf + _
        vbCrLf + "In the structure, A can be either a String Litteral, or a String Vairable" + _
        vbCrLf + "Note: In this version onle ONE String/Variable can be sent per SENDKEYS COMMAND"
    Demo(17) = "SHELL" + _
        vbCrLf + "Starts Another Application" + vbCrLf + _
        vbCrLf + "Structure:" + _
        vbCrLf + "SHELL A" + vbCrLf + _
        vbCrLf + "Examples:" + _
        vbCrLf + "SHELL " + Chr(34) + "C:/Windows/Calc.EXE" + Chr(34) + _
        vbCrLf + "SENDKEYS " + Chr(34) + "5 {+} 5 {=}" + Chr(34) + vbCrLf + _
        vbCrLf + "In the structure, A can be either a String Litteral, or a String Vairable" + _
        vbCrLf + "Note: In this version onle ONE String/Variable can be sent per SHELL COMMAND"
    Demo(18) = "Quick Intro" + _
        vbCrLf + "Below is short guide line to get you started" + vbCrLf + _
        vbCrLf + "To start a program you have saved on the DeskTop, Just Drag and Drop the program" + _
        vbCrLf + "onto the BSSOK BASIC Shortcut (See Windows Help for setting up ShortCuts)" + vbCrLf + _
        vbCrLf + "You can use the Code Window to write programs, or any other Text Editor" + _
        vbCrLf + "as long as you save the File as Text (a .txt extension)" + vbCrLf + _
        vbCrLf + "Use the Help Files for more Information"
    Demo(19) = "Known Bugs" + _
        vbCrLf + "At this time there are no BUGS known that can adversly affect your system" + vbCrLf + _
        vbCrLf + "Arithmetic:" + _
        vbCrLf + "The arithmetic routine will not work with imbedded parenthesies in this version" + _
        vbCrLf + "Full Mathematical Stements will be allowed in the next version" + vbCrLf + _
        vbCrLf + "Only Two Data Types:" + _
        vbCrLf + "There are only Integer and String Data Types in this version." + vbCrLf + _
        vbCrLf + "TAB:" + _
        vbCrLf + "There is a problem with iterpreting TAB characters in this version" + _
        vbCrLf + "The work around is to use the SPACE Key instead of the TAB Key" + _
        vbCrLf + "when writing code" + vbCrLf + _
        vbCrLf + "If you find any BUGs using this program, please send an email to" + _
        vbCrLf + "bugs@bssok.bizhosting.com" + _
        vbCrLf + "include a discription of the problem along with as much detail as" + _
        vbCrLf + "possible on when it occurs"
    Demo(20) = "ABOUT MY DESKTOPBASICS" + _
        vbCrLf + "This is an APLHA Version of BASIC using The Language Creation Core" + vbCrLf + _
        vbCrLf + "MY DESKTOPBASICS is a small, functional BASIC Interpreter intended for the DeskTop. " + _
        vbCrLf + "This ALPHA Version has only  two types: Strings and Integers." + _
        vbCrLf + "As it develops, MY DESKTOPBASICS will incorporate more Types and Keywords." + _
        vbCrLf + "After the Core is completed. There are plans for a small integrated compiler. " + _
        vbCrLf + "If you have any comments or suggestions, please send an email to " + _
        vbCrLf + "basic@bssok.bizhosting.com" + vbCrLf + _
        vbCrLf + "If you want to be put on an upgrade list for future versions as they" + _
        vbCrLf + "become available, send an email to " + _
        vbCrLf + "addme@bssok.bizhosting.com" + vbCrLf + _
        vbCrLf + "We thank you for using MY DESKTOPBASICS" + vbCrLf + _
        vbCrLf + "'MY DESKTOPBASICS' Copyright 1999 BSSOK" + _
        vbCrLf + "'The Language Creation Core' Copyright 1999 Robert Spoons"
    
        
    DName(0) = "KeyWords"
    DName(1) = "Assignment"
    DName(2) = "IF...THEN...END"
    DName(3) = "DO Loops"
    DName(4) = "FOR Loops"
    DName(5) = "INPUT"
    DName(6) = "INKEY"
    DName(7) = "RANDOM"
    DName(8) = "CONACT"
    DName(9) = "PRINT"
    DName(10) = "CLS"
    DName(11) = "SPACE"
    DName(12) = "DIR"
    DName(13) = "PATH"
    DName(14) = "CHANGEDIR"
    DName(15) = "DISPLAYFILE"
    DName(16) = "SENDKEYS"
    DName(17) = "SHELL"
    DName(18) = "Quick Intro"
    DName(19) = "Known Bugs"
    DName(20) = "ABOUT BSSOKBASIC"
    InitHelp = Limit
    
    Examples(0) = "DO UNTIL X = " + Chr(34) + "Q" + Chr(34) + _
    vbCrLf + "CLS" + _
    vbCrLf + "GAMES=GAMES+1" + _
    vbCrLf + "PRINT " + Chr(34) + "I AM THINKING OF A NUMBER BTEWEEN 1 AND 2" + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "TRY GUESSING WHAT IT IS" + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "YOUR GUESS " + Chr(34) + " ;" + _
    vbCrLf + "RESPONSE = " + Chr(34) + "WRONG" + Chr(34) + _
    vbCrLf + "R = RANDOM 1 2" + _
    vbCrLf + "INPUT Guess" + _
    vbCrLf + "IF Guess = R THEN" + _
    vbCrLf + "         SCORE = SCORE + 2" + _
    vbCrLf + "         CORRECT = CORRECT + 1" + _
    vbCrLf + "         RESPONSE = " + Chr(34) + "CORRECT" + Chr(34) + _
    vbCrLf + "END IF" + _
    vbCrLf + "PRINT " + Chr(34) + "MY NUMBER WAS    " + Chr(34) + " ;" + _
    vbCrLf + "PRINT R" + _
    vbCrLf + "PRINT " + Chr(34) + "YOUR GUESS WAS    " + Chr(34) + " ;" + _
    vbCrLf + "PRINT RESPONSE" + _
    vbCrLf + "SCORE = SCORE - 1" + _
    vbCrLf + "PRINT " + Chr(34) + "YOUR SCORE IS NOW    " + Chr(34) + " ;" + _
    vbCrLf + "PRINT SCORE" + _
    vbCrLf + "PRINT " + Chr(34) + " " + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "DO YOU WANT TO SEE YOUR SCORE? (Y/N): " + Chr(34) + " ;" + _
    vbCrLf + "ANSWER = INKEY" + _
    vbCrLf + "IF ANSWER = " + Chr(34) + "Y" + Chr(34) + " THEN" + vbCrLf
    Examples(1) = "            CLS" + _
    vbCrLf + "SPACE (10)" + _
    vbCrLf + "            DATE" + _
    vbCrLf + "SPACE (10)" + _
    vbCrLf + "             TIME" + _
    vbCrLf + "            PRINT " + Chr(34) + "YOU HAVE PLAYED " + Chr(34) + " ;" + _
    vbCrLf + "            PRINT GAMES ;" + _
    vbCrLf + "            PRINT " + Chr(34) + " GAMES" + _
    vbCrLf + "            PRINT " + Chr(34) + "YOU HAVE WON " + Chr(34) + " ;" + _
    vbCrLf + "            PRINT CORRECT ;" + _
    vbCrLf + "            PRINT " + Chr(34) + "OF THEM." + Chr(34) + _
    vbCrLf + "            PRINT " + Chr(34) + "YOUR CURRENT SCORE IS " + Chr(34) + " ;" + _
    vbCrLf + "            PRINT SCORE" + _
    vbCrLf + "            PRINT " + Chr(34) + " " + Chr(34) + _
    vbCrLf + "            PRINT " + Chr(34) + "(Hit any key to continue)" + Chr(34) + _
    vbCrLf + "            DUMMY = INKEY" + _
    vbCrLf + "END IF" + _
    vbCrLf + "PRINT " + Chr(34) + " " + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "Any Key Except 'Q' Continues The Game" + Chr(34) + " ;" + _
    vbCrLf + "X = INKEY" + _
    vbCrLf + "CLS" + _
    vbCrLf + "LOOP" + _
    vbCrLf + "DATE" + _
    vbCrLf + "TIME" + _
    vbCrLf + "PRINT " + Chr(34) + "AFTER PLAING " + Chr(34) + " ;" + vbCrLf
    Examples(2) = "PRINT GAMES ;" + _
    vbCrLf + "PRINT " + Chr(34) + "GAMES" + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "YOUR SCORE IS " + Chr(34) + " ;" + _
    vbCrLf + "PRINT SCORE" + _
    vbCrLf + "PRINT " + Chr(34) + " " + Chr(34) + _
    vbCrLf + "PRINT " + Chr(34) + "PLAY AGAIN SOON" + _
    vbCrLf + "PRINT " + Chr(34) + " " + Chr(34) + _
    vbCrLf + "FOR i = 1 TO 100" + _
    vbCrLf + "NEXT" + _
    vbCrLf + "PRINT " + Chr(34) + "ANY KEY TERMINATES THE PROGRAM" + _
    vbCrLf + "X = INKEY"
    
End Function

Public Function HelpName(lInt As Integer) As String
Dim lReturn As String

    lReturn = ""
    
    If lInt > -1 And lInt < Limit Then
        lReturn = DName(lInt)
    End If
    
    HelpName = lReturn
    
End Function

Public Function Example(lInt As Integer) As String

    Example = Examples(lInt)
    
End Function
