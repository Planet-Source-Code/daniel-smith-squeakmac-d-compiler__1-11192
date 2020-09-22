VERSION 5.00
Begin VB.Form frmRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D++ APP"
   ClientHeight    =   3720
   ClientLeft      =   4125
   ClientTop       =   3780
   ClientWidth     =   6960
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6960
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.TextBox txtIn 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6480
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   2640
   End
   Begin VB.Label input1 
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C)2000 D++ Linker
'Created by SqueakMac (webmaster@pagemac.zzn.com)
'http://pagemac.cjb.net
'Version 2.2.2

Private userinput, pausetime, box1, box2, textput, newvariable, numat
Private ifstate, WavVal, add1, add2, sub1, sub2, mul1, mul2, CurrentLine
Private filedel As Variant, inputvar As Variant, InIf As Boolean, InElse As Boolean
Private VarNames As New Collection, VarData As New Collection
Private Legnth As String

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
    txtText.SelStart = Len(txtIn.Text)
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If txtText.Locked = False Then
    If KeyAscii = vbKeyReturn Then
        If numat = 0 Then
            KeyAscii = 0
            Exit Sub
        Else
            txtText.Locked = True
        End If
    ElseIf KeyAscii = 8 Then
        If numat <= 0 Then
            KeyAscii = 0
            Exit Sub
        Else
            numat = numat - 1
            userinput = Mid(userinput, 1, Len(userinput) - 1)
        End If
    Else
        numat = numat + 1
        userinput = userinput & Chr(KeyAscii)
    End If
End If
End Sub

Private Sub Form_Load()
'On Error GoTo Errorh
    
    Me.Show
    
    'Open App.Path + "\" + App.EXEName + ".EXE" For Binary As #1
    Open "C:\WINDOWS\DESKTOP\D++APP1.EXE" For Binary As #1
    
    FileSize = LOF(1)
    FileData$ = Space$(LOF(1))
    
    Get #1, , FileData$
    
    For i = 1 To FileSize
        If Mid(FileData$, i, 4) = "DPP:" Then
            i = i + 4
            FileChunk$ = String(1000, 0)
            Get #1, i, FileChunk$
            txtIn.Text = FileChunk$
            If txtIn.Text = "" Then
                MsgBox "Error: No syntax found in program!", vbCritical, "Error"
                End
            End If
            LinkCode
            Exit Sub
        End If
    Next i
    
    Close #1
    
'Errorh:
'MsgBox "Error #" & Err.Number & " has occured: " & Err.Description, vbCritical, "Error"
End Sub

Sub LinkCode()
'Executes commands
Legnth = Len(txtIn.Text)
For i = 1 To Legnth
    d = 0
    
    If InElse = True Then
        If LCase(Mid(txtIn.Text, i, 5)) = "endif" Then
            InElse = False
            i = i + 5
        End If
    End If
    
    If InIf = True Then
        If LCase(Mid(txtIn.Text, i, 5)) = "endif" Then
            InIf = False
            i = i + 5
        ElseIf LCase(Mid(txtIn.Text, i, 4)) = "else" Then
            d = i
            Do Until LCase(Mid(txtIn.Text, i, 5)) = "endif"
                If i = d + 5000 Then
                    MsgBox "Syntax Error: Block if without endif: " & d, vbCritical, "Syntax Error"
                    End
                End If
                i = i + 1
            Loop
            InIf = False
            InElse = False
            i = i + 4
        End If
    Else
        If LCase(Mid(txtIn.Text, i, 4)) = "else" Then
            InElse = True
            InIf = False
            i = i + 4
        End If
    End If
    
    'output text to user
    If LCase(Mid(txtIn.Text, i, 11)) = "screenout """ Then
        i = i + 11
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Expected ';' at " & i & "; Found end of program.", vbCritical, "Error"
                End
            End If
            WriteText Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
    'output text to user all at once
    ElseIf LCase(Mid(txtIn.Text, i, 11)) = "screenput """ Then
        i = i + 11
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
        If i = d + 256 Then
            MsgBox "Expected ';' at " & i & "; Found end of program.", vbCritical, "Error"
            End
        End If
        textput = textput & Mid(txtIn.Text, i, 1)
        i = i + 1
        Loop
        txtText.Text = txtText.Text & textput
        textput = ""
        
    'output to user using variable
    ElseIf LCase(Mid(txtIn.Text, i, 10)) = "screenout " Then
        i = i + 10
        d = i
        
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            textput = textput & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop

        If FindVar(textput) = False Then
            MsgBox "Couldn't find the variable requested: " & TheVar, vbCritical, "Error"
            End
        Else
            WriteText GetVar(textput)
        End If
        textput = ""
    
    'output all at once using variable
    ElseIf LCase(Mid(txtIn.Text, i, 10)) = "screenput " Then
        i = i + 10
        d = i
        
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            textput = textput & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        If FindVar(textput) = False Then
            MsgBox "Couldn't find the variable requested: " & TheVar, vbCritical, "Error"
            End
        Else
            txtText.Text = txtText.Text & GetVar(textput)
        End If
        textput = ""
        
    'get input from user
    ElseIf LCase(Mid(txtIn.Text, i, 9)) = "screenin " Then
        i = i + 9
        d = i
        
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            inputvar = inputvar & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        If FindVar(inputvar) = True Then
            txtText.Locked = False
            wait = Timer
            numat = 0
            Do
                DoEvents
            Loop Until txtText.Locked = True
            SetVar inputvar, userinput
        Else
            MsgBox "Couldn't find the variable requested: " & TheVar, vbCritical, "Error"
            End
        End If
        inputvar = ""
        userinput = ""
        
    'title the application
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = "title """ Then
        i = i + 7
        d = i
        
        Me.Caption = ""
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            Me.Caption = Me.Caption & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        App.Title = Me.Caption

    'delete file
    ElseIf LCase(Mid(txtIn.Text, i, 8)) = "delete """ Then
        i = i + 8
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            filedel = filedel & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        If FileExist(filedel) = False Then
            MsgBox "Run Time Error: File not found", vbCritical, "Run Time Error"
            End
        Else
            Kill filedel
        End If
        
    'delete file by variable
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = "delete " Then
        i = i + 7
        d = i

        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            filedel = filedel & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        If FindVar(filedel) = True Then
            If FileExist(GetVar(filedel)) = False Then
                MsgBox "Error!  File not found!", vbCritical, "File Not Found"
            Else
                Kill GetVar(filedel)
            End If
        Else
            MsgBox "Couldn't find the variable requested: " & TheVar, vbCritical, "Error"
            End
        End If
        filedel = ""
        
    'comment
    ElseIf Mid(txtIn.Text, i, 1) = ">" Then
        i = i + 1
        Do Until Mid(txtIn.Text, i, 1) = Chr(13)
            i = i + 1
        Loop
        i = i + 1
        
    'create a message box
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = "box """ Then
        i = i + 5
        d = i
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box1 = box1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        i = i + 4
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            box2 = box2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        MsgBox box1, vbExclamation, box2
        box1 = ""
        box2 = ""
        
    'pause for given time
    ElseIf LCase(Mid(txtIn.Text, i, 6)) = "pause " Then
        i = i + 6
        d = i
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            pausetime = pausetime & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        Pause pausetime
        pausetime = ""
        
    'create a new variable
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = "newvar " Then
        i = i + 7
        d = i
        Do Until Mid(txtIn.Text, i, 1) = ";"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & d & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            newvariable = newvariable & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        If FindVar(newvariable) = True Then
            MsgBox "Run Time Error: " & newvariable & " variable already exits!", vbCritical, "Run Time Error"
            End
        Else
            VarNames.Add newvariable
            VarData.Add ""
        End If
        newvariable = ""
    
    'if statments
    ElseIf LCase(Mid(txtIn.Text, i, 3)) = "if " Then
        i = i + 3
        d = i
        Do Until LCase(Mid(txtIn.Text, i, 5)) = " then"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected 'then' at " & d & "; Found other.", vbCritical, "Syntax Error"
                End
            End If
            ifstate = ifstate & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        i = d
        
        If Eval(ifstate) = True Then
            InIf = True
            InElse = False
        ElseIf Eval(ifstate) = False Then
            Do Until LCase(Mid(txtIn.Text, i, 4)) = "else"
                If i = d + 5000 Then
                    i = d
                    Do Until LCase(Mid(txtIn.Text, i, 5)) = "endif"
                        If i = d + 5000 Then
                            MsgBox "Syntax Error: No else found: " & d, vbCritical, "Syntax Error"
                            End
                        End If
                        i = i + 1
                    Loop
                    Exit Do
                End If
                i = i + 1
            Loop
            InIf = False
            If LCase(Mid(txtIn.Text, i, 4)) = "else" Then
                InElse = True
            Else
                InElse = False
            End If
        End If
        ifstate = ""
        
    'clear console
    ElseIf LCase(Mid(txtIn.Text, i, 6)) = "clear;" Then
        i = i + 6
        txtText.Text = ""
    
    'Hide Console
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = "hide;" Then
        i = i + 5
        Me.Visible = False
    
    'Show Console
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = "show;" Then
        i = i + 5
        Me.Visible = True
        
    'end program
    ElseIf LCase(Mid(txtIn.Text, i, 4)) = "end;" Then
        i = i + 4
        End
        
    'a return
    ElseIf LCase(Mid(txtIn.Text, i, 7)) = "screen;" Then
        i = i + 7
        txtText.Text = txtText.Text & vbCrLf
        
    'add two values
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("add """) Then
        i = i + 5
        d = i
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            add1 = add1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        i = i + 4
        d = i
    
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            add2 = add2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(add1) + CDbl(add2)
        add1 = ""
        add2 = ""
        
    'subtract two numbers
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("sub """) Then
        i = i + 5
        d = i
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            sub1 = sub1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        i = i + 4
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            sub2 = sub2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(box1) - CDbl(box2)
        sub1 = ""
        sub2 = ""
        
    'multiply two numbers
    ElseIf LCase(Mid(txtIn.Text, i, 5)) = LCase("mul """) Then
        i = i + 5
        d = i
    
        Do Until Mid(txtIn.Text, i, 4) = """, """
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ',' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            mul1 = mul1 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        
        i = i + 4
        d = i
        
        Do Until Mid(txtIn.Text, i, 2) = """;"
            If i = d + 256 Then
                MsgBox "Syntax Error: Expected ';' at " & i & "; Found end of program.", vbCritical, "Syntax Error"
                End
            End If
            mul2 = mul2 & Mid(txtIn.Text, i, 1)
            i = i + 1
        Loop
        txtText.Text = txtText.Text & CDbl(mul1) * CDbl(mul2)
        mul1 = ""
        mul2 = ""
    
    Else
        If Mid(txtIn.Text, i, 1) <> "" And Mid(txtIn.Text, i, 1) = Chr(13) And Mid(txtIn.Text, i, 1) = Chr(10) And Mid(txtIn.Text, i, 1) <> " " Then
            MsgBox "Syntax Error: Invalid syntax at " & i & ". (" & Mid(txtIn.Text, i, 1) & ")", vbCritical, "Syntax Error"
        End If
    End If
    
    If Mid(txtIn.Text, i, 1) = "=" Then
        CurrentLine = ""
        For x = i To 1 Step -1
            If Mid(txtIn.Text, x, 1) = Chr(13) Then
                CurrentLine = Mid(CurrentLine, 1, (Len(CurrentLine) - 1))
                CurrentLine = StrReverse(CurrentLine)
                For y = i + 1 To Legnth
                    If Mid(txtIn.Text, y, 1) = Chr(13) Then
                        SetVar Trim(Mid(CurrentLine, 1, InStr(1, CurrentLine, "=") - 1)), Equation(Mid(CurrentLine, InStr(1, CurrentLine, "=") + 1))
                        Exit For
                    End If
                    CurrentLine = CurrentLine & Mid(txtIn.Text, y, 1)
                Next y
            End If
            CurrentLine = CurrentLine & Mid(txtIn.Text, x, 1)
        Next x
    End If
    
Next i
End Sub

Sub Pause(interval)
'Pauses
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub WriteText(TextToPut)
'Prints text to the screen
On Error Resume Next
For sd = 1 To Len(TextToPut)
txtText.SelStart = Len(txtText)
txtText.SelText = Mid(TextToPut, sd, 1)
DoEvents
Pause 0.01
Next sd
txtText.SelStart = Len(txtText)
End Sub

Function FileExist(ByVal FileName As String) As Boolean
'Determines if a file exists
    Dim fileFile As Integer
    fileFile = FreeFile
    On Error Resume Next
    Open FileName For Input As fileFile
    If Err Then
        FileExist = False
    Else
        Close fileFile
        FileExist = True
    End If
End Function

Private Function FindVar(TheVar As Variant) As Boolean
'Determines if a variable exists
For x = 1 To VarNames.Count
    If VarNames(x) = TheVar Then
        FindVar = True
        Exit Function
    End If
Next x
FindVar = False
End Function

Private Function GetVar(TheVar As Variant) As Variant
'Gets a variables value
For x = 1 To VarNames.Count
    If VarNames(x) = TheVar Then
        GetVar = VarData(x)
        Exit Function
    End If
Next x
End Function

Private Sub SetVar(TheVar As Variant, NewVal As Variant)
'Sets the value of a variable
For x = VarNames.Count To 1 Step -1
    If VarNames(x) = TheVar Then
        VarNames.Remove x
        VarData.Remove x
        VarNames.Add TheVar
        VarData.Add NewVal
        Exit Sub
    End If
Next x
End Sub

Private Function Eval(ByVal sFunction As String) As Boolean
'This parses a string into a left value, operator, and right value
Dim LeftVal, RightVal, Operator
Dim sChar, OpFound As Boolean
OpFound = False

    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        If sChar = ">" Or sChar = "<" Or sChar = "=" Then
            Operator = Operator & sChar
            OpFound = True
        Else
            If OpFound = True Then
                RightVal = RightVal & sChar
            Else
                LeftVal = LeftVal & sChar
            End If
        End If
    Next x
    
    LeftVal = Equation(LeftVal)
    RightVal = Equation(RightVal)
    
    Select Case Operator
        Case ">"
            If LeftVal > RightVal Then Eval = True
        Case "<"
            If LeftVal < RightVal Then Eval = True
        Case "="
            If LeftVal = RightVal Then Eval = True
        Case "<>"
            If LeftVal <> RightVal Then Eval = True
        Case ">="
            If LeftVal >= RightVal Then Eval = True
        Case "<="
            If LeftVal <= RightVal Then Eval = True
        Case Else
            MsgBox "Syntax Error: Invalid operator: " & Operator, vbCritical, "Syntax Error"
            End
    End Select
End Function

Private Function Equation(ByVal sFunction As String) As Variant
'This sub basically looks for parentheses, and solves what's in them
Dim Paren1 As Integer, Paren2 As Integer, sChar As String
Do
    For x = 1 To Len(sFunction)
        sChar = Mid(sFunction, x, 1)
        Select Case sChar
            Case Chr(34) 'Character 34 is the "
                x = InStr(x + 1, sFunction, Chr(34))
            Case "("
                Paren1 = x
            Case ")"
                Paren2 = x
                Exit For
        End Select
    Next x
    If Paren1 = 0 Then
        Exit Do
    Else
        sFunction = Mid(sFunction, 1, Paren1 - 1) & " " & Chr(34) & Solve(Mid(sFunction, Paren1 + 1, Paren2 - (Paren1 + 1))) & Chr(34) & " " & Mid(sFunction, Paren2 + 1)
        Paren1 = 0
        Paren2 = 0
    End If
Loop
Equation = Solve(sFunction)
End Function

Private Function Solve(sFunction As String) As Variant
'This sub solves equations like name="SqueakMac" or "5" + ("7" * "3")
Dim Quote As Integer, sChar As String, variable As String
Dim Num2 As Variant, Operator As String, Num1
    
For x = 1 To Len(sFunction)
    sChar = Mid(sFunction, x, 1)
    If sChar = Chr(34) Then
        Quote = InStr(x + 1, sFunction, Chr(34))
        Num2 = Mid(sFunction, x + 1, Quote - (x + 1))
        x = Quote
        If Operator <> "" Then
            Select Case Operator
                Case "+"
                    Solve = Solve + Num2
                Case "-"
                    Solve = Solve - Num2
                Case "/"
                    Solve = Solve / Num2
                Case "\"
                    Solve = Solve \ Num2
                Case "^"
                    Solve = Solve ^ Num2
                Case "*"
                    Solve = Solve * Num2
                Case "&"
                    Solve = Solve & Num2
            End Select
            Operator = ""
        Else
            Solve = Num2
        End If
    ElseIf sChar = "+" Or sChar = "-" Or sChar = "/" Or sChar = "\" Or sChar = "^" Or sChar = "&" Or sChar = "*" Then
        If Num1 <> 0 Then
            Num2 = GetVar(Trim(Mid(sFunction, Num1, x - (Num1 + 1))))
            If Operator <> "" Then
                Select Case Operator
                    Case "+"
                        Solve = Solve + Num2
                    Case "-"
                        Solve = Solve - Num2
                    Case "/"
                        Solve = Solve / Num2
                    Case "\"
                        Solve = Solve \ Num2
                    Case "^"
                        Solve = Solve ^ Num2
                    Case "*"
                        Solve = Solve * Num2
                    Case "&"
                        Solve = Solve & Num2
                End Select
                Operator = ""
            Else
                Solve = Num2
            End If
                
            Num1 = 0
        End If
            
        Operator = sChar
    Else
        If Num1 = 0 Then Num1 = x
        If x >= Len(sFunction) Then
            variable = Trim(Mid(sFunction, Num1, x))
            Num2 = GetVar(variable)
            If Operator <> "" Then
                Select Case Operator
                    Case "+"
                        Solve = Solve + Num2
                    Case "-"
                        Solve = Solve - Num2
                    Case "/"
                        Solve = Solve / Num2
                    Case "\"
                        Solve = Solve \ Num2
                    Case "^"
                        Solve = Solve ^ Num2
                    Case "*"
                        Solve = Solve * Num2
                    Case "&"
                        Solve = Solve & Num2
                End Select
                Operator = ""
            Else
                Solve = Num2
            End If
        End If
    End If
Next x
End Function

