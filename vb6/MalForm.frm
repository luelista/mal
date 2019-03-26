VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MalForm 
   Caption         =   "Transcript"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Align           =   4  'Rechts ausrichten
      BackColor       =   &H00C0FFC0&
      Height          =   9930
      Left            =   10110
      ScaleHeight     =   9870
      ScaleWidth      =   660
      TabIndex        =   2
      Top             =   0
      Width           =   720
      Begin VB.CommandButton cmdActions 
         Caption         =   "edit"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox txtInput 
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   8880
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MalForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   8775
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   15478
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"MalForm.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim env As New LHashmap
Dim actionIdx As Integer
Dim actions As New LList

Sub runTests(lines() As String)
resetEnv
    Dim loopvar
    Dim l As String
    Dim output  As String
    For Each loopvar In lines
        l = loopvar
        If Len(l) >= 1 Then
            Select Case Mid(l, 1, 2)
            Case ";="
                If output = Mid(l, 4) Then
                    append "SUCCESS " + l, &H559955
                Else
                    append "FAIL got: " + output, &H5555FF
                    append "expected: " + l, &H1166DD
                End If
            Case ";;"
                append l, &H555555
            Case ";>"
                append l, &H995555
            Case ";/"
                If Mid(output, 1, 5) = "ERROR" Then
                    append "SUCCESS: " + output + " (expected: " + l + ")", &H1166AA
                Else
                    append "FAIL got: " + output + " (expected error: " + l + ")", &H1166AA
                End If
            Case Else
                append "TEST: " & l, &H666666
                output = rep(l)
            End Select
        End If
    Next
    
End Sub

Function rep(s As String) As String


On Error GoTo ex
'    Dim r As New Reader
'    r.tokenize (s)
'    Dim tokens() As String
'    tokens = r.getTokens()
'    Dim i  As Long
'    For i = 0 To UBound(tokens)
'        append (">>" + tokens(i) + "<<")
'    Next
'
'    Dim frm As Variant
'     r.read_form frm
'    append ("-->" + pr_str(frm))

    Dim frm, output As Variant
    read_str s, frm
    eval frm, env, output
    rep = pr_str(output, True)

    Exit Function
ex:
    'rep = "ERROR:" + Err.Description
    rep = ""
    append "ERROR:" + Err.Description, vbRed
End Function

Sub resetEnv()
    Set env = New LHashmap
    declareInternalFunctions env
    
    env.add "prn", newHostFunction(Me, "repl_prn")
    env.add "println", newHostFunction(Me, "repl_println")
    
    env.add "eval", newHostFunction(Me, "repl_eval")
    
    env.add "clear-transcript", newHostFunction(Me, "repl_clear")
    env.add "create-object", newHostFunction(Me, "repl_createobject")
    env.add "add-action", newHostFunction(Me, "addAction")
    env.add "save-setting", newHostFunction(Me, "repl_saveSetting")
    env.add "get-setting", newHostFunction(Me, "repl_getSetting")
    env.add "show-editor", newHostFunction(Me, "repl_showEditor")
    
    env.add "*host-language*", "Visual Basic 6 (GUI)"
    
    env.add "transcript-form", Me
    env.add "editor-form", MalEdit
    
    
    'env.add "", newInternalFunction(lif_)
    
    rep internalDefs

    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdActions_Click(Index As Integer)
    Dim output As Variant
    On Error GoTo e
    eval actions.item(Index), env, output
    If IsNothing(output) Then
    Else
        append "actionbutton output: " & pr_str(output, True)
    End If
    Exit Sub
e:
    append "actionbutton ERROR: " & Err.Description, vbRed
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE Or KeyCode = vbKeyT) And Shift = vbCtrlMask Then
        MalEdit.Show
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
'    ChDir "C:\_test\lisp\interpreter"
'    txtTests.Text = FileToString("testcases.txt")
    resetEnv
    rep "(println (str ""Mal ["" *host-language* ""]""))"
    actionIdx = 0
    
    Dim l As String: l = GetSetting("MalVB", "GUI", "startup-script", "")
    If l <> "" Then
        rep "(load-file " + pr_str(l, True) + ")"
    End If
End Sub

Private Sub Form_Resize()
    txtInput.Top = Me.Height - 1400
    RichTextBox1.Height = Me.Height - 1600
    txtInput.Width = Me.Width - 800
    RichTextBox1.Width = Me.Width - 800
End Sub


Function addAction(params As LList) As Variant
    Dim i As Integer
    If params.length >= 3 Then
        i = params.item(2)
    Else
        i = actionIdx + 1
    End If
    actionIdx = i
    
    On Error Resume Next
    Load cmdActions(i)
    On Error GoTo 0
    cmdActions(i).Left = 0
    cmdActions(i).Top = i * 500
    actions.assign i, params.item(1)
    cmdActions(i).Caption = params.item(0)
    cmdActions(i).Visible = True
    
    Set addAction = Nothing
End Function

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        append txtInput.Text
        append "; " + rep(txtInput.Text)
        KeyCode = 0
    End If
End Sub

Function repl_eval(params As LList) As Variant
    eval params.item(0), env, repl_eval
End Function


Function repl_prn(params As LList) As Variant
        append print_multi(params, True, " "), &H999900
        Set repl_prn = Nothing
End Function
Function repl_println(params As LList) As Variant
        append print_multi(params, False, " "), &H999900
        Set repl_println = Nothing
End Function

Function repl_clear(params As LList) As Variant
    RichTextBox1.Text = ""
    Set repl_clear = Nothing
End Function

Function repl_saveSetting(params As LList) As Variant
    SaveSetting "MalVB", params.item(0), params.item(1), params.item(2)
    Set repl_saveSetting = Nothing
End Function
Function repl_getSetting(params As LList) As Variant
    repl_getSetting = GetSetting("MalVB", params.item(0), params.item(1))
End Function

Function repl_createobject(params As LList) As Variant
    Set repl_createobject = CreateObject(params.item(0))
End Function

Function repl_showEditor(params As LList) As Variant
    MalEdit.Show
    
    Set repl_showEditor = Nothing
End Function

Sub append(s As String, Optional col As ColorConstants = vbBlack)

RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelColor = col
RichTextBox1.SelText = s + vbNewLine
RichTextBox1.SelStart = RichTextBox1.SelStart + Len(s) + 2

End Sub

