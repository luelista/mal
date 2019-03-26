VERSION 5.00
Begin VB.Form MalEdit 
   Caption         =   "MalEdit"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   672
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Align           =   4  'Rechts ausrichten
      BackColor       =   &H00C0FFFF&
      Height          =   9345
      Left            =   8850
      ScaleHeight     =   9285
      ScaleWidth      =   2550
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2610
      Begin VB.FileListBox File1 
         Height          =   5160
         Left            =   90
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3795
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   555
         Width           =   2415
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         Width           =   2415
      End
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Width           =   7815
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Unten ausrichten
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9345
      Width           =   11460
      Begin VB.CommandButton cmdRunsel 
         Caption         =   "run selection"
         Height          =   495
         Left            =   3465
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "save"
         Height          =   495
         Left            =   6840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdRunscript 
         Caption         =   "run script"
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdRuntests 
         Caption         =   "run tests"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "MalEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" _
    Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, _
    ByVal hMenu As Long, ByVal hInstance As Long, _
    lpParam As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" (ByVal lpLibFileName As String) _
    As Integer
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wp As Long, _
    ByVal lp As Long) As Long
Private Declare Function SendMessageString Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wp As Long, _
    ByVal lp As Any) As Long
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal m As Long, _
    ByVal left As Long, ByVal top As Long, _
    ByVal width As Long, ByVal height As Long, _
    ByVal flags As Long) As Long
Dim sci As Long

Function getSelText() As String
Dim length As Integer: length = SendMessage(sci, SCI_GETSELTEXT, 0, 0)
Dim buf As String: buf = Space(length + 1)
SendMessageString sci, SCI_GETSELTEXT, 0, buf
getSelText = Mid(buf, 1, length - 1)
End Function
Function getContentText() As String
Dim length As Integer: length = SendMessage(sci, SCI_GETTEXT, 0, 0)
Dim buf As String: buf = Space(length + 1)
SendMessageString sci, SCI_GETTEXT, length, buf
getContentText = Mid(buf, 1, length - 1)
End Function
Sub setContentText(s As String)
SendMessageString sci, SCI_SETTEXT, 0, s
End Sub


Function sci_send(msg As Long, a As Long, b As Long) As Long
sci_send = SendMessage(sci, msg, a, b)
End Function
Function sci_getstring(msg As Long, a As Long, buflen As Long) As String
Dim buf As String: buf = Space(buflen + 1)
SendMessageString sci, msg, a, buf
sci_getstring = buf
End Function
Function sci_setstring(msg As Long, a As Long, b As String) As Long
sci_setstring = SendMessageString(sci, msg, a, b)
End Function


Private Sub cmdRunsel_Click()
    Dim script As String: script = "(do " + getSelText() + ")"
    MalForm.append "; script result:  " + MalForm.rep(script)
End Sub

Private Sub cmdRuntests_Click()
    

    Dim lines() As String
    lines = Split(getContentText, vbCrLf)
    MalForm.runTests lines
    
End Sub


Private Sub Form_Activate()
SetFocusEx sci
End Sub

Private Sub Form_Click()
SetFocusEx sci
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbTab) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    LoadLibrary ("SciLexer.DLL")
    sci = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", _
        "TEST", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, _
        hwnd, 0, App.hInstance, 0)
    SendMessage sci, SCI_SETSELBACK, 1, &HFFEFD0
    SendMessageString sci, SCI_ADDTEXT, 16, "Hello non-Python"

    On Error Resume Next
    
    Dir1.Path = GetSetting("MalVB", "GUI", "Last Folder")
    Drive1.Drive = Mid(Dir1.Path, 1, 2)
    txtFilename.Text = GetSetting("MalVB", "GUI", "Last Open File")
    openFile txtFilename.Text
    
End Sub

Private Sub cmdSave_Click()
    StringToFile txtFilename.Text, getContentText
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
SaveSetting "MalVB", "GUI", "Last Folder", Dir1.Path
ChDir Dir1.Path
    
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Sub openFile(name As String)
On Error GoTo x
    cmdSave.Enabled = False
    setContentText FileToString(name)
    txtFilename.Text = name
    SaveSetting "MalVB", "GUI", "Last Open File", name
    cmdSave.Enabled = True
    Exit Sub
x:
    MsgBox "failed to open file" & vbNewLine & Err.Description
End Sub

Private Sub File1_Click()
    openFile File1.Path + "\" + File1.filename
End Sub

Private Sub cmdRunscript_Click()
    Dim script As String: script = "(do " + getContentText + ")"
    MalForm.append "; script result:  " + MalForm.rep(script)
End Sub


Private Sub Form_Resize()
    SetWindowPos sci, 0, 2, 25, ScaleWidth - 180, ScaleHeight - 75, 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Shift = vbShiftMask Then
            cmdRunsel_Click
            KeyCode = 0
        ElseIf Shift = vbCtrlMask Then
            cmdSave_Click
            cmdRunscript_Click
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyS And Shift = vbCtrlMask Then
        cmdSave_Click
    End If
    If KeyCode = vbKeyF5 Then
        cmdRunscript_Click
    End If
    If KeyCode = vbKeyT And Shift = vbCtrlMask Then
        MalForm.Show
        KeyCode = 0
    End If
    If KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        Unload MalForm
        Unload MalEdit
    End If
End Sub
Function indentlines(ByRef s As String, indent As Boolean) As String
    Dim lines() As String, i As Integer
    lines = Split(s, vbNewLine)
    For i = 0 To UBound(lines)
        If indent Then
            lines(i) = vbTab + lines(i)
        ElseIf Mid(lines(i), 1, 1) = vbTab Then
            lines(i) = Mid(lines(i), 2)
        End If
    Next
    indentlines = Join(lines, vbNewLine)
End Function


Private Sub txtTests_KeyPress(KeyAscii As Integer)
    Debug.Print KeyAscii
    
End Sub
