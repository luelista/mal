VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmdLineHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim env As LHashmap
Dim hStdIn, hStdout, hStdErr As Long

Sub Class_Initialize()
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    hStdout = GetStdHandle(STD_OUTPUT_HANDLE)
    hStdErr = GetStdHandle(STD_ERROR_HANDLE)
    resetEnv
End Sub

Function repl_readline(params As LList) As Variant
    append hStdout, params.item(0)
    Dim s As String: s = readLine
    If s = "" Then
        Set repl_readline = Nothing
    Else
        repl_readline = Mid(s, 1, Len(s) - 2)
    End If
End Function

Function readLine() As String
    Dim inbuf As String, lBytesRead As Long
    inbuf = Space(2000)
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    Dim res As Long: res = ReadFile(hStdIn, inbuf, Len(inbuf) - 1, lBytesRead)
    If res = 0 Then Err.Raise 999, "Lisp host", "ReadLine failed " & Err.LastDllError
    readLine = Mid(inbuf, 1, lBytesRead)
End Function

Sub append(ByVal handle As Long, s As String)
    Dim n As Long, r As Long: n = 0
    s = s
    r = WriteFile(handle, s, Len(s), n)
    Debug.Print s, n, r
End Sub
Sub werr(s As String)
append hStdErr, s & vbNewLine
End Sub
Function repl_eval(params As LList) As Variant
    eval params.item(0), env, repl_eval
End Function

Sub runfile(filename As String)
'werr "runfile " & filename
    rep "(load-file " & pr_str(filename, True) & ")"
End Sub

Function rep(s As String) As String
On Error GoTo ex
    Dim frm, output As Variant
    read_str s, frm
    eval frm, env, output
    rep = pr_str(output, True)

    Exit Function
ex:
    rep = "ERROR:" + Err.Description
    werr rep
End Function

Sub repl()
    Dim s As String
    Dim frm, output As Variant
    rep "(println (str ""Mal ["" *host-language* ""]""))"
    While True
        append hStdout, "> "
        s = readLine
        If s = "" Then Exit Sub
        s = Trim(Mid(s, 1, Len(s) - 2))
        If s = "" Then GoTo continueloop
        
again:
        On Error Resume Next
        
        
        Err.Clear
        read_str s, frm
        
        If Err.Number = 1002 Then 'list/map literal not terminated
            append hStdout, "... "
            s = s + readLine
            GoTo again
        ElseIf Err.Number <> 0 Then
            werr "Parse error: " & Err.Description
            GoTo continueloop
        
        End If
        
        Err.Clear
        eval frm, env, output
        If Err.Number <> 0 Then
            werr "Eval error: " & Err.Description
            GoTo continueloop
        End If
        
        Err.Clear
        append hStdout, pr_str(output, True) & vbNewLine
        If Err.Number <> 0 Then
            werr "Error while printing: " & Err.Description
            GoTo continueloop
        End If
        
continueloop:
    Wend
End Sub


Function repl_prn(params As LList) As Variant
        append hStdout, print_multi(params, True, " ") & vbNewLine
        Set repl_prn = Nothing
End Function
Function repl_println(params As LList) As Variant
        append hStdout, print_multi(params, False, " ") & vbNewLine
        Set repl_println = Nothing
End Function


Function repl_createobject(params As LList) As Variant
    Set repl_createobject = CreateObject(params.item(0))
End Function

Function showTranscript(params As LList) As Variant
    MalForm.Show
    Set showTranscript = Nothing
End Function
Function showEditor(params As LList) As Variant
    MalEdit.Show
    Set showEditor = Nothing
End Function



Sub resetEnv()
    Set env = New LHashmap
    declareInternalFunctions env
    
    
    env.add "prn", newHostFunction(Me, "repl_prn")
    env.add "pr-str", newInternalFunction(pr_str_)
    env.add "println", newHostFunction(Me, "repl_println")
    env.add "eval", newHostFunction(Me, "repl_eval")
   ' env.add "show-editor", newHostFunction(Me, "showEditor")
    'env.add "show-transcript", newHostFunction(Me, "showTranscript")
    env.add "transcript-form", MalForm
    env.add "editor-form", MalEdit
        
    env.add "*host-language*", "Visual Basic 6 (command line)"

    env.add "create-object", newHostFunction(Me, "repl_createobject")
    env.add "readline", newHostFunction(Me, "repl_readline")
    
    'env.add "", newInternalFunction(lif_)
    
    rep internalDefs
    
End Sub

