VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim tokens() As String
Public idx As Long
Const WHITESPACE = " ," + vbTab + vbCr + vbLf
Const SPECIAL = "[]{}()'`~^@"


Function getTokens() As String()
    getTokens = tokens
End Function
Function isAtEnd() As Boolean
    isAtEnd = idx > UBound(tokens)
End Function

Sub read_form(ByRef element As Variant)
    Dim tok As String: tok = tokens(idx)
        Dim f, g As Variant
    Select Case Mid(tok, 1, 1)
    Case "("
        idx = idx + 1
        Set element = New LList
        read_list element, ")"
    Case "["
        idx = idx + 1
        Set element = New LList
        element.isVector = True
        read_list element, "]"
    Case "{"
        idx = idx + 1
        Set element = New LHashmap
        read_hash element
    Case ":"
        element = Chr(255) + Mid(tok, 2)
        idx = idx + 1
    Case "'"
        Set element = New LList
        element.add newSymbol("quote")
        idx = idx + 1
        read_form f
        element.add f
    Case "`"
        Set element = New LList
        element.add newSymbol("quasiquote")
        idx = idx + 1
        read_form f
        element.add f
    Case "~"
        Set element = New LList
        element.add newSymbol(IIf(tok = "~@", "splice-unquote", "unquote"))
        idx = idx + 1
        read_form f
        element.add f
    Case "^"
        Set element = New LList
        element.add newSymbol("with-meta")
        idx = idx + 1
        read_form f
        read_form g
        element.add g
        element.add f
    Case "@"
        Set element = New LList
        element.add newSymbol("deref")
        idx = idx + 1
        read_form f
        element.add f
    Case Else
        read_atom element
        idx = idx + 1
    End Select
End Sub

Function isvalidstring(tok As String) As Boolean
    If Len(tok) < 2 Then Exit Function
    isvalidstring = Mid(tok, 1, 1) = """" And Mid(tok, Len(tok)) = """"
End Function
Sub read_atom(ByRef element As Variant)
    Dim tok As String: tok = tokens(idx)
    If IsNumeric(tok) Then
        If InStr(tok, ".") > 0 Then
            element = Val(tok)
        Else
            element = CLng(tok)
        End If
    ElseIf tok = "true" Then
        element = True
    ElseIf tok = "false" Then
        element = False
    ElseIf tok = "nil" Then
        Set element = Nothing
    ElseIf isvalidstring(tok) Then
        element = Replace(Replace(Replace(Mid(tok, 2, Len(tok) - 2), "\""", """"), "\n", vbNewLine), "\\", "\")
    ElseIf Mid(tok, 1, 1) = """" Then
        Err.Raise 1002, "Lisp", "string literal not terminated"
    ElseIf isvalidsymbol(tok) Then
        Set element = newSymbol(tok)
    Else
        Err.Raise 1001, "Lisp", "invalid atom: " + tok
    End If
End Sub

Function read_list(ByRef list As Variant, endmarker As String)
    Do While idx <= UBound(tokens)
        If tokens(idx) = endmarker Then
            idx = idx + 1
            GoTo fini
        End If
        Dim el As Variant
        read_form el
        list.add el
    Loop
    Err.Raise 1002, "Lisp", "unterminated list, expected " & endmarker
fini:
End Function

Function read_hash(ByRef list As Variant)
    Do While idx <= UBound(tokens)
        If tokens(idx) = "}" Then
            idx = idx + 1
            GoTo fini
        End If
        Dim key As Variant
        read_form key
        Dim value As Variant
        read_form value
        list.add key, value
    Loop
    Err.Raise 1002, "Lisp", "unterminated hashmap"
fini:
End Function

Sub tokenize(indata As String)
    ReDim tokens(1000)
    Dim tokenIdx As Long: tokenIdx = 0
    Dim i As Long: i = 1
    Dim strlen As Long: strlen = Len(indata)
    Dim tokenStart As Long
    Dim char As String
    While i <= strlen
        tokenStart = i
        char = Mid(indata, i, 1)
        If InStr(WHITESPACE, char) <> 0 Then
            While i <= strlen
                i = i + 1
                If InStr(WHITESPACE, Mid(indata, i, 1)) = 0 Then
                    GoTo notoken
                End If
            Wend
            GoTo notoken
        End If
        
        If char = "~" And Mid(indata, i + 1, 1) = "@" Then
            i = i + 1
            GoTo addtoken
        End If
        
        If InStr(SPECIAL, char) <> 0 Then
            GoTo addtoken
        End If
        
        If char = ";" Then
            i = InStr(i, indata, vbLf)
            If i = 0 Then i = strlen + 1
            GoTo notoken
        End If
        
        If char = """" Then
            While i <= strlen
                i = i + 1
                char = Mid(indata, i, 1)
                If char = """" Then
                    GoTo addtoken
                ElseIf char = "\" Then
                    i = i + 1
                End If
            Wend
            GoTo addtoken
        End If
        
        While i <= strlen
            If InStr(WHITESPACE + "[]{}('""`,;)", Mid(indata, i + 1, 1)) <> 0 Then
                GoTo addtoken
            End If
            i = i + 1
        Wend
        
addtoken:
        If tokenIdx > UBound(tokens) Then ReDim Preserve tokens(UBound(tokens) + 1000)
        tokens(tokenIdx) = Mid(indata, tokenStart, i - tokenStart + 1)
        tokenIdx = tokenIdx + 1
        i = i + 1
notoken:
    Wend
    ReDim Preserve tokens(tokenIdx - 1)
    idx = 0
End Sub

Function isvalidsymbol(s As String) As Boolean
    isvalidsymbol = True
'    Dim i As Long
'    For i = 1 To Len(s)
'        Select Case Asc(Mid(s, i, 1))
'        Case 65 To 90, 97 To 122
'        Case 48 To 57
'            If i = 1 Then isvalidsymbol = False
'        Case Else
'            isvalidsymbol = False
'        End Select
'    Next
End Function
