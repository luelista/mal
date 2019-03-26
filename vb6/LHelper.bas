Attribute VB_Name = "LHelper"
Option Explicit
Public lispErrorObject As Variant

Public Declare Function GetTickCount Lib "kernel32" () As Long

Function FileToString(strFilename As String) As String
  Dim iFile As Integer: iFile = FreeFile
  Open strFilename For Input As #iFile
    If LOF(iFile) > 0 Then FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function
Sub StringToFile(strFilename As String, value As String)
  Dim iFile As Integer: iFile = FreeFile
  Open strFilename For Output As #iFile
    Print #iFile, value
  Close #iFile
End Sub

Sub read_str(s As String, ByRef v As Variant)
    Dim r As New Reader
    r.tokenize s
    r.read_form v
    If Not r.isAtEnd Then Err.Raise 1003, "Lisp", "unused tokens at the end >" & r.getTokens()(r.idx) & "<"
End Sub

Sub variantAssign(ByRef dest As Variant, src As Variant)
    If IsObject(src) Then Set dest = src Else dest = src
End Sub

'''printer
Function print_multi(ByRef params As LList, print_readably As Boolean, joiner As String) As String
    If params.length = 0 Then Exit Function
    Dim out() As String, i As Long
    ReDim out(params.length - 1)
    For i = 0 To UBound(out)
        out(i) = pr_str(params.item(i), print_readably)
    Next
    print_multi = Join(out, joiner)
End Function
Function pr_str(ByVal element As Variant, print_readably As Boolean) As String
    Dim s() As String
    Dim i As Long

    Select Case VarType(element)
    Case vbEmpty
        pr_str = "EMPTY"
    Case vbNull
        pr_str = "NULL"
    Case vbDouble
        pr_str = "" & Replace(FormatNumber(element, 3, , , vbFalse), ",", ".")
    Case vbInteger, vbLong
        pr_str = "" & element
    Case vbBoolean
        pr_str = IIf(element, "true", "false")
    Case vbString
        If IsKeyword(element) Then
            pr_str = ":" + Mid(element, 2) 'keyword
        ElseIf print_readably Then
            pr_str = """" + Replace(Replace(Replace(element, "\", "\\"), vbCrLf, "\n"), """", "\""") + """"
        Else
            pr_str = element
        End If
        
    Case vbObject
        If element Is Nothing Then
            pr_str = "nil"
        ElseIf TypeOf element Is LList Then
            Dim mylist As LList: Set mylist = element
            If mylist.length > 0 Then
                ReDim s(mylist.length - 1)
                For i = 0 To mylist.length - 1
                    s(i) = pr_str(mylist.item(i), print_readably)
                Next
            End If
            If mylist.isVector Then
                pr_str = "[" + Join(s, " ") + "]"
            Else
                pr_str = "(" + Join(s, " ") + ")"
            End If
        ElseIf TypeOf element Is LHashmap Then
            Dim myhash As LHashmap: Set myhash = element
            If myhash.length > 0 Then
                ReDim s(myhash.length - 1)
                For i = 0 To myhash.length - 1
                    s(i) = pr_str(myhash.key(i), print_readably) + " " + pr_str(myhash.value(i), print_readably)
                Next
            End If
            pr_str = "{" + Join(s, " ") + "}"
        ElseIf TypeOf element Is LSymbol Then
            Dim sym As LSymbol: Set sym = element
            pr_str = sym.name
        ElseIf TypeOf element Is LFunction Or TypeOf element Is LInternalFunction Then
            pr_str = "#" + element.name
        ElseIf TypeOf element Is LAtom Then
            pr_str = "@" + pr_str(element.value, True)
        Else
            pr_str = "#???"
        End If
    Case Else
        pr_str = "unknown(" & VarType(element) & ")"
    End Select
End Function



'''helper
Function IsKeyword(s As Variant) As Boolean
    If VarType(s) = vbString Then
        If s <> "" Then
            IsKeyword = Asc(s) = 255
        End If
    End If
End Function
Function keywordToString(s As Variant, Optional doRaise As Boolean = True) As String
    If IsKeyword(s) Then
        keywordToString = Mid(s, 2)
        Exit Function
    End If
    If doRaise Then
        Err.Raise 1009, "Lisp", "keyword expected, got " & pr_str(s, True)
    End If
End Function
Function symbolToString(s As Variant, Optional doRaise As Boolean = True) As String
    If IsValidObject(s) Then
        If TypeOf s Is LSymbol Then
            Dim sym As LSymbol: Set sym = s
            symbolToString = s.name
            Exit Function
        End If
    End If
    If doRaise Then
        Err.Raise 1009, "Lisp", "symbol expected, got " & pr_str(s, True)
    End If
End Function
Function newSymbol(sym As String) As LSymbol
    Dim c As New LSymbol
    c.name = sym
    Set newSymbol = c
End Function
Function newInternalFunction(which As InternalFunctionEnum) As LInternalFunction
    Dim c As New LInternalFunction
    c.which = which
    c.name = "#internal-" & which
    Set newInternalFunction = c
End Function
Function newHostFunction(host As Object, method As String, Optional calltype As VbCallType = VbMethod) As LInternalFunction
    Dim c As New LInternalFunction
    c.which = lif_hostcall
    Set c.host = host
    c.hostmethod = method
    c.calltype = calltype
    c.name = "#host-" + method
    Set newHostFunction = c
End Function

Function newEnvWithOuter(ByRef parent As LHashmap) As LHashmap
    Dim c As New LHashmap
    Set c.outer = parent
    Set newEnvWithOuter = c
End Function
Function IsValidObject(ByRef o As Variant) As Boolean
    If IsObject(o) Then
        If Not o Is Nothing Then
            IsValidObject = True
        End If
    End If
End Function

Function IsMacroCall(ByRef o As Variant, ByRef env As LHashmap, ByRef macroFunction As LFunction) As Boolean
    If IsNonemptyList(o, False) Then
        Dim appl As LList: Set appl = o
        
        Dim name As String: name = symbolToString(appl.item(0), False)
        If name <> "" Then
            Dim foundfun As Variant
            
            env.findrec name, foundfun, False
            If IsValidObject(foundfun) Then
                If TypeOf foundfun Is LFunction Then
                    Set macroFunction = foundfun
                    IsMacroCall = macroFunction.is_macro
                End If
            End If
        End If
    End If
End Function
Function macroexpand(ByRef ast As Variant, ByRef env As LHashmap) As Boolean
    Dim macro As LFunction
    Dim inter As Variant
    While IsMacroCall(ast, env, macro)
        Debug.Print "before macro expand", macro.name, pr_str(ast, True)
        macro.run ast.slice(1), inter
        If IsObject(inter) Then Set ast = inter Else ast = inter
        Debug.Print "after macro expand", pr_str(ast, True)
    Wend
End Function
Function IsMap(ByRef o As Variant) As Boolean
    If IsValidObject(o) Then
        IsMap = TypeOf o Is LHashmap
    End If
End Function
Function IsList(ByRef o As Variant, Optional allowVector As Boolean = False) As Boolean
    If IsValidObject(o) Then
        If TypeOf o Is LList Then
            If Not allowVector Then
                If o.isVector = False Then IsList = True
            Else
                IsList = True
            End If
        End If
    End If
End Function

Function IsListOfLength(ByRef o As Variant, length As Long) As Boolean
    If IsList(o, True) Then
        IsListOfLength = o.length = length
    End If
End Function

Function IsNonemptyList(ByRef o As Variant, Optional allowVector As Boolean = False) As Boolean
    If IsList(o, allowVector) Then
        IsNonemptyList = o.length > 0
    End If
End Function

Function newPair(ByRef a As Variant, ByRef b As Variant) As LList
    Dim L As New LList
    L.init 2
    L.add a
    L.add b
    Set newPair = L
End Function

Function IsFalsey(v As Variant) As Boolean
    If IsObject(v) Then
        If v Is Nothing Then
            IsFalsey = True
        End If
        Exit Function
    End If
    If VarType(v) = vbBoolean And v = False Then
        IsFalsey = True
        Exit Function
    End If
End Function
Function IsNothing(v As Variant) As Boolean
    If IsObject(v) Then
        If v Is Nothing Then
            IsNothing = True
        End If
    End If
End Function
Function compare(a As Variant, b As Variant) As Boolean
    Dim i  As Long
    If VarType(a) <> VarType(b) Then compare = False: Exit Function
    If VarType(a) = vbObject Then
        If a Is b Then compare = True
        If a Is Nothing Then Exit Function
        If b Is Nothing Then Exit Function
        If TypeOf a Is LList And TypeOf b Is LList Then
            Dim list1 As LList: Set list1 = a
            Dim list2 As LList: Set list2 = b
            If list1.length <> list2.length Then compare = False: Exit Function
            For i = 0 To list1.length - 1
                If Not compare(list1.item(i), list2.item(i)) Then compare = False: Exit Function
            Next
            compare = True
        End If
        If TypeOf a Is LHashmap And TypeOf b Is LHashmap Then
            Dim hma As LHashmap: Set hma = a
            Dim hmb As LHashmap: Set hmb = b
            If hma.length <> hmb.length Then compare = False: Exit Function
            For i = 0 To hma.length - 1
                If Not compare(hma.value(i), hmb.getitem(hma.key(i))) Then compare = False: Exit Function
            Next
            compare = True
        End If
        If TypeOf a Is LSymbol And TypeOf b Is LSymbol Then
            compare = symbolToString(a) = symbolToString(b)
        End If
    Else
        If a = b Then compare = True: Exit Function
    End If
End Function

Function buildErrorObjectFromNative() As LHashmap
            Set buildErrorObjectFromNative = New LHashmap
            buildErrorObjectFromNative.add Chr(255) + "number", Err.Number
            buildErrorObjectFromNative.add Chr(255) + "desc", Err.Description
            buildErrorObjectFromNative.add Chr(255) + "source", Err.Source
End Function
Sub getErrorObject(ByRef result As Variant)
    If Err.Number = 1000 And Err.Source = "Lisp" Then
        If VarType(lispErrorObject) = vbObject Then
            Set result = lispErrorObject
        Else
            result = lispErrorObject
        End If
    Else
        Set result = buildErrorObjectFromNative()
    End If
    Set lispErrorObject = Nothing
End Sub

''''eval
Sub eval(ByVal ast As Variant, ByVal env As LHashmap, ByRef output As Variant)
tco:
Dim f As Variant, i  As Long, fun As LFunction
Dim newenv As LHashmap
    Select Case VarType(ast)
    Case vbObject
        If ast Is Nothing Then
            Set output = Nothing
        ElseIf IsNonemptyList(ast, False) Then
            '''MAKROS!
            macroexpand ast, env
            If Not IsList(ast, False) Then
                eval_ast ast, env, output
                Exit Sub
            End If
            
            Dim mylist As LList: Set mylist = ast
            '''wunderfunktionen mit interpretermagie
            Dim func As Variant: mylist.retr 0, func
            If TypeOf func Is LSymbol Then
                Select Case func.name
                Case "def!"
                    mylist.assertLength 3, 3, "def!"
                    eval mylist.item(2), env, output
                    env.setitem symbolToString(mylist.item(1)), output
                    Exit Sub
                Case "defmacro!"
                    mylist.assertLength 3, 3, "defmacro!"
                    eval mylist.item(2), env, output
                    output.is_macro = True: output.name = symbolToString(mylist.item(1))
                    env.setitem symbolToString(mylist.item(1)), output
                    Exit Sub
                Case "let*"
                    mylist.assertLength 3, 3, "let* requires bindlist and body"
                    Set newenv = newEnvWithOuter(env)
                    
                    Dim bindlist As LList
                    mylist.retr 1, bindlist
                    For i = 0 To bindlist.length - 1 Step 2
                        eval bindlist.item(i + 1), newenv, f
                        newenv.add symbolToString(bindlist.item(i)), f
                    Next
                    
                    mylist.retr 2, ast
                    Set env = newenv
                    
                    GoTo tco
                Case "do"
                    For i = 1 To mylist.length - 2
                        eval mylist.item(i), env, output
                    Next
                    mylist.retr mylist.length - 1, ast
                    GoTo tco
                Case "if"
                    mylist.assertLength 3, 4, "if"
                    eval mylist.item(1), env, output
                    If IsFalsey(output) Then
                        If mylist.length < 4 Then
                            Set output = Nothing
                        Else
                            mylist.retr 3, ast
                            GoTo tco
                        End If
                    Else
                        mylist.retr 2, ast
                        GoTo tco
                    End If
                    Exit Sub
                Case "fn*"
                    mylist.assertLength 3, 3, "fn*"
                    Set fun = New LFunction
                    Set fun.env = env
                    Set fun.binds = mylist.item(1)
                    mylist.retr 2, f
                    If IsObject(f) Then Set fun.body = f Else fun.body = f
                    Set output = fun
                    Exit Sub
                Case "quote"
                    mylist.assertLength 2, 2, "quote"
                    mylist.retr 1, output
                    Exit Sub
                Case "quasiquote"
                    mylist.assertLength 2, 2, "quasiquote"
                    mylist.retr 1, f
                    eval_quasiquote f, env, ast
                    Debug.Print pr_str(ast, True)
                    GoTo tco
                Case "macroexpand"
                    mylist.assertLength 2, 2, "macroexpand"
                    Set output = mylist.item(1)
                    macroexpand output, env
                    Exit Sub
                Case "try*"
                    mylist.assertLength 3, 3, "try*"
                    Dim catch As LList: Set catch = mylist.item(2)
                    If symbolToString(catch.item(0)) <> "catch*" Then Err.Raise 1010, "Lisp", "expected catch*"
                    
                    On Error GoTo trycatch
                    eval mylist.item(1), env, output
                    Exit Sub
trycatch:
                    Set newenv = newEnvWithOuter(env)
                    getErrorObject f
                    newenv.add symbolToString(catch.item(1)), f
                    eval catch.item(2), newenv, output
                    Exit Sub
                'Case "catch*"
                '    Set output = mylist
                '    Exit Sub
                End Select
            End If
            
            eval_ast ast, env, output
            Set mylist = output
            mylist.retr 0, func
            If TypeOf func Is LFunction Then
                Set fun = func
                Set env = newEnvWithOuter(fun.env)
                env.bind fun.binds, mylist.slice(1)
                If IsObject(fun.body) Then Set ast = fun.body Else ast = fun.body
                GoTo tco
            Else
                func.run mylist.slice(1), output
            End If
            Exit Sub
        End If
        
        eval_ast ast, env, output
    Case Else
        output = ast
    End Select
End Sub
Sub eval_ast(ast As Variant, env As LHashmap, ByRef output As Variant)
    Dim i As Long
    Dim el As Variant
    Select Case VarType(ast)
    Case vbObject
        If ast Is Nothing Then
            Set output = Nothing
        ElseIf TypeOf ast Is LList Then
            Dim mylist As LList: Set mylist = ast
            Dim newlist As New LList: newlist.isVector = mylist.isVector
            
            For i = 0 To mylist.length - 1
                eval mylist.item(i), env, el
                newlist.add el
            Next
            Set output = newlist
        ElseIf TypeOf ast Is LHashmap Then
            Dim myhash As LHashmap: Set myhash = ast
            Dim newhash As New LHashmap
            For i = 0 To myhash.length - 1
                eval myhash.value(i), env, el
                newhash.add myhash.key(i), el
            Next
            Set output = newhash
        ElseIf TypeOf ast Is LSymbol Then
            Dim sym As LSymbol: Set sym = ast
            env.findrec sym.name, output
        Else
            Set output = ast
        End If
    Case Else
        output = ast
    End Select
End Sub


Sub eval_quasiquote(ast As Variant, env As LHashmap, ByRef output As Variant)
    Dim i As Long
    Dim el As Variant
    Dim astlist As LList, childlist As LList, outlist As LList
    If IsNonemptyList(ast, True) Then
        Set astlist = ast
        If symbolToString(astlist.item(0), False) = "unquote" Then
            ast.retr 1, output
            Exit Sub
        ElseIf IsNonemptyList(astlist.item(0), True) Then
            Set childlist = astlist.item(0)
            If symbolToString(childlist.item(0), False) = "splice-unquote" Then
                Set outlist = New LList
                outlist.init astlist.length + 2
                outlist.add newSymbol("concat")
                outlist.add childlist.item(1)
                eval_quasiquote astlist.slice(1), env, el
                outlist.add el
                Set output = outlist
                Exit Sub
            End If
        End If
        Set outlist = New LList
        outlist.add newSymbol("cons")
        eval_quasiquote astlist.item(0), env, el
        outlist.add el
        eval_quasiquote astlist.slice(1), env, el
        outlist.add el
        Set output = outlist
    Else
        Set output = newPair(newSymbol("quote"), ast)
    End If
End Sub





