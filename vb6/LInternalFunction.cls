VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LInternalFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum InternalFunctionEnum
    add
    subtract
    multiply
    divide
    pr_str_
    str_
    list
    list_q
    empty_q
    count_
    eq
    lt
    le
    gt
    ge
    not_
    read_string
    slurp
    
    lif_atom
    lif_atom_q
    lif_deref
    lif_reset_
    lif_swap_
    
    lif_cons
    lif_concat
    
    lif_nth
    lif_slice
    lif_first
    lif_rest
    
    lif_throw
    
    lif_apply
    lif_map
    
    lif_nil_p
    lif_true_p
    lif_false_p
    lif_symbol_p
    
    lif_callbyname
    lif_hostcall
    
    lif_symbol
    lif_keyword
    lif_keyword_p
    lif_vector
    lif_vector_p
    lif_hash_map
    lif_map_p
    lif_assoc
    lif_dissoc
    lif_get
    lif_contains_p
    lif_keys
    lif_vals
    lif_sequential_p
    lif_time_ms
    lif_charcodeat
    lif_str_p
End Enum

Public name As String

Public host As Object
Public hostmethod As String
Public calltype As VbCallType

Public which As InternalFunctionEnum

Sub run(params As LList, ByRef result As Variant)
On Error GoTo intfunError
    Dim i  As Long, ls As LList, atom As LAtom, f As Variant, fun As Variant, funcparams As LList, newhash As LHashmap, hash As LHashmap
    Select Case which
    Case add
        result = params.item(0) + params.item(1)
    Case subtract
        result = params.item(0) - params.item(1)
    Case multiply
        result = params.item(0) * params.item(1)
    Case divide
        result = params.item(0) / params.item(1)
        
    Case pr_str_
        result = print_multi(params, True, " ")
    Case str_
        result = print_multi(params, False, "")
        
    Case list
        Set result = params
    Case list_q
        result = IsList(params.item(0), False)
    Case empty_q
        Set ls = params.item(0)
        result = ls.length = 0
    Case count_
        If VarType(params.item(0)) = vbString Then result = Len(params.item(0)): Exit Sub
        If IsEmpty(params.item(0)) Or params.item(0) Is Nothing Then result = 0: Exit Sub
        Set ls = params.item(0)
        result = ls.length
    Case eq
        result = compare(params.item(0), params.item(1))
    Case lt
        result = params.item(0) < params.item(1)
    Case le
        result = params.item(0) <= params.item(1)
    Case gt
        result = params.item(0) > params.item(1)
    Case ge
        result = params.item(0) >= params.item(1)
    Case not_
        result = IsFalsey(params.item(0))
    Case read_string
        read_str params.item(0), result
    Case slurp:
        result = FileToString(params.item(0))
        
    Case lif_atom:
        Set atom = New LAtom
        params.retr 0, f
        If IsObject(f) Then Set atom.value = f Else atom.value = f
        Set result = atom
    Case lif_atom_q:
        result = False
        If VarType(params.item(0)) = vbObject Then
            If params.item(0) Is Nothing Then Exit Sub
            If TypeOf params.item(0) Is LAtom Then
                result = True
            End If
        End If
    Case lif_deref:
        Set atom = params.item(0)
        atom.retr result
    Case lif_reset_:
        Set atom = params.item(0)
        params.retr 1, result
        If IsObject(result) Then Set atom.value = result Else atom.value = result
    Case lif_swap_:
        params.retr 0, atom
        params.retr 1, f
        Set ls = params.slice(1)
        ls.assign 0, atom.value
        f.run ls, result
        If IsObject(result) Then Set atom.value = result Else atom.value = result
        
        
    Case lif_cons
        Set ls = New LList
        ls.add params.item(0)
        ls.addall params.item(1)
        Set result = ls
    Case lif_concat
        Set ls = New LList
        For i = 0 To params.length - 1
            ls.addall params.item(i)
        Next
        Set result = ls
        
    Case lif_nth
        If VarType(params.item(0)) = vbString Then
            result = Mid(params.item(0), params.item(1) - 1, 1)
        Else
            Set ls = params.item(0)
            ls.retr params.item(1), result
        End If
    Case lif_slice
        If params.length = 3 Then i = params.item(3) Else i = -1
        If VarType(params.item(0)) = vbString Then
            If i < 0 Then i = Len(params.item(0)) + i + 1
            result = Mid(params.item(0), params.item(1) - 1, i)
        Else
            Set ls = params.item(0)
            Set result = ls.slice(params.item(1), i)
        End If
    Case lif_first
        If IsNonemptyList(params.item(0), True) Then
            Set ls = params.item(0)
            ls.retr 0, result
        Else
            Set result = Nothing
        End If
    Case lif_rest
        If params.item(0) Is Nothing Then
            Set result = New LList
        Else
            Set ls = params.item(0)
            Set result = ls.slice(1)
        End If
        
    Case lif_throw
         params.retr 0, lispErrorObject
        Err.Raise 1000, "Lisp", "thrown: " + pr_str(params.item(0), True)
        
    Case lif_apply
        Set fun = params.item(0)
        If params.length = 1 Then
            Set funcparams = New LList
        Else
            Set funcparams = params.slice(1, params.length - 2)
            Set ls = params.item(params.length - 1)
            funcparams.addall ls
        End If
        
        fun.run funcparams, result
        
    Case lif_map
        Set fun = params.item(0)
        Set ls = params.item(1)
        Dim outlist As New LList
        outlist.init ls.length
        
        Set funcparams = New LList
        For i = 0 To ls.length - 1
            funcparams.assign 0, ls.item(i)
            fun.run funcparams, f
            outlist.add f
        Next
        Set result = outlist
        
    Case lif_nil_p
        result = IsNothing(params.item(0))
    Case lif_true_p
        If VarType(params.item(0)) = vbBoolean Then
            result = params.item(0)
        Else
            result = False
        End If
    Case lif_false_p
        If VarType(params.item(0)) = vbBoolean Then
            result = Not params.item(0)
        Else
            result = False
        End If
    Case lif_symbol_p
        result = symbolToString(params.item(0), False) <> ""
        
    Case lif_callbyname
        Dim optionParm As String: optionParm = keywordToString(params.item(0), False)
        i = 0
        calltype = VbMethod
        If optionParm <> "" Then
            i = 1
            Select Case optionParm
            Case "get": calltype = VbGet
            Case "set": calltype = VbSet
            Case "let": calltype = VbLet
            End Select
        End If
        Dim funcname As String: funcname = keywordToString(params.item(i + 1))
        Select Case params.length - i
        Case 2
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype)
        Case 3
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2))
        Case 4
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2), params.item(i + 3))
        Case 5
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2), params.item(i + 3), params.item(i + 4))
        Case 6
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2), params.item(i + 3), params.item(i + 4), params.item(i + 5))
        Case 7
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2), params.item(i + 3), params.item(i + 4), params.item(i + 5), params.item(i + 6))
        Case 8
            variantAssign result, CallByName(params.item(i + 0), funcname, calltype, params.item(i + 2), params.item(i + 3), params.item(i + 4), params.item(i + 5), params.item(i + 6), params.item(i + 7))
        Case Else
            Err.Raise 1012, "Lisp", "CallByName with " & params.length & " args not supported :-("
        End Select
        
    Case lif_hostcall
        variantAssign result, CallByName(host, hostmethod, calltype, params)
        
    Case lif_symbol
        Set result = newSymbol(params.item(0))
    Case lif_keyword
        result = Chr(255) + params.item(0)
    Case lif_keyword_p
        result = IsKeyword(params.item(0))
    Case lif_vector
        Set result = params
        result.isVector = True
    Case lif_vector_p
        If IsValidObject(params.item(0)) Then
            Set ls = params.item(0)
            result = ls.isVector
        End If
    Case lif_hash_map
        Set newhash = New LHashmap
        For i = 0 To params.length - 1 Step 2
            newhash.add params.item(i), params.item(i + 1)
        Next
        Set result = newhash
    Case lif_map_p
        result = IsMap(params.item(0))
    Case lif_assoc
        Set newhash = New LHashmap
        newhash.addall params.item(0)
        For i = 1 To params.length - 1 Step 2
            newhash.setitem params.item(i), params.item(i + 1)
        Next
        Set result = newhash
    Case lif_dissoc
        Set newhash = New LHashmap
        Set hash = params.item(0)
        newhash.enlarge hash.length
        'Set ls = params.item(1)
        For i = 0 To hash.length - 1
            'If ls.indexof(hash.key(i)) = -1 Then
            If params.indexof(hash.key(i)) < 1 Then
                newhash.add hash.key(i), hash.value(i)
            End If
        Next
        Set result = newhash
    Case lif_get
        If Not IsMap(params.item(0)) Then Set result = Nothing: Exit Sub
        Set hash = params.item(0)
        i = hash.find(params.item(1))
        If i = -1 Then Set result = Nothing Else hash.retr i, result
    Case lif_contains_p
        If Not IsMap(params.item(0)) Then Set result = Nothing: Exit Sub
        Set hash = params.item(0)
        i = hash.find(params.item(1))
        result = i > -1
    Case lif_keys
        If Not IsMap(params.item(0)) Then Set result = Nothing: Exit Sub
        Set hash = params.item(0)
        Set ls = New LList
        ls.init hash.length
        For i = 0 To hash.length - 1
            ls.add hash.key(i)
        Next
        Set result = ls
    Case lif_vals
        If Not IsMap(params.item(0)) Then Set result = Nothing: Exit Sub
        Set hash = params.item(0)
        Set ls = New LList
        ls.init hash.length
        For i = 0 To hash.length - 1
            ls.add hash.value(i)
        Next
        Set result = ls
    Case lif_sequential_p
        result = IsList(params.item(0), True)
    Case lif_time_ms
        result = GetTickCount
    Case lif_str_p
        result = VarType(params.item(0)) And Not IsKeyword(params.item(0))
    Case lif_charcodeat
        i = 1
        If params.length = 2 Then i = params.item(1) + 1
        result = Asc(Mid(params.item(0), i, 1))
    End Select
    Exit Sub
intfunError:
    Dim errNo As Long: errNo = Err.Number
    Dim errsource As String: errsource = Err.Source
    Dim errDesc As String: errDesc = Err.Description
    On Error GoTo 0
    Err.Raise errNo, errsource, errDesc & vbNewLine & " during internal function call " & which & " " & pr_str(params, True) ' put user errors through
    
End Sub



