VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LHashmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim initialized As Boolean
Public outer As LHashmap


Dim keys() As Variant
Dim values() As Variant
Dim L As Long

Function GetKeys() As Variant()
GetKeys = keys
End Function

Function length() As Long
length = L
End Function

Property Get key(idx As Long) As Variant
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(keys(idx)) Then
        Set key = keys(idx)
    Else
        key = keys(idx)
    End If
End Property
Property Get value(idx As Long) As Variant
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(values(idx)) Then
        Set value = values(idx)
    Else
        value = values(idx)
    End If
End Property
Sub retr(idx As Long, ByRef v As Variant)
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(values(idx)) Then
        Set v = values(idx)
    Else
        v = values(idx)
    End If
End Sub

Sub setvalue(idx As Long, ByRef value As Variant)
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(value) Then
        Set values(idx) = value
    Else
        values(idx) = value
    End If
End Sub

Function find(ByRef key As Variant) As Long
    Dim i As Long
    For i = 0 To L - 1
        If keys(i) = key Then
            find = i
            Exit Function
        End If
    Next
    find = -1
End Function

Sub enlarge(size As Long)
    ReDim Preserve keys(size)
    ReDim Preserve values(size)
End Sub

Function getitem(ByRef key As Variant) As Variant
    Dim index As Long: index = find(key)
    If index = -1 Then
        Set getitem = Nothing
    Else
        retr index, getitem
    End If
End Function
Sub setitem(ByRef key As Variant, ByRef value As Variant)
    Dim index As Long: index = find(key)
    If index = -1 Then
        add key, value
    Else
        setvalue index, value
    End If
End Sub

Sub add(ByRef key As Variant, ByRef value As Variant)
    If Not initialized Then enlarge 100: initialized = True
    
    If L > UBound(keys) Then enlarge UBound(keys) + 100
    
    L = L + 1
    keys(L - 1) = key
    setvalue L - 1, value
End Sub
Sub addall(other As LHashmap)
    Dim i As Long
    For i = 0 To other.length - 1
        add other.key(i), other.value(i)
    Next
End Sub

Sub findrec(ByRef key As Variant, ByRef retrvalue As Variant, Optional doRaise As Boolean = True)
    Dim index As Long: index = find(key)
    If index = -1 Then
        If outer Is Nothing Then
            If doRaise Then
                Err.Raise 1007, "Lisp", "environment entry " + key + " not found"
            Else
                Set retrvalue = Nothing
            End If
        Else
            outer.findrec key, retrvalue, doRaise
        End If
    Else
        retr index, retrvalue
    End If
End Sub

Sub bind(binds As LList, exprs As LList)
    Dim i As Long
    Dim f As Variant
    Dim varName As String
    For i = 0 To binds.length - 1
        varName = symbolToString(binds.item(i))
        If varName = "&" Then
            varName = symbolToString(binds.item(i + 1))
            add varName, exprs.slice(i)
            Exit Sub
        End If
        'exprs.findrec varName, f
        'add varName, f
        add varName, exprs.item(i)
    Next
End Sub


