VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim initialized As Boolean

Dim arr() As Variant
Dim l As Integer

Function GetArray() As Variant()
GetArray = arr
End Function
Sub SetArray(a() As Variant)
arr = a
End Sub

Function length() As Integer
length = l
End Function

Property Get Item(idx As Integer) As Variant
    If Not initialized Or idx < 0 Or idx >= l Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(arr(idx)) Then
        Set Item = arr(idx)
    Else
        Item = arr(idx)
    End If
End Property

Sub retr(idx As Integer, ByRef v As Variant)
    If Not initialized Or idx < 0 Or idx >= l Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(arr(idx)) Then
        Set v = arr(idx)
    Else
        v = arr(idx)
    End If
End Sub

Sub assign(idx As Integer, ByRef v As Variant)
    If Not initialized Then ReDim arr(100): initialized = True
    If idx > UBound(arr) Then ReDim Preserve arr(idx + 100)
    If IsObject(v) Then
        Set arr(idx) = v
    Else
         arr(idx) = v
    End If
    If idx + 1 > l Then l = idx + 1
End Sub

Sub add(ByRef v As Variant)
    assign l, v
End Sub




