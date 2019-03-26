VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim initialized As Boolean

Dim arr() As Variant
Dim L As Long

Public isVector As Boolean

Function GetArray() As Variant()
    If L = 0 Then
        Erase arr
        initialized = False
    Else
        ReDim Preserve arr(L - 1)
    End If
    GetArray = arr
End Function
Function slice(Optional start As Long = 0, Optional leng As Long = -1) As LList
    Dim a As New LList
    a.addall Me, start, leng
    Set slice = a
End Function
Sub addall(ByRef otherlist As LList, Optional startother As Long = 0, Optional leng As Long = -1)
    If leng < 0 Then leng = otherlist.length - startother + leng + 1
    If leng > 0 Then
        expect leng - 1
        Dim i As Long
        For i = 0 To leng - 1
            add otherlist.item(i + startother)
        Next
    End If
End Sub
Sub SetArray(a() As Variant)
arr = a
End Sub

Function length() As Long
length = L
End Function

Property Get item(ByVal idx As Long) As Variant
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(arr(idx)) Then
        Set item = arr(idx)
    Else
        item = arr(idx)
    End If
End Property

Sub retr(idx As Long, ByRef v As Variant)
    If Not initialized Or idx < 0 Or idx >= L Then Err.Raise 1004, "Lisp", "index of of range: " & idx
    If IsObject(arr(idx)) Then
        Set v = arr(idx)
    Else
        v = arr(idx)
    End If
End Sub

Sub assign(ByVal idx As Long, ByRef v As Variant)
    If Not initialized Then init 100
    If idx > UBound(arr) Then ReDim Preserve arr(idx + 100)
    If IsObject(v) Then
        Set arr(idx) = v
    Else
         arr(idx) = v
    End If
    If idx + 1 > L Then L = idx + 1
End Sub

Sub init(capacity As Long)
    ReDim arr(capacity): initialized = True
End Sub
Sub expect(capacity As Long)
    If Not initialized Then init capacity
    If L + capacity > UBound(arr) Then ReDim Preserve arr(L + capacity)
End Sub
Sub add(ByRef v As Variant)
    assign L, v
End Sub

Function indexof(ByRef v As Variant) As Long
    Dim i As Long
    For i = 0 To L - 1
        If compare(arr(i), v) Then
            indexof = i
            Exit Function
        End If
    Next
    indexof = -1
End Function

Sub assertLength(ByVal min As Long, ByVal max As Long, message As String)
If L < min Then Err.Raise 1004, "Lisp", "expected list of min. " & min & " items, got " & L & ": " & message
If max > -1 And L > max Then Err.Raise 1004, "Lisp", "expected list of max. " & max & " items, got " & L & ": " & message
End Sub
