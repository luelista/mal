VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public env As LHashmap
Public binds As LList
Public body As Variant
Public is_macro As Boolean
Public name As String

Sub run(params As LList, ByRef result As Variant)
    Dim newenv As New LHashmap
    Set newenv.outer = env
    
    newenv.bind binds, params
    
    eval body, newenv, result
End Sub

