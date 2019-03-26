Attribute VB_Name = "appMain"
Option Explicit

Sub Main()
Debug.Print GetStdHandle(STD_INPUT_HANDLE)

    If GetStdHandle(STD_INPUT_HANDLE) = 0 Or Command = "-gui" Then
        MalForm.Show
        Exit Sub
    End If
    
    Dim host As New CmdLineHost
    If Command = "" Or Command = """" Then
        
        host.repl
        
    Else
        Dim filename As String: filename = Command
        If Mid(filename, 1, 1) = """" Then
            filename = Mid(filename, 2, Len(filename) - 2)
        End If
        
        host.runfile filename
        
    End If

End Sub
