Attribute VB_Name = "CmdLineHelper"
Option Explicit
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" _
(ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long

Public Declare Function WriteFile Lib "kernel32" _
(ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&



