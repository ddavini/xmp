Attribute VB_Name = "modConsoleFuntion"
Option Explicit
'Versione 1.2

Public Type COORD
  X As Long
  Y As Long
End Type

Public Type SMALL_RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type PCCONSOLE_SCREEN_BUFFER_INFO
  dwSize As Long
  dwCursorPosition As Long
  wAttributes As Long
  srWindow As SMALL_RECT
  dwMaximumWindowSize As Long
End Type

Public Const FOREGROUND_BLUE = &H1
Public Const FOREGROUND_GREEN = &H2
Public Const FOREGROUND_RED = &H4
Public Const BACKGROUND_BLUE = &H10
Public Const BACKGROUND_GREEN = &H20
Public Const BACKGROUND_RED = &H40
Public Const BACKGROUND_INTENSITY = &H80&
Public Const BACKGROUND_SEARCH = &H20&
Public Const FOREGROUND_INTENSITY = &H8&
Public Const FOREGROUND_SEARCH = (&H10&)

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const INVALID_HANDLE_VALUE = -1&

Private Const ENABLE_LINE_INPUT = &H2&
Private Const ENABLE_ECHO_INPUT = &H4&
Private Const ENABLE_MOUSE_INPUT = &H10&
Private Const ENABLE_PROCESSED_INPUT = &H1&
Private Const ENABLE_WINDOW_INPUT = &H8&
Private Const ENABLE_PROCESSED_OUTPUT = &H1&
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2&

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_SHARE_DELETE = &H4

Private Const CONSOLE_TEXTMODE_BUFFER = &H1

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long

Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, ByRef dwCursorPosition As COORD) As Long

Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByRef ConsoleScreenBufferInfo As PCCONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function CreateConsoleScreenBuffer Lib "kernel32" (dwDesiredAccess As Long, dwShareMode As Long, lpSecurityAttributes As Long, dwFlags As Long, lpScreenBufferData As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function Printf Lib "vbstdio.dll" Alias "vbprintf" (ByVal Str As String) As Long
Public Declare Function Clrscr Lib "vbstdio.dll" Alias "vbclrscr" () As Long


Private Function GetStdOut() As Long
    GetStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
End Function

Private Function GetStdIn() As Long
    GetStdIn = GetStdHandle(STD_INPUT_HANDLE)
End Function

Private Function GetStdErr() As Long
    GetStdErr = GetStdHandle(STD_ERROR_HANDLE)
End Function

Public Function CreateCmdScreenBuffer(Optional DesiredAccess As Long = GENERIC_ALL, _
                                      Optional ShareMode As Long = FILE_SHARE_WRITE) As Long
     CreateCmdScreenBuffer = CreateConsoleScreenBuffer(DesiredAccess, ShareMode, vbNull, CONSOLE_TEXTMODE_BUFFER, vbNull)
     
End Function

Public Function SetCmdCursorPosition(Coordinates As COORD) As Long
    SetCmdCursorPosition = SetConsoleCursorPosition(GetStdOut, Coordinates)
End Function

Public Function SetCmdTitle(ByVal sTitle As String) As Long
    SetCmdTitle = SetConsoleTitle(sTitle)
End Function

Public Function GetCmdTitle() As String
    Dim sTitle As String * 256
    Call GetConsoleTitle(sTitle, 256)
    GetCmdTitle = sTitle
End Function

Public Function AllocCmd() As Long
    AllocCmd = AllocConsole()
End Function

Public Function GetCmdInfo(ByRef Value As PCCONSOLE_SCREEN_BUFFER_INFO) As Long
    GetCmdInfo = GetConsoleScreenBufferInfo(GetStdOut, Value)
End Function

Public Sub sPrintf(Optional ByVal StrOutput As String)
    Dim Written As Long
    StrOutput = StrOutput + vbCrLf
    Call WriteFile(GetStdOut, ByVal StrOutput, Len(StrOutput) + 1, Written)
End Sub

Public Sub WriteSTDIN(ByVal StrToWrite As String)
    Dim Written As Long
    Call WriteFile(GetStdIn, ByVal StrToWrite, Len(StrToWrite), Written, ByVal vbNullChar)
End Sub


Public Sub Writeln(Optional ByVal StrOutput As String)
    Dim Written As Long
    StrOutput = StrOutput + vbCrLf
    Call WriteConsole(GetStdOut, StrOutput, Len(StrOutput), Written, vbNull)
End Sub

Public Sub Readln(ByRef StrInput As String)
    
    If 256 - Len(StrInput) > 0 Then
        StrInput = StrInput & String$(256 - Len(StrInput), vbNullChar)
    Else
        StrInput = Left$(StrInput, 256)
    End If
    
    Call ReadConsole(GetStdIn, StrInput, Len(StrInput), vbNull, vbNull)
    StrInput = Left$(StrInput, InStr(StrInput, Chr$(0)) - 3) + vbCrLf
End Sub

Public Sub Writes(Optional ByVal StrOutput As String)
    Dim Written As Long
    Call WriteConsole(GetStdOut, StrOutput, Len(StrOutput), Written, vbNull)
End Sub

Public Sub Reads(ByRef StrInput As String)

    If 256 - Len(StrInput) > 0 Then
        StrInput = StrInput & String$(256 - Len(StrInput), vbNullChar)
    Else
        StrInput = Left$(StrInput, 256)
    End If
    
    Call ReadConsole(GetStdIn, StrInput, Len(StrInput), vbNull, vbNull)
    StrInput = Left$(StrInput, InStr(StrInput, Chr$(0)) - 3)
End Sub

Public Sub CloseInCmd()
    Call CloseHandle(GetStdIn)
End Sub

Public Sub CloseOutCmd()
    Call CloseHandle(GetStdOut)
End Sub

Public Sub FreeCmd()
    FreeConsole
End Sub

Public Sub ColorCmd(Color As Long)
    SetConsoleTextAttribute GetStdOut, Color
End Sub

Public Function CommandToVector(ByVal StrCommand As String) As String()
Dim Pos As Integer, I As Integer, PosC34 As Integer
Dim StrVector() As String
Dim StrStore As String

On Error GoTo Errore:
    StrCommand = Trim(StrCommand)
    Do
        Pos = InStr(1, StrCommand, " ", vbTextCompare)
        StrStore = Trim$(Left$(StrCommand, Pos))
        If Left$(StrStore, 1) = Chr$(34) Then
            PosC34 = InStr(2, StrCommand, Chr$(34), vbTextCompare)
            StrStore = Left$(StrCommand, PosC34)
            StrCommand = Trim$(Right$(StrCommand, Len(StrCommand) - PosC34))
        Else
            StrCommand = Right$(StrCommand, Len(StrCommand) - Pos)
        End If
        ReDim Preserve StrVector(I)
        StrVector(I) = StrStore
        I = I + 1
    Loop Until Pos = 0
    StrVector(I - 1) = StrCommand
    CommandToVector = StrVector
Exit Function
Errore:
    Call ErrorGST(Err.Description + " " + Err.source)
End Function

Public Function ParameterExist(ByVal Parametro As String, ParamStr() As String, _
                               ByRef Posizione As Integer) As Boolean
Dim I As Integer
On Error GoTo Errore:
  ParameterExist = False
  For I = 0 To UBound(ParamStr)
    If UCase(ParamStr(I)) = Parametro Then
        Posizione = I
        ParameterExist = True
        Exit For
    End If
  Next I
Exit Function
Errore:
    Call ErrorGST(Err.Description + " " + Err.source)
End Function

Public Sub EraseC34(ByRef Str As String)
On Error GoTo Errore:
    If Left$(Str, 1) = Chr$(34) Then
        Str = Right$(Str, Len(Str) - 1)
    End If
    If Right$(Str, 1) = Chr$(34) Then
        Str = Left$(Str, Len(Str) - 1)
    End If
    Exit Sub
Errore:
    Call ErrorGST(Err.Description + " " + Err.source)
End Sub
