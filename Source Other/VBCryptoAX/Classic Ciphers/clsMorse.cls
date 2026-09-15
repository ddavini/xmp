VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMorse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Morse Code Encoder/Decoder
' By: David Midkiff (mznull@earthlink.net)

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

Private Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Public Function MorseEncode(Text As String) As String
    Dim xLoop As Long, R As String
    Reset
    For xLoop = 1 To Len(Text)
         R = Mid$(LCase$(Text), xLoop, 1)
        If R = "a" Then
            Append ".- "
        ElseIf R = "b" Then
            Append "-... "
        ElseIf R = "c" Then
            Append "-.-. "
        ElseIf R = "d" Then
            Append "-.. "
        ElseIf R = "e" Then
            Append ". "
        ElseIf R = "f" Then
            Append "..-. "
        ElseIf R = "g" Then
            Append "--. "
        ElseIf R = "h" Then
            Append ".... "
        ElseIf R = "i" Then
            Append ".. "
        ElseIf R = "j" Then
            Append ".--- "
        ElseIf R = "k" Then
            Append "-.- "
        ElseIf R = "l" Then
            Append ".-.. "
        ElseIf R = "m" Then
            Append "-- "
        ElseIf R = "n" Then
            Append "-. "
        ElseIf R = "o" Then
            Append "--- "
        ElseIf R = "p" Then
            Append ".--. "
        ElseIf R = "q" Then
            Append "--.- "
        ElseIf R = "r" Then
            Append ".-. "
        ElseIf R = "s" Then
            Append "... "
        ElseIf R = "t" Then
            Append "- "
        ElseIf R = "u" Then
            Append "..- "
        ElseIf R = "v" Then
            Append "...- "
        ElseIf R = "w" Then
            Append ".-- "
        ElseIf R = "x" Then
            Append "-..- "
        ElseIf R = "y" Then
            Append "-.-- "
        ElseIf R = "z" Then
            Append "--.. "
        ElseIf R = "1" Then
            Append ".---- "
        ElseIf R = "2" Then
            Append "..--- "
        ElseIf R = "3" Then
            Append "...-- "
        ElseIf R = "4" Then
            Append "....- "
        ElseIf R = "5" Then
            Append "..... "
        ElseIf R = "6" Then
            Append "-.... "
        ElseIf R = "7" Then
            Append "--... "
        ElseIf R = "8" Then
            Append "---.. "
        ElseIf R = "9" Then
            Append "----. "
        ElseIf R = "0" Then
            Append "----- "
        ElseIf R = " " Then
            Append "\ "
        ElseIf R = "!" Then
            Append "\ "
        ElseIf R = "@" Then
            Append "\ "
        ElseIf R = "#" Then
            Append "\ "
        ElseIf R = "$" Then
            Append "\ "
        ElseIf R = "%" Then
            Append "\ "
        ElseIf R = "^" Then
            Append "\ "
        ElseIf R = "&" Then
            Append "\ "
        ElseIf R = "*" Then
            Append "\ "
        ElseIf R = "<" Then
            Append "\ "
        ElseIf R = "=" Then
            Append "\ "
        ElseIf R = "+" Then
            Append "\ "
        ElseIf R = "]" Then
            Append "\ "
        ElseIf R = "[" Then
            Append "\ "
        ElseIf R = "{" Then
            Append "\ "
        ElseIf R = "}" Then
            Append "\ "
        ElseIf R = "|" Then
            Append "\ "
        ElseIf R = "." Then
            Append ".-.-.- "
        ElseIf R = "," Then
            Append "--..-- "
        ElseIf R = ":" Then
            Append "---... "
        ElseIf R = "?" Then
            Append "..--.. "
        ElseIf R = "'" Then
            Append ".----. "
        ElseIf R = "-" Then
            Append "-....- "
        ElseIf R = "/" Then
            Append "-..-. "
        ElseIf R = "(" Then
            Append "-.--. "
        ElseIf R = ")" Then
            Append "-.--.- "
        ElseIf R = ";" Then
            Append "-.-.-. "
        ElseIf R = Chr$(34) Then
            Append ".-..-. "
        ElseIf R = "_" Then
            Append "..--.- "
        ElseIf R = Chr$(13) Then
            Append "~ "
        End If
    Next
    MorseEncode = GData
    Reset
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property

Public Function MorseDecode(Text As String) As String
    Dim fLen As Double, X As Double, xStr As String, tStr As String
    If Text <> "" Then
        fLen = Len(Text)
        Do
            X = InStr(Text, " ")
            xStr = Right$(Text, Len(Text) - X)
            tStr = Left$(Text, X)
            If tStr = ".- " Then
                Append "a"
            ElseIf tStr = "-... " Then
                Append "b"
            ElseIf tStr = "-.-. " Then
                Append "c"
            ElseIf tStr = "-.. " Then
                Append "d"
            ElseIf tStr = ". " Then
                Append "e"
            ElseIf tStr = "..-. " Then
                Append "f"
            ElseIf tStr = "--. " Then
                Append "g"
            ElseIf tStr = ".... " Then
                Append "h"
            ElseIf tStr = ".. " Then
                Append "i"
            ElseIf tStr = ".--- " Then
                Append "j"
            ElseIf tStr = "-.- " Then
                Append "k"
            ElseIf tStr = ".-.. " Then
                Append "l"
            ElseIf tStr = "-- " Then
                Append "m"
            ElseIf tStr = "-. " Then
                Append "n"
            ElseIf tStr = "--- " Then
                Append "o"
            ElseIf tStr = ".--. " Then
                Append "p"
            ElseIf tStr = "--.- " Then
                Append "q"
            ElseIf tStr = ".-. " Then
                Append "r"
            ElseIf tStr = "... " Then
                Append "s"
            ElseIf tStr = "- " Then
                Append "t"
            ElseIf tStr = "..- " Then
                Append "u"
            ElseIf tStr = "...- " Then
                Append "v"
            ElseIf tStr = ".-- " Then
                Append "w"
            ElseIf tStr = "-..- " Then
                Append "x"
            ElseIf tStr = "-.-- " Then
                Append "y"
            ElseIf tStr = "--.. " Then
                Append "z"
            ElseIf tStr = ".---- " Then
                Append "1"
            ElseIf tStr = "..--- " Then
                Append "2"
            ElseIf tStr = "...-- " Then
                Append "3"
            ElseIf tStr = "....- " Then
                Append "4"
            ElseIf tStr = "..... " Then
                Append "5"
            ElseIf tStr = "-.... " Then
                Append "6"
            ElseIf tStr = "--... " Then
                Append "7"
            ElseIf tStr = "---.. " Then
                Append "8"
            ElseIf tStr = "----. " Then
                Append "9"
            ElseIf tStr = "----- " Then
                Append "0"
            ElseIf tStr = ".-.-.- " Then
                Append "."
            ElseIf tStr = "--..-- " Then
                Append ","
            ElseIf tStr = "---... " Then
                Append ":"
            ElseIf tStr = "..--.. " Then
                Append "?"
            ElseIf tStr = ".----. " Then
                Append "'"
            ElseIf tStr = "-....- " Then
                Append "-"
            ElseIf tStr = "-..-. " Then
                Append "/"
            ElseIf tStr = "-.-.-. " Then
                Append ";"
            ElseIf tStr = "-.--. " Then
                Append "("
            ElseIf tStr = "-.--.- " Then
                Append ")"
            ElseIf tStr = ".-..-. " Then
                Append Chr$(34)
            ElseIf tStr = "..--.- " Then
                Append "_"
            ElseIf tStr = "\ " Then
                Append " "
            ElseIf tStr = "~ " Then
                Append vbCrLf
            Else
                Append " "
                GoTo bypass
            End If
            Text = xStr
        Loop Until Text = ""
bypass:
    Else
    End If
    MorseDecode = GData
    Reset
End Function
