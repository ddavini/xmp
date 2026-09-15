VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClassicCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Classic Ciphers
' By: David Midkiff (mznull@earthlink.net)
'
' This contains a collection of classic ciphers and routines for use
' in basic cryptographic schemes. There is little to no security in
' any of these routines. They are somewhat fun as they bring back
' nostalgic memories of the past.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Dim intBin(7) As Integer

Private Sub Class_Initialize()
intBin(0) = 1
intBin(1) = 2
intBin(2) = 4
intBin(3) = 8
intBin(4) = 16
intBin(5) = 32
intBin(6) = 64
intBin(7) = 128
End Sub

Public Function Ascii2Binary(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iAsc As Integer, hLoop As Integer, tInt As Integer, mLoop As Integer
    Reset
    For mLoop = 1 To Len(Text)
        iAsc = Asc(Mid$(Text, mLoop, 1))
        For hLoop = 7 To 0 Step -1
            If intBin(hLoop) <= iAsc - tInt Then
                Append "1"
                tInt = tInt + intBin(hLoop)
            Else
                Append "0"
            End If
        Next
        tInt = 0
    Next
    Ascii2Binary = Trim(GData)
    Reset
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
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
Private Function HexDecode(Text As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Text) Step 2
        Append Chr$(Val("&H" & Mid$(Text, iCount, 2)))
    Next
    HexDecode = GData
    Reset
End Function
Private Function HexEncode(Text As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Text)
        sTemp = Hex$(Asc(Mid$(Text, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    HexEncode = GData
    Reset
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function Binary2Ascii(BinaryText As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iLoop As Integer, t As Integer, mLoop As Integer, S As String, n As Integer
    For mLoop = 1 To Len(BinaryText) / 8
        For iLoop = ((mLoop - 1) * 8) + 1 To ((mLoop - 1) * 8) + 8
            n = n + 1
            t = t + (Val(Mid(BinaryText, iLoop, 1)) * intBin(8 - n))
        Next
        S = S & Chr(t)
        t = 0: n = 0
    Next
    Binary2Ascii = S
End Function
Public Function ReverseString(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer
    Reset
    For iLoop = Len(Text) To 1 Step -1
        Append Mid$(Text, iLoop, 1)
    Next
    ReverseString = GData
    Reset
End Function
Public Function CaesarShift(Text As String, Difference As Integer) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    While Difference > 26
        Difference = Difference - 26
    Wend
    Reset
    Dim iLoop As Integer, tAsc As String
    For iLoop = 1 To Len(Text)
        tAsc = Mid$(Text, iLoop, 1)
        If tAsc = " " Then Append " "
        If tAsc = LCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 122 Then Append Chr$(97 + ((Asc(tAsc) + Difference) - 122)) Else Append Chr$(Asc(tAsc) + Difference)
        ElseIf tAsc = UCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 90 Then Append Chr$(65 + ((Asc(tAsc) + Difference) - 90)) Else Append Chr$(Asc(tAsc) + Difference)
        End If
    Next
    CaesarShift = GData
    Reset
End Function
Public Function Ascii2Transposition(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    Reset
    dInt = Format$(Len(Text) / 2, 0)
    For iLoop = 1 To Len(Text) / 2
        Append Mid$(Text, iLoop, 1)
        Append Mid$(Text, iLoop + dInt, 1)
    Next
    If Len(Text) / 2 < dInt Then Append Mid$(Text, dInt, 1)
    Ascii2Transposition = GData
    Reset
End Function
Public Function Transposition2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    dInt = Len(Text) / 2
    Reset
    For iLoop = 1 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next
    For iLoop = 2 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next iLoop
    Transposition2Ascii = GData
    Reset
End Function
Public Function Ascii2PigLatin(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        sString = sString & Right$(n(iInt), Len(n(iInt)) - 1) & Left$(n(iInt), 1) & "ay "
    Next
    sString = Trim(sString)
    Ascii2PigLatin = sString
End Function

Public Function PigLatin2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        n(iInt) = Left$(n(iInt), Len(n(iInt)) - 2)
        sString = sString & Right$(n(iInt), 1) & Left$(n(iInt), Len(n(iInt)) - 1) & " "
    Next
    sString = Trim(sString)
    PigLatin2Ascii = sString
End Function

Function Ascii2Random(CustomValue As String, EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim A As String, I As Long, B As String, D As String, Q%, C As String
    Q = Val(Trigger(EncryptCode))
    A = CustomValue
    Reset
    For I = 1 To Len(A)
        DoEvents
        B = Mid$(A, I, 1)
        C = Asc(B) + Q
        If C < 0 Then C = C - C - C
        If Len(C) = 1 Then C = "00" & C
        If Len(C) = 2 Then C = "0" & C
        Append C
    Next I
    D = GData
    Reset
    Dim F As Integer, E As String, G As Integer
    For I = 1 To Len(D)
        DoEvents
        F = Mid$(D, I, 1)
        Randomize
        G = Rnd * 19
        Select Case F
            Case 0: Append Chr(G + 55)
            Case 1: Append Chr(G + 75)
            Case 2: Append Chr(G + 95)
            Case 3: Append Chr(G + 115)
            Case 4: Append Chr(G + 135)
            Case 5: Append Chr(G + 155)
            Case 6: Append Chr(G + 175)
            Case 7: Append Chr(G + 195)
            Case 8: Append Chr(G + 215)
            Case 9: Append Chr(G + 235)
        End Select
    Next I
    Ascii2Random = GData
    Reset
End Function

Function Random2Ascii(ByVal CustomValue As String, ByVal EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim X As String, Y As String, z As String, Q As Integer, A As String, I As Long, B As String, C As String
    X = Left$(EncryptCode, 1)
    Y = Right$(EncryptCode, 1)
    z = Mid$(EncryptCode, Format$(Len(EncryptCode) / 2, "#"), 1)
    Q = Format$((Asc(X) + Asc(Y) + Asc(z)) / 6, "##")
    A = CustomValue
    For I = 1 To Len(A)
        DoEvents
        B = Val(Asc(Mid$(A, I, 1)))
        If B >= 55 And B < 75 Then C = C & 0
        If B >= 75 And B < 95 Then C = C & 1
        If B >= 95 And B < 115 Then C = C & 2
        If B >= 115 And B < 135 Then C = C & 3
        If B >= 135 And B < 155 Then C = C & 4
        If B >= 155 And B < 175 Then C = C & 5
        If B >= 175 And B < 195 Then C = C & 6
        If B >= 195 And B < 215 Then C = C & 7
        If B >= 215 And B < 235 Then C = C & 8
        If B >= 235 And B < 255 Then C = C & 9
        If B < 55 Or B >= 255 Then
            Random2Ascii = ""
            Exit Function
        End If
    Next I
    Reset
    Dim D As Integer, F As Integer
    For I = 1 To Len(C) Step 3
        DoEvents
        D = Mid$(C, I, 3)
        F = D - Q
        If F < 0 Then F = F - F - F
        If F > 255 Or F < 0 Then
            Random2Ascii = ""
            Exit Function
        End If
        Append Chr$(F)
    Next I
    Random2Ascii = GData
    Reset
End Function

Private Function Trigger(CodeWord As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    If CodeWord = "" Then Trigger = "": Exit Function
    Dim X$, Y$, z$, Q%
    X = Left$(CodeWord, 1)
    Y = Right$(CodeWord, 1)
    z = Mid$(CodeWord, Format$(Len(CodeWord) / 2, "#"), 1)
    Q = Format$((Asc(X) + Asc(Y) + Asc(z)) / 6, "##")
    Trigger = Q
End Function
