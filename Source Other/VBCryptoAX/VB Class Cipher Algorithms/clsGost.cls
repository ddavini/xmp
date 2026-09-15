VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Visual Basic GOST Implementation
' By: David Midkiff (mznull@earthlink.net)
'
' Standard Gost implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Gost supports 256-bit keys
' and is a stable algorithm developed by the former Soviet Union but there
' are many notable weaknesses and it remains in question if it really is
' much stronger than DES. Recommended for medium risk level solutions.
'
' Information on the algorithm can be found at:
' http://www.jetico.sci.fi/index.htm#/gost.htm

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub addLongs Lib "myDLL" (ByRef number1 As Long, ByVal number2 As Long)

Event Progress(Percent As Long)

Private m_KeyValue As String
Private K(1 To 8) As Long
Private k87(0 To 255) As Byte
Private k65(0 To 255) As Byte
Private k43(0 To 255) As Byte
Private k21(0 To 255) As Byte
Private sBox(0 To 7, 0 To 255) As Byte
Private m_RunningCompiled As Boolean

Private m_bytIndex(0 To 63) As Byte
Private m_bytReverseIndex(0 To 255) As Byte
Private Const k_bytEqualSign As Byte = 61
Private Const k_bytMask1 As Byte = 3
Private Const k_bytMask2 As Byte = 15
Private Const k_bytMask3 As Byte = 63
Private Const k_bytMask4 As Byte = 192
Private Const k_bytMask5 As Byte = 240
Private Const k_bytMask6 As Byte = 252
Private Const k_bytShift2 As Byte = 4
Private Const k_bytShift4 As Byte = 16
Private Const k_bytShift6 As Byte = 64
Private Const k_lMaxBytesPerLine As Long = 152
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private cB64 As New clsBase64

Private Function FileExist(Filename As String) As Boolean
    On Error GoTo errorhandler
    Call FileLen(Filename)
    FileExist = True
    Exit Function
    
errorhandler:
    FileExist = False
End Function
Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long
    Dim x1(0 To 3) As Byte, x2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, value As Long, A As Long
    Call CopyMem(x1(0), Data1, 4)
    Call CopyMem(x2(0), Data2, 4)
    Rest = 0
    For A = 0 To 3
        value = CLng(x1(A)) + CLng(x2(A)) + Rest
        xx(A) = value And 255
        Rest = value \ 256
    Next
    Call CopyMem(UnsignedAdd, xx(0), 4)
End Function
Private Function UnsignedDel(Data1 As Long, Data2 As Long) As Long
    Dim x1(0 To 3) As Byte, x2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, value As Long, A As Long
    Call CopyMem(x1(0), Data1, 4)
    Call CopyMem(x2(0), Data2, 4)
    Call CopyMem(xx(0), UnsignedDel, 4)
    For A = 0 To 3
        value = CLng(x1(A)) - CLng(x2(A)) - Rest
        If (value < 0) Then
            value = value + 256
            Rest = 1
        Else
            Rest = 0
        End If
        xx(A) = value
    Next
    Call CopyMem(UnsignedDel, xx(0), 4)
End Function

Private Static Sub PutWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
    Dim bb(0 To 3) As Byte
    Call CopyMem(bb(0), LongValue, 4)
    CryptBuffer(Offset) = bb(3)
    CryptBuffer(Offset + 1) = bb(2)
    CryptBuffer(Offset + 2) = bb(1)
    CryptBuffer(Offset + 3) = bb(0)
End Sub
Private Static Sub GetWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
    Dim bb(0 To 3) As Byte
    bb(3) = CryptBuffer(Offset)
    bb(2) = CryptBuffer(Offset + 1)
    bb(1) = CryptBuffer(Offset + 2)
    bb(0) = CryptBuffer(Offset + 3)
    Call CopyMem(LongValue, bb(0), 4)
End Sub
Private Static Sub DecryptBlock(LeftWord As Long, RightWord As Long)
    Dim I As Long
    RightWord = RightWord Xor F(LeftWord, K(1))
    LeftWord = LeftWord Xor F(RightWord, K(2))
    RightWord = RightWord Xor F(LeftWord, K(3))
    LeftWord = LeftWord Xor F(RightWord, K(4))
    RightWord = RightWord Xor F(LeftWord, K(5))
    LeftWord = LeftWord Xor F(RightWord, K(6))
    RightWord = RightWord Xor F(LeftWord, K(7))
    LeftWord = LeftWord Xor F(RightWord, K(8))
    For I = 1 To 3
        RightWord = RightWord Xor F(LeftWord, K(8))
        LeftWord = LeftWord Xor F(RightWord, K(7))
        RightWord = RightWord Xor F(LeftWord, K(6))
        LeftWord = LeftWord Xor F(RightWord, K(5))
        RightWord = RightWord Xor F(LeftWord, K(4))
        LeftWord = LeftWord Xor F(RightWord, K(3))
        RightWord = RightWord Xor F(LeftWord, K(2))
        LeftWord = LeftWord Xor F(RightWord, K(1))
    Next
End Sub
Private Static Sub EncryptBlock(LeftWord As Long, RightWord As Long)
    Dim I As Long
    For I = 1 To 3
        RightWord = RightWord Xor F(LeftWord, K(1))
        LeftWord = LeftWord Xor F(RightWord, K(2))
        RightWord = RightWord Xor F(LeftWord, K(3))
        LeftWord = LeftWord Xor F(RightWord, K(4))
        RightWord = RightWord Xor F(LeftWord, K(5))
        LeftWord = LeftWord Xor F(RightWord, K(6))
        RightWord = RightWord Xor F(LeftWord, K(7))
        LeftWord = LeftWord Xor F(RightWord, K(8))
    Next
    RightWord = RightWord Xor F(LeftWord, K(8))
    LeftWord = LeftWord Xor F(RightWord, K(7))
    RightWord = RightWord Xor F(LeftWord, K(6))
    LeftWord = LeftWord Xor F(RightWord, K(5))
    RightWord = RightWord Xor F(LeftWord, K(4))
    LeftWord = LeftWord Xor F(RightWord, K(3))
    RightWord = RightWord Xor F(LeftWord, K(2))
    LeftWord = LeftWord Xor F(RightWord, K(1))
End Sub
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String, Optional OutputIn64 As Boolean) As Boolean
    On Error GoTo errorhandler
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    Call EncryptByte(Buffer(), Key)
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        If OutputIn64 = True Then
            Put #FileO, , cB64.EncodeArray64(Buffer())
        Else
            Put #FileO, , Buffer()
        End If
    Close #FileO
    EncryptFile = True
    Exit Function

errorhandler:
    EncryptFile = False
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String, Optional IsFileIn64 As Boolean) As Boolean
    On Error GoTo errorhandler
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    If IsFileIn64 = True Then Buffer() = cB64.DecodeArray64(StrConv(Buffer(), vbUnicode))
    Call DecryptByte(Buffer(), Key)
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    DecryptFile = True
    Exit Function

errorhandler:
    DecryptFile = False
End Function
Private Static Function F(R As Long, K As Long) As Long
    Dim X As Long, xb(0 To 3) As Byte, xx(0 To 3) As Byte, A As Byte, B As Byte, C As Byte, D As Byte
    If (m_RunningCompiled) Then X = R + K Else X = UnsignedAdd(R, K)
    D = X And &HFF
    X = X \ 256
    C = X And &HFF
    X = X \ 256
    B = X And &HFF
    X = X \ 256
    A = X And &HFF
    xb(0) = k21(A)
    xb(1) = k43(B)
    xb(2) = k65(C)
    xb(3) = k87(D)
    xx(0) = ((xb(3) And 31) * 8) Or ((xb(2) And 224) \ 32)
    xx(1) = ((xb(0) And 31) * 8) Or ((xb(3) And 224) \ 32)
    xx(2) = ((xb(1) And 31) * 8) Or ((xb(0) And 224) \ 32)
    xx(3) = ((xb(2) And 31) * 8) Or ((xb(1) And 224) \ 32)
    Call CopyMem(F, xx(0), 4)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
    Dim byteArray() As Byte
    If IsTextIn64 = True Then Text = cB64.Decode64(Text)
    byteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(byteArray(), Key)
    DecryptString = StrConv(byteArray(), vbUnicode)
End Function
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
    Dim byteArray() As Byte
    byteArray() = StrConv(Text, vbFromUnicode)
    Call EncryptByte(byteArray(), Key)
    EncryptString = StrConv(byteArray(), vbUnicode)
    If OutputIn64 = True Then EncryptString = cB64.Encode64(EncryptString)
End Function
Private Static Function lBSL(ByVal lInput As Long, bShiftBits As Byte) As Long
    lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
    If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then lBSL = (lBSL Or &H80000000)
End Function
Private Static Function lBSR(ByVal lInput As Long, bShiftBits As Byte) As Long
    If bShiftBits = 31 Then
      If lInput < 0 Then lBSR = &HFFFFFFFF Else lBSR = 0
    Else
      lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
    End If
End Function
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
    Dim Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long
    
    If (Len(Key) > 0) Then Me.Key = Key
    OrigLen = UBound(byteArray) + 1
    CipherLen = OrigLen + 12
    If (CipherLen Mod 8 <> 0) Then CipherLen = CipherLen + 8 - (CipherLen Mod 8)
    ReDim Preserve byteArray(CipherLen - 1)
    Call CopyMem(byteArray(12), byteArray(0), OrigLen)
    Call CopyMem(byteArray(8), OrigLen, 4)
    Call Randomize
    Call CopyMem(byteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(byteArray(4), CLng(2147483647 * Rnd), 4)
    For Offset = 0 To (CipherLen - 1) Step 8
        Call GetWord(LeftWord, byteArray(), Offset)
        Call GetWord(RightWord, byteArray(), Offset + 4)
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight
        Call EncryptBlock(LeftWord, RightWord)
        Call PutWord(LeftWord, byteArray(), Offset)
        Call PutWord(RightWord, byteArray(), Offset + 4)
        CipherLeft = LeftWord
        CipherRight = RightWord
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Progress(CurrPercent)
        End If
    Next
    If (CurrPercent <> 100) Then RaiseEvent Progress(100)
End Sub
Public Sub DecryptByte(byteArray() As Byte, Optional Key As String)
    On Error GoTo errorhandler
    Dim Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long
    
    If (Len(Key) > 0) Then Me.Key = Key
    CipherLen = UBound(byteArray) + 1
    For Offset = 0 To (CipherLen - 1) Step 8
        Call GetWord(LeftWord, byteArray(), Offset)
        Call GetWord(RightWord, byteArray(), Offset + 4)
        Call DecryptBlock(RightWord, LeftWord)
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight
        Call GetWord(CipherLeft, byteArray(), Offset)
        Call GetWord(CipherRight, byteArray(), Offset + 4)
        Call PutWord(LeftWord, byteArray(), Offset)
        Call PutWord(RightWord, byteArray(), Offset + 4)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Progress(CurrPercent)
        End If
    Next
    Call CopyMem(OrigLen, byteArray(8), 4)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Gost decryption")
    Call CopyMem(byteArray(0), byteArray(12), OrigLen)
    ReDim Preserve byteArray(OrigLen - 1)
    If (CurrPercent <> 100) Then RaiseEvent Progress(100)

errorhandler:
End Sub
Public Property Let Key(New_Value As String)
    Class_Initialize
    Dim A As Long, Key() As Byte, KeyLen As Long, byteArray() As Byte
    If (m_KeyValue = New_Value) Then Exit Property
    KeyLen = Len(New_Value)
    Key() = StrConv(New_Value, vbFromUnicode)
    ReDim byteArray(0 To 31)
    For A = 0 To 31
        byteArray(A) = Key(A Mod KeyLen)
    Next
    Call CopyMem(K(1), byteArray(0), 32)
    m_KeyValue = New_Value
End Property
Private Sub Class_Initialize()
    Dim A As Long, B As Long, C As Long, LeftWord As Long, S(0 To 7) As Variant
      
    On Local Error Resume Next
    m_RunningCompiled = ((2147483647 + 1) < 0)
      
    S(0) = Array(6, 5, 1, 7, 14, 0, 4, 10, 11, 9, 3, 13, 8, 12, 2, 15)
    S(1) = Array(14, 13, 9, 0, 8, 10, 12, 4, 7, 15, 6, 11, 3, 1, 5, 2)
    S(2) = Array(6, 5, 1, 7, 2, 4, 10, 0, 11, 13, 14, 3, 8, 12, 15, 9)
    S(3) = Array(8, 7, 3, 9, 6, 4, 14, 5, 2, 13, 0, 12, 1, 11, 10, 15)
    S(4) = Array(10, 9, 6, 11, 5, 1, 8, 4, 0, 13, 7, 2, 14, 3, 15, 12)
    S(5) = Array(5, 3, 0, 6, 11, 13, 4, 14, 10, 7, 1, 12, 2, 8, 15, 9)
    S(6) = Array(2, 1, 12, 3, 11, 13, 15, 7, 10, 6, 9, 14, 0, 8, 4, 5)
    S(7) = Array(6, 5, 1, 7, 8, 9, 4, 2, 15, 3, 13, 12, 10, 14, 11, 0)
    
    For A = 0 To 15
        For B = 0 To 7
            sBox(B, A) = S(B)(A)
        Next
    Next
    For A = 0 To 255
        k87(A) = lBSL(CLng(sBox(7, lBSR(A, 4))), 4) Or sBox(6, A And 15)
        k65(A) = lBSL(CLng(sBox(5, lBSR(A, 4))), 4) Or sBox(4, A And 15)
        k43(A) = lBSL(CLng(sBox(3, lBSR(A, 4))), 4) Or sBox(2, A And 15)
        k21(A) = lBSL(CLng(sBox(1, lBSR(A, 4))), 4) Or sBox(0, A And 15)
    Next
End Sub
