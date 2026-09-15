VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkipjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Visual Basic Skipjack Implementation
' By: David Midkiff (mznull@earthlink.net)
'
' Standard Skipjack implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Skipjack is questionable since
' it was developed by the NSA for use in Clipper and Capstone chips ... ciphers
' from the NSA can suggest backdoors or hidden weaknesses but this appears to
' be a farely decent cipher. I would recommend it for low-risk solutions.
'
' Skipjack is property of the NSA but free for public use.
'
' Information can be found at:
' http://csrc.nist.gov/encryption/tkencryption.html

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Progress(Percent As Long)
Private m_KeyValue As String
Private m_F(0 To 255) As Byte
Private m_Key(0 To 127) As Byte
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
Public Function DecryptByte(byteArray() As Byte, Optional Key As String) As String
    On Error GoTo errorhandler
    Dim I As Long, u As Long, K As Long, temp As Byte, Round As Long, Offset As Long, OrigLen As Long, CipherLen As Long, G(0 To 5) As Byte, NextPercent As Long, CurrPercent As Long, Counter(0 To 32) As Byte, w(0 To 3, 0 To 33) As Integer
    If (Len(Key) > 0) Then Me.Key = Key
    CipherLen = UBound(byteArray) + 1
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = byteArray(Offset)
        byteArray(Offset) = byteArray(Offset + 1)
        byteArray(Offset + 1) = temp
    Next
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(w(0, 32), byteArray(Offset), 8)
        K = 32
        u = 31
        For I = 0 To 32
            Counter(I) = I + 1
        Next
        For Round = 1 To 2
            For I = 1 To 8
                  Call CopyMem(G(4), w(1, K), 2)
                  G(3) = m_F(G(5) Xor m_Key(4 * u + 3)) Xor G(4)
                  G(2) = m_F(G(3) Xor m_Key(4 * u + 2)) Xor G(5)
                  G(0) = m_F(G(2) Xor m_Key(4 * u + 1)) Xor G(3)
                  G(1) = m_F(G(0) Xor m_Key(4 * u)) Xor G(2)
                  Call CopyMem(w(0, K - 1), G(0), 2)
                  w(1, K - 1) = w(0, K - 1) Xor w(2, K) Xor Counter(K - 1)
                  w(2, K - 1) = w(3, K)
                  w(3, K - 1) = w(0, K)
                  u = u - 1
                  K = K - 1
            Next
            For I = 1 To 8
                Call CopyMem(G(4), w(1, K), 2)
                G(3) = m_F(G(5) Xor m_Key(4 * u + 3)) Xor G(4)
                G(2) = m_F(G(3) Xor m_Key(4 * u + 2)) Xor G(5)
                G(0) = m_F(G(2) Xor m_Key(4 * u + 1)) Xor G(3)
                G(1) = m_F(G(0) Xor m_Key(4 * u)) Xor G(2)
                Call CopyMem(w(0, K - 1), G(0), 2)
                w(1, K - 1) = w(2, K)
                w(2, K - 1) = w(3, K)
                w(3, K - 1) = w(0, K) Xor w(1, K) Xor Counter(K - 1)
                u = u - 1
                K = K - 1
            Next
        Next
        w(0, 0) = w(0, 0) Xor w(0, 33)
        w(1, 0) = w(1, 0) Xor w(1, 33)
        w(2, 0) = w(2, 0) Xor w(2, 33)
        w(3, 0) = w(3, 0) Xor w(3, 33)
        Call CopyMem(byteArray(Offset), w(0, 0), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Progress(CurrPercent)
        End If
        Call CopyMem(w(0, 33), w(0, 32), 8)
    Next
    
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = byteArray(Offset)
        byteArray(Offset) = byteArray(Offset + 1)
        byteArray(Offset + 1) = temp
    Next
    Call CopyMem(OrigLen, byteArray(8), 4)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Skipjack decryption")
    Call CopyMem(byteArray(0), byteArray(12), OrigLen)
    ReDim Preserve byteArray(OrigLen - 1)
    If (CurrPercent <> 100) Then RaiseEvent Progress(100)

errorhandler:
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
    Dim byteArray() As Byte
    If IsTextIn64 = True Then Text = cB64.Decode64(Text)
    byteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(byteArray(), Key)
    DecryptString = StrConv(byteArray(), vbUnicode)
End Function
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
    Dim I As Long, K As Long, temp As Byte, Round As Long, Offset As Long, OrigLen As Long, Counter As Long, G(0 To 5) As Byte, CipherLen As Long, NextPercent As Long, CurrPercent As Long, w(0 To 3, 0 To 32) As Integer
    
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
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = byteArray(Offset)
        byteArray(Offset) = byteArray(Offset + 1)
        byteArray(Offset + 1) = temp
    Next
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(w(0, 0), byteArray(Offset), 8)
        w(0, 0) = w(0, 0) Xor w(0, 32)
        w(1, 0) = w(1, 0) Xor w(1, 32)
        w(2, 0) = w(2, 0) Xor w(2, 32)
        w(3, 0) = w(3, 0) Xor w(3, 32)
        K = 0
        Counter = 1
        For Round = 1 To 2
            For I = 1 To 8
                  Call CopyMem(G(0), w(0, K), 2)
                  G(2) = m_F(G(0) Xor m_Key(4 * K)) Xor G(1)
                  G(3) = m_F(G(2) Xor m_Key(4 * K + 1)) Xor G(0)
                  G(5) = m_F(G(3) Xor m_Key(4 * K + 2)) Xor G(2)
                  G(4) = m_F(G(5) Xor m_Key(4 * K + 3)) Xor G(3)
                  Call CopyMem(w(1, K + 1), G(4), 2)
                  w(0, K + 1) = w(1, K + 1) Xor w(3, K) Xor Counter
                  w(2, K + 1) = w(1, K)
                  w(3, K + 1) = w(2, K)
                  Counter = Counter + 1
                  K = K + 1
            Next
            For I = 1 To 8
                Call CopyMem(G(0), w(0, K), 2)
                G(2) = m_F(G(0) Xor m_Key(4 * K)) Xor G(1)
                G(3) = m_F(G(2) Xor m_Key(4 * K + 1)) Xor G(0)
                G(5) = m_F(G(3) Xor m_Key(4 * K + 2)) Xor G(2)
                G(4) = m_F(G(5) Xor m_Key(4 * K + 3)) Xor G(3)
                Call CopyMem(w(1, K + 1), G(4), 2)
                w(0, K + 1) = w(3, K)
                w(2, K + 1) = w(0, K) Xor w(1, K) Xor Counter
                w(3, K + 1) = w(2, K)
                Counter = Counter + 1
                K = K + 1
            Next
        Next
        Call CopyMem(byteArray(Offset), w(0, 32), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Progress(CurrPercent)
        End If
    Next
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = byteArray(Offset)
        byteArray(Offset) = byteArray(Offset + 1)
        byteArray(Offset + 1) = temp
    Next
    If (CurrPercent <> 100) Then RaiseEvent Progress(100)
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
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
    Dim byteArray() As Byte
    byteArray() = StrConv(Text, vbFromUnicode)
    Call EncryptByte(byteArray(), Key)
    EncryptString = StrConv(byteArray(), vbUnicode)
    If OutputIn64 = True Then EncryptString = cB64.Encode64(EncryptString)
End Function
Public Property Let Key(New_Value As String)
    Class_Initialize
    Dim I As Long, pass() As Byte, PassLen As Long
    If (New_Value = m_KeyValue) Then Exit Property
    If (Len(New_Value) = 0) Then Exit Property
    PassLen = Len(New_Value)
    pass() = StrConv(New_Value, vbFromUnicode)
    For I = 0 To 127
        m_Key(I) = pass(I Mod PassLen)
    Next
    m_KeyValue = New_Value
End Property
Private Sub Class_Initialize()
    Dim A As Long, Ftable As Variant
    Ftable = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", "F9", "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", "5F", "66", "0B", "D8", "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", "9B", "76", "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", _
                   "70", "88", "61", "2C", "9F", "0D", "2B", "87", "50", "82", "54", "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", "7F", "AB", "E6", "3E", "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", "8C", "0E", "E2", "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")

    For A = 0 To 255
          m_F(A) = Val("&H" & Ftable(A))
    Next
End Sub
