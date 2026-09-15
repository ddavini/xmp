VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Visual Basic RC4 Implementation
' By: David Midkiff (mznull@earthlink.net)
'
' Standard RC4 implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. RC4 is an extremely fast
' and very secure stream cipher from RSA Data Security, Inc. I recommend
' this for high risk low resource environments. It's speed is very very
' attractive. Patents no longer apply for commercial use.
'
' Information on the algorithm can be found at:
' http://www.rsasecurity.com/rsalabs/faq/3-6-3.html

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Progress(Percent As Long)

Private m_Key As String
Private m_sBox(0 To 255) As Integer
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

Public Sub DecryptByte(byteArray() As Byte, Optional Key As String)
    Call EncryptByte(byteArray(), Key)
End Sub
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
    Dim byteArray() As Byte
    byteArray() = StrConv(Text, vbFromUnicode)
    Call EncryptByte(byteArray(), Key)
    EncryptString = StrConv(byteArray(), vbUnicode)
    If OutputIn64 = True Then EncryptString = cB64.Encode64(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
    Dim byteArray() As Byte
    If IsTextIn64 = True Then Text = cB64.Decode64(Text)
    byteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(byteArray(), Key)
    DecryptString = StrConv(byteArray(), vbUnicode)
End Function
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
Dim I As Long, J As Long, temp As Byte, Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, sBox(0 To 255) As Integer

If (Len(Key) > 0) Then Me.Key = Key
Call CopyMem(sBox(0), m_sBox(0), 512)
OrigLen = UBound(byteArray) + 1
CipherLen = OrigLen

For Offset = 0 To (OrigLen - 1)
    I = (I + 1) Mod 256
    J = (J + sBox(I)) Mod 256
    temp = sBox(I)
    sBox(I) = sBox(J)
    sBox(J) = temp
    byteArray(Offset) = byteArray(Offset) Xor (sBox((sBox(I) + sBox(J)) Mod 256))
    If (Offset >= NextPercent) Then
        CurrPercent = Int((Offset / CipherLen) * 100)
        NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
        RaiseEvent Progress(CurrPercent)
    End If
Next
If (CurrPercent <> 100) Then RaiseEvent Progress(100)
End Sub
Private Function FileExist(Filename As String) As Boolean
    On Error GoTo errorhandler
    Call FileLen(Filename)
    FileExist = True
    Exit Function

errorhandler:
    FileExist = False
End Function
Public Property Let Key(New_Value As String)
    Dim A As Long, B As Long, temp As Byte, Key() As Byte, KeyLen As Long
    If (m_Key = New_Value) Then Exit Property
    m_Key = New_Value
    Key() = StrConv(m_Key, vbFromUnicode)
    KeyLen = Len(m_Key)
    For A = 0 To 255
        m_sBox(A) = A
    Next A
    For A = 0 To 255
        B = (B + m_sBox(A) + Key(A Mod KeyLen)) Mod 256
        temp = m_sBox(A)
        m_sBox(A) = m_sBox(B)
        m_sBox(B) = temp
    Next
End Property

