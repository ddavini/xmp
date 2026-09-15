VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerpent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' 256-bit Serpent VB Implementation
'
' Credits:
' Serpent algorithm developed by: Eli Biham, Ross Anderson and Lars Knudsen
' Original C Implementation by: Dr Brian Gladman (gladman@seven77.demon.co.uk)
'
' Converted/Ported from C by: Anatoliy Razin
' Optimized/Overhauled for VB by: David Midkiff (mznull@earthlink.net)
'
' This is an extremely slow but working Visual Basic implementation of
' the 256-bit Serpent algorithm. Serpent was a finalist for AES and
' appears to be a very secure powerful algorithm. The C implementation
' is by far the fastest and most efficient. Since Visual Basic is bogged
' down by memory problems and lacks the power that C provides, high level
' speeds for this algorithm may never be achieved in VB. If you find ways
' to increase the speed please contact David and let him know.
'
' I would only recommend using this implementation for academic purposes
' and not in commercial schemes due to its speed. It's inherent security
' is a definite plus in any high risk situations.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Event Progress(Percent As Integer)

#Const BLOCK_REVERSE = True

Private l_key(0 To 139) As MemBlock
Private RegX(0 To 3, 0 To 31) As Byte
Private TheByte(0 To 7) As Byte
Private bStr() As Byte
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

Private Type MemBlock
  Data(0 To 3) As Byte
End Type
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

Private Sub PutBits(ByVal Val As Byte)
    Dim Count As Integer
    For Count = 7 To 1 Step -1
      TheByte(Count) = Val Mod 2
      Val = Val \ 2
    Next
    TheByte(0) = Val Mod 2
End Sub
Private Sub GetBits(Val As Byte)
    Dim Count As Integer
    Val = TheByte(0)
    For Count = 1 To 7
        Val = Val * 2
        Val = Val + TheByte(Count)
    Next
End Sub
Private Sub PutRegByte(ByVal Val As Byte, ByteOrder As Integer, Reg As Integer)
    PutBits Val
    RegX(Reg, ByteOrder * 8) = TheByte(0)
    RegX(Reg, ByteOrder * 8 + 1) = TheByte(1)
    RegX(Reg, ByteOrder * 8 + 2) = TheByte(2)
    RegX(Reg, ByteOrder * 8 + 3) = TheByte(3)
    RegX(Reg, ByteOrder * 8 + 4) = TheByte(4)
    RegX(Reg, ByteOrder * 8 + 5) = TheByte(5)
    RegX(Reg, ByteOrder * 8 + 6) = TheByte(6)
    RegX(Reg, ByteOrder * 8 + 7) = TheByte(7)
End Sub
Private Sub GetRegByte(Val As Byte, ByteOrder As Integer, Reg As Integer)
    TheByte(0) = RegX(Reg, ByteOrder * 8)
    TheByte(1) = RegX(Reg, ByteOrder * 8 + 1)
    TheByte(2) = RegX(Reg, ByteOrder * 8 + 2)
    TheByte(3) = RegX(Reg, ByteOrder * 8 + 3)
    TheByte(4) = RegX(Reg, ByteOrder * 8 + 4)
    TheByte(5) = RegX(Reg, ByteOrder * 8 + 5)
    TheByte(6) = RegX(Reg, ByteOrder * 8 + 6)
    TheByte(7) = RegX(Reg, ByteOrder * 8 + 7)
    GetBits Val
End Sub
Private Sub PutReg(Val As MemBlock, Reg As Integer)
    PutRegByte Val.Data(0), 0, Reg
    PutRegByte Val.Data(1), 1, Reg
    PutRegByte Val.Data(2), 2, Reg
    PutRegByte Val.Data(3), 3, Reg
End Sub
Private Sub GetReg(Val As MemBlock, Reg As Integer)
    GetRegByte Val.Data(0), 0, Reg
    GetRegByte Val.Data(1), 1, Reg
    GetRegByte Val.Data(2), 2, Reg
    GetRegByte Val.Data(3), 3, Reg
End Sub
Private Function Bitwise_OR(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
    Dim Count As Integer, val3 As MemBlock
    PutReg Val1, 0
    PutReg Val2, 1
    RegX(2, 0) = (RegX(0, 0) + RegX(1, 0) + 1) \ 2
    RegX(2, 1) = (RegX(0, 1) + RegX(1, 1) + 1) \ 2
    RegX(2, 2) = (RegX(0, 2) + RegX(1, 2) + 1) \ 2
    RegX(2, 3) = (RegX(0, 3) + RegX(1, 3) + 1) \ 2
    RegX(2, 4) = (RegX(0, 4) + RegX(1, 4) + 1) \ 2
    RegX(2, 5) = (RegX(0, 5) + RegX(1, 5) + 1) \ 2
    RegX(2, 6) = (RegX(0, 6) + RegX(1, 6) + 1) \ 2
    RegX(2, 7) = (RegX(0, 7) + RegX(1, 7) + 1) \ 2
    RegX(2, 8) = (RegX(0, 8) + RegX(1, 8) + 1) \ 2
    RegX(2, 9) = (RegX(0, 9) + RegX(1, 9) + 1) \ 2
    RegX(2, 10) = (RegX(0, 10) + RegX(1, 10) + 1) \ 2
    RegX(2, 11) = (RegX(0, 11) + RegX(1, 11) + 1) \ 2
    RegX(2, 12) = (RegX(0, 12) + RegX(1, 12) + 1) \ 2
    RegX(2, 13) = (RegX(0, 13) + RegX(1, 13) + 1) \ 2
    RegX(2, 14) = (RegX(0, 14) + RegX(1, 14) + 1) \ 2
    RegX(2, 15) = (RegX(0, 15) + RegX(1, 15) + 1) \ 2
    RegX(2, 16) = (RegX(0, 16) + RegX(1, 16) + 1) \ 2
    RegX(2, 17) = (RegX(0, 17) + RegX(1, 17) + 1) \ 2
    RegX(2, 18) = (RegX(0, 18) + RegX(1, 18) + 1) \ 2
    RegX(2, 19) = (RegX(0, 19) + RegX(1, 19) + 1) \ 2
    RegX(2, 20) = (RegX(0, 20) + RegX(1, 20) + 1) \ 2
    RegX(2, 21) = (RegX(0, 21) + RegX(1, 21) + 1) \ 2
    RegX(2, 22) = (RegX(0, 22) + RegX(1, 22) + 1) \ 2
    RegX(2, 23) = (RegX(0, 23) + RegX(1, 23) + 1) \ 2
    RegX(2, 24) = (RegX(0, 24) + RegX(1, 24) + 1) \ 2
    RegX(2, 25) = (RegX(0, 25) + RegX(1, 25) + 1) \ 2
    RegX(2, 26) = (RegX(0, 26) + RegX(1, 26) + 1) \ 2
    RegX(2, 27) = (RegX(0, 27) + RegX(1, 27) + 1) \ 2
    RegX(2, 28) = (RegX(0, 28) + RegX(1, 28) + 1) \ 2
    RegX(2, 29) = (RegX(0, 29) + RegX(1, 29) + 1) \ 2
    RegX(2, 30) = (RegX(0, 30) + RegX(1, 30) + 1) \ 2
    RegX(2, 31) = (RegX(0, 31) + RegX(1, 31) + 1) \ 2
    GetReg val3, 2
    Bitwise_OR = val3
End Function
Private Function Bitwise_AND(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
  Dim Count As Integer, val3 As MemBlock
  
  PutReg Val1, 0
  PutReg Val2, 1
  
  RegX(2, 0) = RegX(0, 0) * RegX(1, 0)
  RegX(2, 1) = RegX(0, 1) * RegX(1, 1)
  RegX(2, 2) = RegX(0, 2) * RegX(1, 2)
  RegX(2, 3) = RegX(0, 3) * RegX(1, 3)
  RegX(2, 4) = RegX(0, 4) * RegX(1, 4)
  RegX(2, 5) = RegX(0, 5) * RegX(1, 5)
  RegX(2, 6) = RegX(0, 6) * RegX(1, 6)
  RegX(2, 7) = RegX(0, 7) * RegX(1, 7)
  RegX(2, 8) = RegX(0, 8) * RegX(1, 8)
  RegX(2, 9) = RegX(0, 9) * RegX(1, 9)
  RegX(2, 10) = RegX(0, 10) * RegX(1, 10)
  RegX(2, 11) = RegX(0, 11) * RegX(1, 11)
  RegX(2, 12) = RegX(0, 12) * RegX(1, 12)
  RegX(2, 13) = RegX(0, 13) * RegX(1, 13)
  RegX(2, 14) = RegX(0, 14) * RegX(1, 14)
  RegX(2, 15) = RegX(0, 15) * RegX(1, 15)
  RegX(2, 16) = RegX(0, 16) * RegX(1, 16)
  RegX(2, 17) = RegX(0, 17) * RegX(1, 17)
  RegX(2, 18) = RegX(0, 18) * RegX(1, 18)
  RegX(2, 19) = RegX(0, 19) * RegX(1, 19)
  RegX(2, 20) = RegX(0, 20) * RegX(1, 20)
  RegX(2, 21) = RegX(0, 21) * RegX(1, 21)
  RegX(2, 22) = RegX(0, 22) * RegX(1, 22)
  RegX(2, 23) = RegX(0, 23) * RegX(1, 23)
  RegX(2, 24) = RegX(0, 24) * RegX(1, 24)
  RegX(2, 25) = RegX(0, 25) * RegX(1, 25)
  RegX(2, 26) = RegX(0, 26) * RegX(1, 26)
  RegX(2, 27) = RegX(0, 27) * RegX(1, 27)
  RegX(2, 28) = RegX(0, 28) * RegX(1, 28)
  RegX(2, 29) = RegX(0, 29) * RegX(1, 29)
  RegX(2, 30) = RegX(0, 30) * RegX(1, 30)
  RegX(2, 31) = RegX(0, 31) * RegX(1, 31)

  GetReg val3, 2
  Bitwise_AND = val3
End Function
Private Function Bitwise_XOR(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
    Dim val3 As MemBlock
    PutReg Val1, 0
    PutReg Val2, 1
    RegX(2, 0) = (RegX(0, 0) + RegX(1, 0)) Mod 2
    RegX(2, 1) = (RegX(0, 1) + RegX(1, 1)) Mod 2
    RegX(2, 2) = (RegX(0, 2) + RegX(1, 2)) Mod 2
    RegX(2, 3) = (RegX(0, 3) + RegX(1, 3)) Mod 2
    RegX(2, 4) = (RegX(0, 4) + RegX(1, 4)) Mod 2
    RegX(2, 5) = (RegX(0, 5) + RegX(1, 5)) Mod 2
    RegX(2, 6) = (RegX(0, 6) + RegX(1, 6)) Mod 2
    RegX(2, 7) = (RegX(0, 7) + RegX(1, 7)) Mod 2
    RegX(2, 8) = (RegX(0, 8) + RegX(1, 8)) Mod 2
    RegX(2, 9) = (RegX(0, 9) + RegX(1, 9)) Mod 2
    RegX(2, 10) = (RegX(0, 10) + RegX(1, 10)) Mod 2
    RegX(2, 11) = (RegX(0, 11) + RegX(1, 11)) Mod 2
    RegX(2, 12) = (RegX(0, 12) + RegX(1, 12)) Mod 2
    RegX(2, 13) = (RegX(0, 13) + RegX(1, 13)) Mod 2
    RegX(2, 14) = (RegX(0, 14) + RegX(1, 14)) Mod 2
    RegX(2, 15) = (RegX(0, 15) + RegX(1, 15)) Mod 2
    RegX(2, 16) = (RegX(0, 16) + RegX(1, 16)) Mod 2
    RegX(2, 17) = (RegX(0, 17) + RegX(1, 17)) Mod 2
    RegX(2, 18) = (RegX(0, 18) + RegX(1, 18)) Mod 2
    RegX(2, 19) = (RegX(0, 19) + RegX(1, 19)) Mod 2
    RegX(2, 20) = (RegX(0, 20) + RegX(1, 20)) Mod 2
    RegX(2, 21) = (RegX(0, 21) + RegX(1, 21)) Mod 2
    RegX(2, 22) = (RegX(0, 22) + RegX(1, 22)) Mod 2
    RegX(2, 23) = (RegX(0, 23) + RegX(1, 23)) Mod 2
    RegX(2, 24) = (RegX(0, 24) + RegX(1, 24)) Mod 2
    RegX(2, 25) = (RegX(0, 25) + RegX(1, 25)) Mod 2
    RegX(2, 26) = (RegX(0, 26) + RegX(1, 26)) Mod 2
    RegX(2, 27) = (RegX(0, 27) + RegX(1, 27)) Mod 2
    RegX(2, 28) = (RegX(0, 28) + RegX(1, 28)) Mod 2
    RegX(2, 29) = (RegX(0, 29) + RegX(1, 29)) Mod 2
    RegX(2, 30) = (RegX(0, 30) + RegX(1, 30)) Mod 2
    RegX(2, 31) = (RegX(0, 31) + RegX(1, 31)) Mod 2
    GetReg val3, 2
    Bitwise_XOR = val3
End Function
Private Function Bitwise_NOT(Val As MemBlock) As MemBlock
    Dim Count As Integer, RetVal As MemBlock
    PutReg Val, 0
    
    RegX(0, 0) = (RegX(0, 0) + 1) Mod 2
    RegX(0, 1) = (RegX(0, 1) + 1) Mod 2
    RegX(0, 2) = (RegX(0, 2) + 1) Mod 2
    RegX(0, 3) = (RegX(0, 3) + 1) Mod 2
    RegX(0, 4) = (RegX(0, 4) + 1) Mod 2
    RegX(0, 5) = (RegX(0, 5) + 1) Mod 2
    RegX(0, 6) = (RegX(0, 6) + 1) Mod 2
    RegX(0, 7) = (RegX(0, 7) + 1) Mod 2
    RegX(0, 8) = (RegX(0, 8) + 1) Mod 2
    RegX(0, 9) = (RegX(0, 9) + 1) Mod 2
    RegX(0, 10) = (RegX(0, 10) + 1) Mod 2
    RegX(0, 11) = (RegX(0, 11) + 1) Mod 2
    RegX(0, 12) = (RegX(0, 12) + 1) Mod 2
    RegX(0, 13) = (RegX(0, 13) + 1) Mod 2
    RegX(0, 14) = (RegX(0, 14) + 1) Mod 2
    RegX(0, 15) = (RegX(0, 15) + 1) Mod 2
    RegX(0, 16) = (RegX(0, 16) + 1) Mod 2
    RegX(0, 17) = (RegX(0, 17) + 1) Mod 2
    RegX(0, 18) = (RegX(0, 18) + 1) Mod 2
    RegX(0, 19) = (RegX(0, 19) + 1) Mod 2
    RegX(0, 20) = (RegX(0, 20) + 1) Mod 2
    RegX(0, 21) = (RegX(0, 21) + 1) Mod 2
    RegX(0, 22) = (RegX(0, 22) + 1) Mod 2
    RegX(0, 23) = (RegX(0, 23) + 1) Mod 2
    RegX(0, 24) = (RegX(0, 24) + 1) Mod 2
    RegX(0, 25) = (RegX(0, 25) + 1) Mod 2
    RegX(0, 26) = (RegX(0, 26) + 1) Mod 2
    RegX(0, 27) = (RegX(0, 27) + 1) Mod 2
    RegX(0, 28) = (RegX(0, 28) + 1) Mod 2
    RegX(0, 29) = (RegX(0, 29) + 1) Mod 2
    RegX(0, 30) = (RegX(0, 30) + 1) Mod 2
    RegX(0, 31) = (RegX(0, 31) + 1) Mod 2
    GetReg RetVal, 0
    Bitwise_NOT = RetVal
End Function
Private Function Bitwise_LShift(Val As MemBlock, Pos As Integer) As MemBlock
    Dim Count As Integer, RetVal As MemBlock
    PutReg Val, 0
    If Pos = 0 Then
        GetReg RetVal, 0
        Bitwise_LShift = RetVal
        Exit Function
    End If
    If Pos >= 32 Or Pos < 0 Then
        SetInt RetVal, 0
        Bitwise_LShift = RetVal
        Exit Function
    End If
    For Count = 0 To 31 - Pos
        RegX(2, Count) = RegX(0, Pos + Count)
    Next
    For Count = 32 - Pos To 31
        RegX(2, Count) = 0
    Next Count
    GetReg RetVal, 2
    Bitwise_LShift = RetVal
End Function
Private Function ulBitwise_LShift(Val As MemBlock, ulPos As MemBlock) As MemBlock
    Dim Pos As Integer
    GetInt ulPos, Pos
    Pos = Pos Mod 32
    ulBitwise_LShift = Bitwise_LShift(Val, Pos)
End Function
Private Function Bitwise_RShift(Val As MemBlock, Pos As Integer) As MemBlock
    Dim Count As Integer, RetVal As MemBlock
    PutReg Val, 0
    If Pos = 0 Then
        GetReg RetVal, 0
        Bitwise_RShift = RetVal
        Exit Function
    End If
    If Pos >= 32 Or Pos < 0 Then
        SetInt RetVal, 0
        Bitwise_RShift = RetVal
        Exit Function
    End If
    For Count = 0 To Pos - 1
        RegX(2, Count) = 0
    Next
    For Count = Pos To 31
        RegX(2, Count) = RegX(0, Count - Pos)
    Next Count
    GetReg RetVal, 2
    Bitwise_RShift = RetVal
End Function
Private Function ulBitwise_RShift(Val As MemBlock, ulPos As MemBlock) As MemBlock
    Dim Pos As Integer
    GetInt ulPos, Pos
    Pos = Pos Mod 32
    ulBitwise_RShift = Bitwise_RShift(Val, Pos)
End Function
Private Function rotl(Val As MemBlock, Pos As Integer) As MemBlock
    rotl = Bitwise_OR(Bitwise_LShift(Val, Pos), Bitwise_RShift(Val, 32 - Pos))
End Function
Private Function rotr(Val As MemBlock, Pos As Integer) As MemBlock
    rotr = Bitwise_OR(Bitwise_RShift(Val, Pos), Bitwise_LShift(Val, 32 - Pos))
End Function
Private Function ul_rotl(Val As MemBlock, ulPos As MemBlock) As MemBlock
    Dim Pos As Integer
    GetInt ulPos, Pos
    Pos = Pos Mod 32
    ul_rotl = Bitwise_OR(Bitwise_LShift(Val, Pos), Bitwise_RShift(Val, 32 - Pos))
End Function
Private Function ul_rotr(Val As MemBlock, ulPos As MemBlock) As MemBlock
    Dim Pos As Integer
    GetInt ulPos, Pos
    Pos = Pos Mod 32
    ul_rotr = Bitwise_OR(Bitwise_RShift(Val, Pos), Bitwise_LShift(Val, 32 - Pos))
End Function
Private Function bswap(Val As MemBlock) As MemBlock
    Dim Temp1 As MemBlock, Temp2 As MemBlock
    SetData Temp1, 0, 255, 0, 255
    SetData Temp2, 255, 0, 255, 0
    bswap = Bitwise_OR(Bitwise_AND(rotl(Val, 8), Temp1), Bitwise_AND(rotr(Val, 8), Temp2))
End Function
Private Function Sum(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
    Dim Count As Integer, val3 As MemBlock
    
    PutReg Val1, 0
    PutReg Val2, 1
    
    Dim NextPos As Integer
    NextPos = 0
    For Count = 31 To 0 Step -1
      RegX(2, Count) = (RegX(0, Count) + RegX(1, Count) + NextPos) Mod 2
      NextPos = (RegX(0, Count) + RegX(1, Count) + NextPos) \ 2
    Next Count
    GetReg val3, 2
    Sum = val3
End Function
Private Function Prod(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
    Dim Count As Integer, SumVal As MemBlock
    SetInt SumVal, 0
    PutReg Val2, 3
    For Count = 31 To 0 Step -1
      If RegX(3, Count) = 1 Then SumVal = Sum(SumVal, Bitwise_LShift(Val1, 31 - Count))
    Next
    Prod = SumVal
End Function
Private Function ChangeSign(Val As MemBlock) As MemBlock
    Dim Count As Integer, RetVal As MemBlock, ToChange As Boolean
    PutReg Val, 0
    ToChange = False
    For Count = 31 To 0 Step -1
        If ToChange Then
            RegX(2, Count) = (RegX(0, Count) + 1) Mod 2
        Else
            RegX(2, Count) = RegX(0, Count)
            If RegX(2, Count) = 1 Then ToChange = True
        End If
    Next
    GetReg RetVal, 2
    ChangeSign = RetVal
End Function
Private Function Substr(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
    Substr = Sum(Val1, ChangeSign(Val2))
End Function
Private Sub SetInt(Dest As MemBlock, IntValue As Byte)
    Dest.Data(0) = 0
    Dest.Data(1) = 0
    Dest.Data(2) = 0
    Dest.Data(3) = IntValue
End Sub
Private Sub SetData(Dest As MemBlock, IVal1 As Byte, IVal2 As Byte, IVal3 As Byte, IVal4 As Byte)
    Dest.Data(0) = IVal1
    Dest.Data(1) = IVal2
    Dest.Data(2) = IVal3
    Dest.Data(3) = IVal4
End Sub
Private Sub GetInt(Src As MemBlock, IntValue As Integer)
    IntValue = Src.Data(3)
End Sub
Private Sub MCopy(Src As MemBlock, Dest As MemBlock)
    PutReg Src, 0
    GetReg Dest, 0
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
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Private Sub SetBuffer(Buffer() As Byte, Str As String, Length As Double)
    Dim CF As Double, StrLength As Double
    StrLength = Len(Str)
    For CF = 0 To Length - 1
        If CF < StrLength Then Buffer(CF) = Asc(Mid$(Str, (CF + 1), 1)) Else Buffer(CF) = 0
    Next
End Sub
Private Sub SetBlock(Blk() As MemBlock, Buffer() As Byte, Pos As Double, Length As Double)
    On Local Error Resume Next
    Dim Count As Double

    For Count = 0 To Length - 1
        Blk(Count \ 4).Data(Count Mod 4) = Buffer(Pos + Count)
    Next
End Sub
Private Sub GetBlock(Blk() As MemBlock, Buffer() As Byte, Pos As Double, Length As Double)
    On Local Error Resume Next
    Dim Count As Integer
    For Count = 0 To Length - 1
        Buffer(Pos + Count) = Blk(Count \ 4).Data(Count Mod 4)
    Next
End Sub
Private Sub GetBuffer(Buffer() As Byte, Str As String, Length As Double)
    Dim Count As Integer
    Str = ""
    Reset
    For Count = 0 To Length - 1
        Append Chr$(Buffer(Count))
    Next Count
    Str = GData
    Reset
End Sub
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
    On Local Error Resume Next
    Dim kblk(0 To 7) As MemBlock, blk1(0 To 3) As MemBlock, blk2(0 To 3) As MemBlock
    Dim StrBuffer() As Byte, KeyBuffer() As Byte, RetVal() As Byte
    Dim Length As Double, Count As Double, PProg As Integer
    Length = Len(Text)
    Length = IIf(Length = 0, 16, 16 * ((Length - 1) \ 16 + 1))
    ReDim StrBuffer(0 To Length - 1)
    ReDim RetVal(0 To Length - 1)
    SetKey Key
    SetBuffer StrBuffer, Text, Length

    For Count = 0 To Length - 1 Step 16
        SetBlock blk1, StrBuffer, Count, 16
        EncryptBlock blk1, blk2
        GetBlock blk2, RetVal, Count, 16
        PProg = (Count / Length) * 100
        If PProg > 100 Or Count = (Length - 16) Then PProg = 100
        RaiseEvent Progress(PProg)
        DoEvents
    Next
    EncryptString = StrConv(RetVal, vbUnicode)
    If OutputIn64 = True Then EncryptString = cB64.Encode64(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
    On Local Error Resume Next
    Dim blk1(0 To 3) As MemBlock, blk2(0 To 3) As MemBlock, Code() As Byte
    Dim StrBuffer() As Byte, RetVal As String, Count As Double, Length As Double, PProg As Integer
    If IsTextIn64 = True Then Text = cB64.Decode64(Text)
    Code() = StrConv(Text, vbFromUnicode)
    Length = UBound(Code) + 1
    ReDim StrBuffer(0 To Length - 1)
    SetKey Key
    For Count = 0 To Length - 1 Step 16
        SetBlock blk1, Code, Count, 16
        DecryptBlock blk1, blk2
        GetBlock blk2, StrBuffer, Count, 16
        PProg = (Count / Length) * 100
        If PProg > 100 Then PProg = 100
        RaiseEvent Progress(PProg)
        DoEvents
    Next
    GetBuffer StrBuffer, RetVal, Length
    DecryptString = RetVal
End Function
Private Sub ib5(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(C): T2 = Bitwise_AND(B, T1): t3 = Bitwise_XOR(D, T2)
    t4 = Bitwise_AND(A, t3): t5 = Bitwise_XOR(B, T1): H = Bitwise_XOR(t4, t5)
    t7 = Bitwise_OR(B, H): t8 = Bitwise_AND(A, t7): F = Bitwise_XOR(t3, t8)
    t10 = Bitwise_OR(A, D): t11 = Bitwise_XOR(T1, t7): E = Bitwise_XOR(t10, t11)
    t13 = Bitwise_XOR(A, C): t14 = Bitwise_AND(B, t10): t15 = Bitwise_OR(t4, t13)
    G = Bitwise_XOR(t14, t15)
End Sub
Private Sub EncryptBlock(in_blk() As MemBlock, out_blk() As MemBlock)
    Dim A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock
    
    #If BLOCK_REVERSE Then
      A = bswap(in_blk(3))
      B = bswap(in_blk(2))
      C = bswap(in_blk(1))
      D = bswap(in_blk(0))
    #Else
      MCopy in_blk(0), A
      MCopy in_blk(1), B
      MCopy in_blk(2), C
      MCopy in_blk(3), D
    #End If
    k_xor 0, A, B, C, D: sb0 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 1, E, F, G, H: sb1 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 2, A, B, C, D: sb2 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 3, E, F, G, H: sb3 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 4, A, B, C, D: sb4 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 5, E, F, G, H: sb5 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 6, A, B, C, D: sb6 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 7, E, F, G, H: sb7 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 8, A, B, C, D: sb0 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 9, E, F, G, H: sb1 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 10, A, B, C, D: sb2 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 11, E, F, G, H: sb3 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 12, A, B, C, D: sb4 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 13, E, F, G, H: sb5 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 14, A, B, C, D: sb6 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 15, E, F, G, H: sb7 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 16, A, B, C, D: sb0 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 17, E, F, G, H: sb1 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 18, A, B, C, D: sb2 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 19, E, F, G, H: sb3 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 20, A, B, C, D: sb4 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 21, E, F, G, H: sb5 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 22, A, B, C, D: sb6 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 23, E, F, G, H: sb7 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 24, A, B, C, D: sb0 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 25, E, F, G, H: sb1 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 26, A, B, C, D: sb2 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 27, E, F, G, H: sb3 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 28, A, B, C, D: sb4 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 29, E, F, G, H: sb5 E, F, G, H, A, B, C, D: rot A, B, C, D
    k_xor 30, A, B, C, D: sb6 A, B, C, D, E, F, G, H: rot E, F, G, H
    k_xor 31, E, F, G, H: sb7 E, F, G, H, A, B, C, D: k_xor 32, A, B, C, D
    #If BLOCK_REVERSE Then
      out_blk(3) = bswap(A)
      out_blk(2) = bswap(B)
      out_blk(1) = bswap(C)
      out_blk(0) = bswap(D)
    #Else
      MCopy A, out_blk(0)
      MCopy B, out_blk(1)
      MCopy C, out_blk(2)
      MCopy D, out_blk(3)
    #End If
End Sub
Private Sub DecryptBlock(in_blk() As MemBlock, out_blk() As MemBlock)
    Dim A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock
    
    #If BLOCK_REVERSE Then
      A = bswap(in_blk(3))
      B = bswap(in_blk(2))
      C = bswap(in_blk(1))
      D = bswap(in_blk(0))
    #Else
      MCopy in_blk(0), A
      MCopy in_blk(1), B
      MCopy in_blk(2), C
      MCopy in_blk(3), D
    #End If
    k_xor 32, A, B, C, D: ib7 A, B, C, D, E, F, G, H: k_xor 31, E, F, G, H
    irot E, F, G, H: ib6 E, F, G, H, A, B, C, D: k_xor 30, A, B, C, D
    irot A, B, C, D: ib5 A, B, C, D, E, F, G, H: k_xor 29, E, F, G, H
    irot E, F, G, H: ib4 E, F, G, H, A, B, C, D: k_xor 28, A, B, C, D
    irot A, B, C, D: ib3 A, B, C, D, E, F, G, H: k_xor 27, E, F, G, H
    irot E, F, G, H: ib2 E, F, G, H, A, B, C, D: k_xor 26, A, B, C, D
    irot A, B, C, D: ib1 A, B, C, D, E, F, G, H: k_xor 25, E, F, G, H
    irot E, F, G, H: ib0 E, F, G, H, A, B, C, D: k_xor 24, A, B, C, D
    irot A, B, C, D: ib7 A, B, C, D, E, F, G, H: k_xor 23, E, F, G, H
    irot E, F, G, H: ib6 E, F, G, H, A, B, C, D: k_xor 22, A, B, C, D
    irot A, B, C, D: ib5 A, B, C, D, E, F, G, H: k_xor 21, E, F, G, H
    irot E, F, G, H: ib4 E, F, G, H, A, B, C, D: k_xor 20, A, B, C, D
    irot A, B, C, D: ib3 A, B, C, D, E, F, G, H: k_xor 19, E, F, G, H
    irot E, F, G, H: ib2 E, F, G, H, A, B, C, D: k_xor 18, A, B, C, D
    irot A, B, C, D: ib1 A, B, C, D, E, F, G, H: k_xor 17, E, F, G, H
    irot E, F, G, H: ib0 E, F, G, H, A, B, C, D: k_xor 16, A, B, C, D
    irot A, B, C, D: ib7 A, B, C, D, E, F, G, H: k_xor 15, E, F, G, H
    irot E, F, G, H: ib6 E, F, G, H, A, B, C, D: k_xor 14, A, B, C, D
    irot A, B, C, D: ib5 A, B, C, D, E, F, G, H: k_xor 13, E, F, G, H
    irot E, F, G, H: ib4 E, F, G, H, A, B, C, D: k_xor 12, A, B, C, D
    irot A, B, C, D: ib3 A, B, C, D, E, F, G, H: k_xor 11, E, F, G, H
    irot E, F, G, H: ib2 E, F, G, H, A, B, C, D: k_xor 10, A, B, C, D
    irot A, B, C, D: ib1 A, B, C, D, E, F, G, H: k_xor 9, E, F, G, H
    irot E, F, G, H: ib0 E, F, G, H, A, B, C, D: k_xor 8, A, B, C, D
    irot A, B, C, D: ib7 A, B, C, D, E, F, G, H: k_xor 7, E, F, G, H
    irot E, F, G, H: ib6 E, F, G, H, A, B, C, D: k_xor 6, A, B, C, D
    irot A, B, C, D: ib5 A, B, C, D, E, F, G, H: k_xor 5, E, F, G, H
    irot E, F, G, H: ib4 E, F, G, H, A, B, C, D: k_xor 4, A, B, C, D
    irot A, B, C, D: ib3 A, B, C, D, E, F, G, H: k_xor 3, E, F, G, H
    irot E, F, G, H: ib2 E, F, G, H, A, B, C, D: k_xor 2, A, B, C, D
    irot A, B, C, D: ib1 A, B, C, D, E, F, G, H: k_xor 1, E, F, G, H
    irot E, F, G, H: ib0 E, F, G, H, A, B, C, D: k_xor 0, A, B, C, D
    #If BLOCK_REVERSE Then
      out_blk(3) = bswap(A)
      out_blk(2) = bswap(B)
      out_blk(1) = bswap(C)
      out_blk(0) = bswap(D)
    #Else
      MCopy A, out_blk(0)
      MCopy B, out_blk(1)
      MCopy C, out_blk(2)
      MCopy D, out_blk(3)
    #End If
End Sub
Private Sub k_get(R As Integer, A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock)
   MCopy A, l_key(4 * R + 8)
   MCopy B, l_key(4 * R + 9)
   MCopy C, l_key(4 * R + 10)
   MCopy D, l_key(4 * R + 11)
End Sub
Private Sub k_set(R As Integer, A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock)
   MCopy l_key(4 * R + 8), A
   MCopy l_key(4 * R + 9), B
   MCopy l_key(4 * R + 10), C
   MCopy l_key(4 * R + 11), D
End Sub
Private Sub irot(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock)
    C = rotr(C, 22)
    A = rotr(A, 5)
    C = Bitwise_XOR(C, Bitwise_XOR(D, Bitwise_LShift(B, 7)))
    A = Bitwise_XOR(A, Bitwise_XOR(B, D))
    D = rotr(D, 7)
    B = rotr(B, 1)
    D = Bitwise_XOR(D, Bitwise_XOR(C, Bitwise_LShift(A, 3)))
    B = Bitwise_XOR(B, Bitwise_XOR(A, C))
    C = rotr(C, 3)
    A = rotr(A, 13)
End Sub
Private Sub ib0(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(A, D): T2 = Bitwise_XOR(C, D): t3 = Bitwise_NOT(T2)
    t4 = Bitwise_OR(A, B): G = Bitwise_XOR(t3, t4): t6 = Bitwise_XOR(B, T1)
    t7 = Bitwise_OR(C, t6): t8 = Bitwise_XOR(A, t7): t9 = Bitwise_AND(T2, t8)
    F = Bitwise_XOR(t6, t9): t11 = Bitwise_NOT(t8): t12 = Bitwise_AND(B, D)
    t13 = Bitwise_OR(F, t12): H = Bitwise_XOR(t11, t13): t15 = Bitwise_XOR(T2, t12)
    t16 = Bitwise_OR(F, H): E = Bitwise_XOR(t15, t16)
End Sub
Private Sub ib1(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(A, D): T2 = Bitwise_AND(A, B): t3 = Bitwise_XOR(B, C)
    t4 = Bitwise_XOR(A, t3): t5 = Bitwise_OR(B, D): H = Bitwise_XOR(t4, t5)
    t7 = Bitwise_OR(C, T1): t8 = Bitwise_XOR(B, t7): t9 = Bitwise_AND(t4, t8)
    F = Bitwise_XOR(T1, t9): t11 = Bitwise_NOT(T2): t12 = Bitwise_AND(H, F)
    t13 = Bitwise_XOR(t9, t11): G = Bitwise_XOR(t12, t13): t15 = Bitwise_AND(A, D)
    t16 = Bitwise_XOR(C, t13): E = Bitwise_XOR(t15, t16)
End Sub
Private Sub ib2(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(B, D): T2 = Bitwise_NOT(T1): t3 = Bitwise_XOR(A, C)
    t4 = Bitwise_XOR(C, T1): t5 = Bitwise_AND(B, t4): E = Bitwise_XOR(t3, t5)
    t7 = Bitwise_OR(A, T2): t8 = Bitwise_XOR(D, t7): t9 = Bitwise_OR(t3, t8)
    H = Bitwise_XOR(T1, t9): t11 = Bitwise_NOT(t4): t12 = Bitwise_OR(E, H)
    F = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(D, t11): t15 = Bitwise_XOR(t3, t12)
    G = Bitwise_XOR(t14, t15)
End Sub
Private Sub ib3(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(B, C): T2 = Bitwise_OR(B, C): t3 = Bitwise_XOR(A, C)
    t4 = Bitwise_XOR(T2, t3): t5 = Bitwise_OR(D, t4): E = Bitwise_XOR(T1, t5)
    t7 = Bitwise_XOR(A, D): t8 = Bitwise_OR(T1, t5): t9 = Bitwise_XOR(T2, t7)
    G = Bitwise_XOR(t8, t9): t11 = Bitwise_AND(A, t4): t12 = Bitwise_OR(E, t9)
    F = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(A, G): t15 = Bitwise_XOR(T2, t14)
    t16 = Bitwise_AND(E, t15): H = Bitwise_XOR(t4, t16)
End Sub
Private Sub ib4(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(C, D): T2 = Bitwise_OR(C, D): t3 = Bitwise_XOR(B, T2)
    t4 = Bitwise_AND(A, t3): F = Bitwise_XOR(T1, t4): t6 = Bitwise_XOR(A, D)
    t7 = Bitwise_OR(B, D): t8 = Bitwise_AND(t6, t7): H = Bitwise_XOR(t3, t8)
    t10 = Bitwise_NOT(A): t11 = Bitwise_XOR(C, H): t12 = Bitwise_OR(t10, t11)
    E = Bitwise_XOR(t3, t12): t14 = Bitwise_OR(C, t4): t15 = Bitwise_XOR(t7, t14)
    t16 = Bitwise_OR(H, t10): G = Bitwise_XOR(t15, t16)
End Sub
Private Sub rot(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock)
    A = rotl(A, 13)
    C = rotl(C, 3)
    D = Bitwise_XOR(D, Bitwise_XOR(C, Bitwise_LShift(A, 3)))
    B = Bitwise_XOR(B, Bitwise_XOR(A, C))
    D = rotl(D, 7)
    B = rotl(B, 1)
    A = Bitwise_XOR(A, Bitwise_XOR(B, D))
    C = Bitwise_XOR(C, Bitwise_XOR(D, Bitwise_LShift(B, 7)))
    A = rotl(A, 5)
    C = rotl(C, 22)
End Sub
Private Sub sb0(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(B, D): T2 = Bitwise_NOT(T1): t3 = Bitwise_OR(A, D)
    t4 = Bitwise_XOR(B, C): H = Bitwise_XOR(t3, t4): t6 = Bitwise_XOR(A, B)
    t7 = Bitwise_OR(A, t4): t8 = Bitwise_AND(C, t6): t9 = Bitwise_OR(T2, t8)
    E = Bitwise_XOR(t7, t9): t11 = Bitwise_XOR(A, H): t12 = Bitwise_AND(T1, t6)
    t13 = Bitwise_XOR(E, t11): F = Bitwise_XOR(t12, t13): t15 = Bitwise_OR(E, F)
    t16 = Bitwise_AND(t3, t15): G = Bitwise_XOR(B, t16)
End Sub
Private Sub sb1(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(A, D): T2 = Bitwise_XOR(B, D): t3 = Bitwise_AND(A, B)
    t4 = Bitwise_NOT(C): t5 = Bitwise_XOR(T2, t3): G = Bitwise_XOR(t4, t5)
    t7 = Bitwise_XOR(A, T2): t8 = Bitwise_OR(B, t4): t9 = Bitwise_OR(D, G)
    t10 = Bitwise_AND(t7, t9): F = Bitwise_XOR(t8, t10): t12 = Bitwise_XOR(C, D)
    t13 = Bitwise_OR(T1, T2): t14 = Bitwise_XOR(F, t12): H = Bitwise_XOR(t13, t14)
    t16 = Bitwise_OR(T1, G): t17 = Bitwise_XOR(t8, t14): E = Bitwise_XOR(t16, t17)
End Sub
Private Sub sb2(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(A): T2 = Bitwise_XOR(B, D): t3 = Bitwise_AND(C, T1)
    E = Bitwise_XOR(T2, t3): t5 = Bitwise_XOR(C, T1): t6 = Bitwise_XOR(C, E)
    t7 = Bitwise_AND(B, t6): H = Bitwise_XOR(t5, t7): t9 = Bitwise_OR(D, t7)
    t10 = Bitwise_OR(E, t5): t11 = Bitwise_AND(t9, t10): G = Bitwise_XOR(A, t11)
    t13 = Bitwise_OR(D, T1): t14 = Bitwise_XOR(T2, H): t15 = Bitwise_XOR(G, t13)
    F = Bitwise_XOR(t14, t15)
End Sub
Private Sub sb3(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(A, C): T2 = Bitwise_OR(A, D): t3 = Bitwise_AND(A, B)
    t4 = Bitwise_AND(A, D): t5 = Bitwise_OR(B, t4): t6 = Bitwise_AND(T1, T2)
    F = Bitwise_XOR(t5, t6): t8 = Bitwise_XOR(B, D): t9 = Bitwise_OR(C, t3)
    t10 = Bitwise_XOR(t6, t8): H = Bitwise_XOR(t9, t10): t12 = Bitwise_XOR(C, t3)
    t13 = Bitwise_AND(T2, H): G = Bitwise_XOR(t12, t13): t15 = Bitwise_NOT(G)
    t16 = Bitwise_XOR(T2, t3): t17 = Bitwise_AND(F, t15): E = Bitwise_XOR(t16, t17)
End Sub
Private Sub sb4(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(A): T2 = Bitwise_XOR(A, D): t3 = Bitwise_XOR(A, B)
    t4 = Bitwise_XOR(C, T1): t5 = Bitwise_OR(T2, t3): E = Bitwise_XOR(t4, t5)
    t7 = Bitwise_NOT(E): t8 = Bitwise_OR(B, t7): H = Bitwise_XOR(T2, t8)
    t10 = Bitwise_AND(A, E): t11 = Bitwise_XOR(B, H): t12 = Bitwise_AND(t8, t11)
    G = Bitwise_XOR(t10, t12): t14 = Bitwise_OR(A, t7): t15 = Bitwise_XOR(t3, t14)
    t16 = Bitwise_AND(H, G): F = Bitwise_XOR(t15, t16)
End Sub
Private Sub sb5(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(A): T2 = Bitwise_XOR(A, B): t3 = Bitwise_XOR(A, D)
    t4 = Bitwise_XOR(C, T1): t5 = Bitwise_OR(T2, t3): E = Bitwise_XOR(t4, t5)
    t7 = Bitwise_NOT(D): t8 = Bitwise_AND(E, t7): F = Bitwise_XOR(T2, t8)
    t10 = Bitwise_OR(B, F): t11 = Bitwise_OR(C, E): t12 = Bitwise_XOR(t7, t10)
    H = Bitwise_XOR(t11, t12): t14 = Bitwise_OR(D, F): t15 = Bitwise_XOR(T1, t14)
    t16 = Bitwise_OR(E, H): G = Bitwise_XOR(t15, t16)
End Sub
Private Sub sb7(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(C): T2 = Bitwise_XOR(B, C): t3 = Bitwise_OR(B, T1)
    t4 = Bitwise_XOR(D, t3): t5 = Bitwise_AND(A, t4): H = Bitwise_XOR(T2, t5)
    t7 = Bitwise_XOR(A, D): t8 = Bitwise_XOR(B, t5): t9 = Bitwise_OR(T2, t8)
    F = Bitwise_XOR(t7, t9): t11 = Bitwise_AND(D, t3): t12 = Bitwise_XOR(t5, F)
    t13 = Bitwise_AND(H, t12): G = Bitwise_XOR(t11, t13): t15 = Bitwise_OR(T1, t4)
    t16 = Bitwise_XOR(t12, G): E = Bitwise_XOR(t15, t16)
End Sub
Private Sub ib6(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_NOT(C): T2 = Bitwise_XOR(A, C): t3 = Bitwise_XOR(B, D)
    t4 = Bitwise_OR(A, T1): F = Bitwise_XOR(t3, t4): t6 = Bitwise_OR(A, B)
    t7 = Bitwise_AND(B, T2): t8 = Bitwise_XOR(F, t6): t9 = Bitwise_OR(t7, t8)
    E = Bitwise_XOR(C, t9): t11 = Bitwise_NOT(F): t12 = Bitwise_OR(D, T2)
    t13 = Bitwise_XOR(t9, t11): H = Bitwise_XOR(t12, t13): t15 = Bitwise_XOR(B, t11)
    t16 = Bitwise_AND(E, H): G = Bitwise_XOR(t15, t16)
End Sub
Private Sub sb6(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_XOR(A, C): T2 = Bitwise_OR(B, D): t3 = Bitwise_XOR(B, C)
    t4 = Bitwise_NOT(t3): t5 = Bitwise_AND(A, D): F = Bitwise_XOR(t4, t5)
    t7 = Bitwise_OR(B, C): t8 = Bitwise_XOR(D, T1): t9 = Bitwise_AND(t7, t8)
    H = Bitwise_XOR(T2, t9): t11 = Bitwise_AND(T1, t7): t12 = Bitwise_XOR(t4, t8)
    t13 = Bitwise_AND(H, t11): E = Bitwise_XOR(t12, t13): t15 = Bitwise_XOR(t3, t11)
    t16 = Bitwise_OR(H, t15): G = Bitwise_XOR(t12, t16)
End Sub
Private Sub ib7(A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock)
    Dim T1 As MemBlock, T2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
    Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
    Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
    Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
    T1 = Bitwise_AND(A, B): T2 = Bitwise_OR(A, B): t3 = Bitwise_OR(C, T1)
    t4 = Bitwise_AND(D, T2): H = Bitwise_XOR(t3, t4): t6 = Bitwise_NOT(D)
    t7 = Bitwise_XOR(B, t4): t8 = Bitwise_XOR(H, t6): t9 = Bitwise_OR(t7, t8)
    F = Bitwise_XOR(A, t9): t11 = Bitwise_XOR(C, t7): t12 = Bitwise_OR(D, F)
    E = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(A, H): t15 = Bitwise_XOR(t3, F)
    t16 = Bitwise_XOR(E, t14): G = Bitwise_XOR(t15, t16)
End Sub
Private Sub k_xor(R As Integer, A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock)
   A = Bitwise_XOR(A, l_key(4 * R + 8))
   B = Bitwise_XOR(B, l_key(4 * R + 9))
   C = Bitwise_XOR(C, l_key(4 * R + 10))
   D = Bitwise_XOR(D, l_key(4 * R + 11))
End Sub
Private Function set_key(in_key() As MemBlock, key_len As Integer) As MemBlock()
    Dim I As Integer, lk As Integer, A As MemBlock, B As MemBlock, C As MemBlock, D As MemBlock, E As MemBlock, F As MemBlock, G As MemBlock, H As MemBlock, RetVal(0 To 0) As MemBlock
    Dim lk1 As MemBlock, One As MemBlock, Temp1 As MemBlock, Temp2 As MemBlock

    If key_len < 0 Or key_len > 256 Then
        SetInt RetVal(0), 0
        set_key = RetVal
        Exit Function
    End If

    I = 0
    lk = (key_len + 31) \ 32
      
    While I < lk
        #If BLOCK_REVERSE Then
            l_key(I) = bswap(in_key(lk - I - 1))
        #Else
            l_key(I) = in_key(I)
        #End If
        I = I + 1
    Wend
    If key_len < 256 Then
        While I < 8
            SetInt l_key(I), 0
            I = I + 1
        Wend
        I = key_len \ 32
        SetInt One, 1
        lk1 = Bitwise_LShift(One, key_len Mod 32)
        l_key(I) = Bitwise_OR(Bitwise_AND(l_key(I), Substr(lk1, One)), lk1)
    End If
    For I = 0 To 131
        lk1 = Bitwise_XOR(l_key(I), l_key(I + 3))
        lk1 = Bitwise_XOR(lk1, l_key(I + 5))
        lk1 = Bitwise_XOR(lk1, l_key(I + 7))
        SetData Temp1, 158, 55, 121, 185
        SetInt Temp2, CByte(I)
        lk1 = Bitwise_XOR(lk1, Temp1)
        lk1 = Bitwise_XOR(lk1, Temp2)
        l_key(I + 8) = Bitwise_OR(Bitwise_LShift(lk1, 11), Bitwise_RShift(lk1, 21))
    Next
    
    k_set 0, A, B, C, D: sb3 A, B, C, D, E, F, G, H: k_get 0, E, F, G, H
    k_set 1, A, B, C, D: sb2 A, B, C, D, E, F, G, H: k_get 1, E, F, G, H
    k_set 2, A, B, C, D: sb1 A, B, C, D, E, F, G, H: k_get 2, E, F, G, H
    k_set 3, A, B, C, D: sb0 A, B, C, D, E, F, G, H: k_get 3, E, F, G, H
    k_set 4, A, B, C, D: sb7 A, B, C, D, E, F, G, H: k_get 4, E, F, G, H
    k_set 5, A, B, C, D: sb6 A, B, C, D, E, F, G, H: k_get 5, E, F, G, H
    k_set 6, A, B, C, D: sb5 A, B, C, D, E, F, G, H: k_get 6, E, F, G, H
    k_set 7, A, B, C, D: sb4 A, B, C, D, E, F, G, H: k_get 7, E, F, G, H
    k_set 8, A, B, C, D: sb3 A, B, C, D, E, F, G, H: k_get 8, E, F, G, H
    k_set 9, A, B, C, D: sb2 A, B, C, D, E, F, G, H: k_get 9, E, F, G, H
    k_set 10, A, B, C, D: sb1 A, B, C, D, E, F, G, H: k_get 10, E, F, G, H
    k_set 11, A, B, C, D: sb0 A, B, C, D, E, F, G, H: k_get 11, E, F, G, H
    k_set 12, A, B, C, D: sb7 A, B, C, D, E, F, G, H: k_get 12, E, F, G, H
    k_set 13, A, B, C, D: sb6 A, B, C, D, E, F, G, H: k_get 13, E, F, G, H
    k_set 14, A, B, C, D: sb5 A, B, C, D, E, F, G, H: k_get 14, E, F, G, H
    k_set 15, A, B, C, D: sb4 A, B, C, D, E, F, G, H: k_get 15, E, F, G, H
    k_set 16, A, B, C, D: sb3 A, B, C, D, E, F, G, H: k_get 16, E, F, G, H
    k_set 17, A, B, C, D: sb2 A, B, C, D, E, F, G, H: k_get 17, E, F, G, H
    k_set 18, A, B, C, D: sb1 A, B, C, D, E, F, G, H: k_get 18, E, F, G, H
    k_set 19, A, B, C, D: sb0 A, B, C, D, E, F, G, H: k_get 19, E, F, G, H
    k_set 20, A, B, C, D: sb7 A, B, C, D, E, F, G, H: k_get 20, E, F, G, H
    k_set 21, A, B, C, D: sb6 A, B, C, D, E, F, G, H: k_get 21, E, F, G, H
    k_set 22, A, B, C, D: sb5 A, B, C, D, E, F, G, H: k_get 22, E, F, G, H
    k_set 23, A, B, C, D: sb4 A, B, C, D, E, F, G, H: k_get 23, E, F, G, H
    k_set 24, A, B, C, D: sb3 A, B, C, D, E, F, G, H: k_get 24, E, F, G, H
    k_set 25, A, B, C, D: sb2 A, B, C, D, E, F, G, H: k_get 25, E, F, G, H
    k_set 26, A, B, C, D: sb1 A, B, C, D, E, F, G, H: k_get 26, E, F, G, H
    k_set 27, A, B, C, D: sb0 A, B, C, D, E, F, G, H: k_get 27, E, F, G, H
    k_set 28, A, B, C, D: sb7 A, B, C, D, E, F, G, H: k_get 28, E, F, G, H
    k_set 29, A, B, C, D: sb6 A, B, C, D, E, F, G, H: k_get 29, E, F, G, H
    k_set 30, A, B, C, D: sb5 A, B, C, D, E, F, G, H: k_get 30, E, F, G, H
    k_set 31, A, B, C, D: sb4 A, B, C, D, E, F, G, H: k_get 31, E, F, G, H
    k_set 32, A, B, C, D: sb3 A, B, C, D, E, F, G, H: k_get 32, E, F, G, H
    set_key = l_key
End Function
Private Sub SetKey(Key As String)
    Dim kblk(0 To 7) As MemBlock, KeyBuffer() As Byte
    ReDim KeyBuffer(0 To 31) As Byte
    SetBuffer KeyBuffer, Key, 32
    SetBlock kblk, KeyBuffer, 0, 32
    set_key kblk, 32 * 8
End Sub

