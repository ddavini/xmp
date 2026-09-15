Attribute VB_Name = "modLED"
Option Explicit

Private Const SRCCOPY = &HCC0020

Public Sub DrawLED(DestPB As Object, Value As String, GFXPBARRAY As Object, Align As Integer)
    Dim CurChar As String
    Dim I, X, Y As Integer
On Error GoTo ErrH
    If Not DestPB.ScaleMode = 3 Then
        DestPB.ScaleMode = 3
    End If
    If Not DestPB.AutoRedraw = True Then
        DestPB.AutoRedraw = True
    End If
    DestPB.Cls
    Select Case Align
        Case 0
            Value = SetupString(Value, DestPB.ScaleWidth)
            Y = (DestPB.ScaleHeight / 2) - 8
            X = ((DestPB.ScaleWidth / 2) - ((16 * Len(Value) / 2))) - 16
            For I = 1 To Len(Value)
                CurChar = Mid(Value, I, 1)
                BitBlt DestPB.hdc, X + (I * 16), Y, 16, 16, GFXPBARRAY(CInt(CurChar)).hdc, 0, 0, SRCCOPY
            Next
        Case 1
            Y = (DestPB.ScaleHeight / 2) - 8
            X = ((DestPB.ScaleWidth / 2) - ((16 * Len(Value) / 2))) - 16
            For I = 1 To Len(Value)
                CurChar = Mid(Value, I, 1)
                BitBlt DestPB.hdc, X + (I * 16), Y, 16, 16, GFXPBARRAY(CInt(CurChar)).hdc, 0, 0, SRCCOPY
            Next
    End Select
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modLed.DrawLED" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub LoadGFX(GFXPBARRAY As Object, Optional CustomColor As Boolean, Optional BackColor As Long, Optional DiodeColor As Long, Optional LitDiodeColor As Long)
Dim Str1to16(1 To 16) As String
Dim R, C, Number As Byte
Dim CurChar As String
On Error GoTo ErrH
    For Number = 0 To 9
        Select Case Number
            Case 0
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "000$$$$$$$$$$000"
             Str1to16(4) = "00$0000000000$00"
             Str1to16(5) = "00$0000000000$00"
             Str1to16(6) = "00$0000000000$00"
             Str1to16(7) = "00$0000000000$00"
             Str1to16(8) = "00$##########$00"
             Str1to16(9) = "00############00"
            Str1to16(10) = "00$0000000000$00"
            Str1to16(11) = "00$0000000000$00"
            Str1to16(12) = "00$0000000000$00"
            Str1to16(13) = "00$0000000000$00"
            Str1to16(14) = "00$0000000000$00"
            Str1to16(15) = "000$$$$$$$$$$000"
            Str1to16(16) = "0000000000000000"
             
            Case 1
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "00############00"
             Str1to16(3) = "0#0##########$00"
             Str1to16(4) = "0##0000000000$00"
             Str1to16(5) = "0##0000000000$00"
             Str1to16(6) = "0##0000000000$00"
             Str1to16(7) = "0##0000000000$00"
             Str1to16(8) = "0#0##########$00"
             Str1to16(9) = "00############00"
            Str1to16(10) = "0#00000000000$00"
            Str1to16(11) = "0##0000000000$00"
            Str1to16(12) = "0##0000000000$00"
            Str1to16(13) = "0##0000000000$00"
            Str1to16(14) = "0#0##########$00"
            Str1to16(15) = "00############00"
            Str1to16(16) = "0000000000000000"
            
            Case 2
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "0#$$$$$$$$$$$000"
             Str1to16(4) = "0##0000000000$00"
             Str1to16(5) = "0##0000000000$00"
             Str1to16(6) = "0##0000000000$00"
             Str1to16(7) = "0##0000000000$00"
             Str1to16(8) = "0#00000000000$00"
             Str1to16(9) = "00$$$$$$$$$$$000"
            Str1to16(10) = "0$000000000000#0"
            Str1to16(11) = "0$00000000000##0"
            Str1to16(12) = "0$00000000000##0"
            Str1to16(13) = "0$00000000000##0"
            Str1to16(14) = "0$000000000000#0"
            Str1to16(15) = "00$$$$$$$$$$$$#0"
            Str1to16(16) = "0000000000000000"
            
            Case 3
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "0#$$$$$$$$$$$000"
             Str1to16(4) = "0##0000000000$00"
             Str1to16(5) = "0##0000000000$00"
             Str1to16(6) = "0##0000000000$00"
             Str1to16(7) = "0##0000000000$00"
             Str1to16(8) = "0#00000000000$00"
             Str1to16(9) = "00$$$$$$$$$$$$00"
            Str1to16(10) = "0#00000000000$00"
            Str1to16(11) = "0##0000000000$00"
            Str1to16(12) = "0##0000000000$00"
            Str1to16(13) = "0##0000000000$00"
            Str1to16(14) = "0#00000000000$00"
            Str1to16(15) = "00$$$$$$$$$$$000"
            Str1to16(16) = "0000000000000000"
             
            Case 4
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "00############00"
             Str1to16(3) = "0#$##########$00"
             Str1to16(4) = "0#$0000000000$00"
             Str1to16(5) = "0#$0000000000$00"
             Str1to16(6) = "0#$0000000000$00"
             Str1to16(7) = "0#$0000000000$00"
             Str1to16(8) = "0#$0000000000$00"
             Str1to16(9) = "000$$$$$$$$$$$00"
            Str1to16(10) = "0#00000000000$00"
            Str1to16(11) = "0##0000000000$00"
            Str1to16(12) = "0##0000000000$00"
            Str1to16(13) = "0##0000000000$00"
            Str1to16(14) = "0#0##########$00"
            Str1to16(15) = "00###########$00"
            Str1to16(16) = "0000000000000000"
            
            Case 5
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "000$$$$$$$$$$$#0"
             Str1to16(4) = "0#$0000000000##0"
             Str1to16(5) = "0#$0000000000##0"
             Str1to16(6) = "0#$0000000000##0"
             Str1to16(7) = "0#$0000000000##0"
             Str1to16(8) = "00$00000000000#0"
             Str1to16(9) = "000$$$$$$$$$$$00"
            Str1to16(10) = "0#000000000000$0"
            Str1to16(11) = "0##00000000000$0"
            Str1to16(12) = "0##00000000000$0"
            Str1to16(13) = "0##00000000000$0"
            Str1to16(14) = "0#000000000000$0"
            Str1to16(15) = "0#0$$$$$$$$$$$00"
            Str1to16(16) = "0000000000000000"
            
            Case 6
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "000$$$$$$$$$$$#0"
             Str1to16(4) = "00$0000000000##0"
             Str1to16(5) = "00$0000000000##0"
             Str1to16(6) = "00$0000000000##0"
             Str1to16(7) = "00$0000000000##0"
             Str1to16(8) = "00$00000000000#0"
             Str1to16(9) = "000$$$$$$$$$$$00"
            Str1to16(10) = "00$00000000000$0"
            Str1to16(11) = "00$00000000000$0"
            Str1to16(12) = "00$00000000000$0"
            Str1to16(13) = "00$00000000000$0"
            Str1to16(14) = "00$00000000000$0"
            Str1to16(15) = "000$$$$$$$$$$$00"
            Str1to16(16) = "0000000000000000"
             
            Case 7
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "0#0$$$$$$$$$$$00"
             Str1to16(4) = "0##00000000000$0"
             Str1to16(5) = "0##00000000000$0"
             Str1to16(6) = "0##00000000000$0"
             Str1to16(7) = "0##00000000000$0"
             Str1to16(8) = "0#0##########0$0"
             Str1to16(9) = "00############00"
            Str1to16(10) = "0#000000000000$0"
            Str1to16(11) = "0##00000000000$0"
            Str1to16(12) = "0##00000000000$0"
            Str1to16(13) = "0##00000000000$0"
            Str1to16(14) = "0#0##########0$0"
            Str1to16(15) = "00############00"
            Str1to16(16) = "0000000000000000"
             
            Case 8
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "000$$$$$$$$$$$00"
             Str1to16(4) = "00$00000000000$0"
             Str1to16(5) = "00$00000000000$0"
             Str1to16(6) = "00$00000000000$0"
             Str1to16(7) = "00$00000000000$0"
             Str1to16(8) = "00$00000000000$0"
             Str1to16(9) = "000$$$$$$$$$$$00"
            Str1to16(10) = "00$00000000000$0"
            Str1to16(11) = "00$00000000000$0"
            Str1to16(12) = "00$00000000000$0"
            Str1to16(13) = "00$00000000000$0"
            Str1to16(14) = "00$00000000000$0"
            Str1to16(15) = "000$$$$$$$$$$$00"
            Str1to16(16) = "0000000000000000"
            
            Case 9
             Str1to16(1) = "0000000000000000"
             Str1to16(2) = "0000000000000000"
             Str1to16(3) = "000$$$$$$$$$$$0"
             Str1to16(4) = "00$00000000000$0"
             Str1to16(5) = "00$00000000000$0"
             Str1to16(6) = "00$00000000000$0"
             Str1to16(7) = "00$00000000000$0"
             Str1to16(8) = "00$00000000000$0"
             Str1to16(9) = "000$$$$$$$$$$$00"
            Str1to16(10) = "0#000000000000$0"
            Str1to16(11) = "0##00000000000$0"
            Str1to16(12) = "0##00000000000$0"
            Str1to16(13) = "0##00000000000$0"
            Str1to16(14) = "0#000000000000$0"
            Str1to16(15) = "00$$$$$$$$$$$$00"
            Str1to16(16) = "0000000000000000"
            
        End Select
        GFXPBARRAY(Number).AutoRedraw = True
        GFXPBARRAY(Number).ScaleMode = 3
        
        For C = 1 To 16
            For R = 1 To 16
                CurChar = Mid(Str1to16(C), R, 1)
                If Not CustomColor Then
                    If CurChar = "0" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), &H8000&
                    If CurChar = "$" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), vbGreen
                    If CurChar = "#" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), &H7000&
                Else
                    If CurChar = "0" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), BackColor
                    If CurChar = "$" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), LitDiodeColor
                    If CurChar = "#" Then GFXPBARRAY(Number).PSet (R - 1, C - 1), DiodeColor
                End If
            Next
        Next
    Next
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modLed.LoadGFX" & vbCrLf & Err.Description)
    End If
End Sub

Function SetupString(Str1 As String, Width As Integer) As String
    Dim Caps As Integer
    Dim OString As String
    Dim I As Integer
On Error GoTo ErrH
    Caps = Width \ 16
    If Len(Str1) < Caps Then
        OString = ""
        For I = 1 To Caps - Len(Str1)
            OString = OString & "0"
        Next
        SetupString = OString & Str1
    Else
        SetupString = Str1
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modLed.SetupString" & vbCrLf & Err.Description)
    End If
End Function
