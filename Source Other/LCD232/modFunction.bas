Attribute VB_Name = "modFunction"
Option Explicit

Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Public Sub SendString(sLine() As String)
    Dim I As Integer
    Dim bCols As Byte
    
    bCols = CByte(GetIni(App.Path & "\" & ININAME, "LCD", "COLS", "20"))
    
On Error Resume Next
    frmMain.ComCtrl.Output = Chr(12)
      
    For I = 0 To UBound(sLine)
        sLine(I) = Left$(sLine(I) + Space(bCols), bCols)
        frmMain.ComCtrl.Output = sLine(I)
    Next I
    
End Sub

Public Sub Pause(Optional Durata As Double = 0.03, Optional DoEv As Boolean = False)
On Error GoTo ErrH
    Dim Tick As Single
    Tick = Timer
    While Timer < Tick + Durata
        Call SleepEx(1, 1)
        If DoEv Then
            DoEvents
        End If
    Wend
ErrH:

End Sub

