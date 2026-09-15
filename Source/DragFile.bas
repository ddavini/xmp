Attribute VB_Name = "modDragFile"
'Init http://go.to/abubakar

Option Explicit

Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_DROPFILES = &H233
Public PrevWndFunc As Long
Public DragEngineObj As DragEngine

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error GoTo ErrH
    Dim N As Long, iLoop As Long, FileInfo As Long
    Dim Buffer As String * 256
    Dim Length As Long
    If msg = WM_DROPFILES Then
        DragEngineObj.ClearFileNames
        FileInfo = wParam
        N = DragQueryFile(FileInfo, -1&, vbNullString, 0)
        For iLoop = 0 To N - 1
            Buffer = ""
            Length = DragQueryFile(FileInfo, iLoop, ByVal Buffer, 256)
            'Buffer = RipulisciBuffer(Buffer)
            DragEngineObj.AddInFileNames Trim$(Buffer)
        Next
        
        DragEngineObj.NowRaiseEvent
        
        DragFinish FileInfo 'wParam
        WndProc = 0
    Else
        WndProc = CallWindowProc(PrevWndFunc, hWnd, msg, wParam, lparam)
    End If
ErrH:
    'Martellone paura
    Err.Clear
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "DragFile.WndProc" & vbCrLf & Err.Description)
    End If
End Function

Private Function RipulisciBuffer(Buffer As String) As String
On Error GoTo ErrH
    Dim Posizione As Integer
    Dim I As Byte
        Buffer = Replace$(Trim$(Buffer), vbNullChar, "")
        Do
            Posizione = InStr(1, UCase$(Buffer), Estensioni(I))
            I = I + 1
        Loop Until Posizione <> 0 Or I > UBound(Estensioni)
        RipulisciBuffer = Left$(Buffer, Posizione + 3)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "DragFile.RipulisciBuffer" & vbCrLf & Err.Description)
    End If
End Function
