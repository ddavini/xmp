Attribute VB_Name = "modHotKeys"
''''''''''''''''''''''''''''''''''''''
'Based on Craig Jasper, Sample   '
''''''''''''''''''''''''''''''''''''''

Option Explicit

'Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long

Private Const SW_SHOW = 5

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Const MOD_ALT = &H1      ' ALT key must be held down.
Private Const MOD_CONTROL = &H2  ' CTRL key must be held down.
Private Const MOD_SHIFT = &H4    ' SHIFT key must be held down.

Private Const WM_HOTKEY = &H312

Private Const GWL_WNDPROC = -4

Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private g_lngOldWindowProc    As Long
Private g_lngWinRet           As Long
Private g_lngHook             As Long
Private bloSuccess()          As Boolean

'*********************************************************************
' Routine:             CallBack
' Description:         Passes wParam (HotKey) to HotKeyRoutine.
'*********************************************************************
Public Function CallBack(ByVal hWnd As Long, _
     ByVal lMsg As Long, _
     ByVal wParam As Long, _
     ByVal lparam As Long) As Long
 
  On Error GoTo ErrH

  Select Case lMsg
    Case WM_HOTKEY
      HotKeyRoutine wParam
      CallBack = 1
    '  Exit Function
  End Select
    
  ' If the message wasn't a HotKey then let
  ' default message pass:
  CallBack = CallWindowProc(g_lngWinRet, _
     hWnd, lMsg, wParam, lparam)
  
  Exit Function
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modHotkeys.CallBack" & vbCrLf & Err.Description)
    End If
End Function

Public Sub frmHotKey(KeyAscii As Integer)
  Select Case KeyAscii
    Case Asc("p") 'Play
        Call HotKeyRoutine(0)
    Case Asc("s") 'Stop
        Call HotKeyRoutine(1)
    Case Asc("b") 'Back
        Call HotKeyRoutine(2)
    Case Asc("n") 'Next
        Call HotKeyRoutine(3)
    Case Asc("k") 'Kill
        Call HotKeyRoutine(4)
    Case Asc("h") 'Hide
        Call HotKeyRoutine(5)
    Case Asc("v") 'Show
        Call HotKeyRoutine(6)
    Case Asc("+") 'Volume +
        Call HotKeyRoutine(7)
    Case Asc("-") 'Volume -
        Call HotKeyRoutine(8)
    Case Asc("m") 'Mute
        Call HotKeyRoutine(9)
    Case Asc("f") 'Pause
        Call HotKeyRoutine(10)
    Case Asc("x") 'XSound
        Call HotKeyRoutine(11)
  End Select
End Sub


Public Sub HotKeyRoutine(ByVal KeyID As Byte)
    Dim lVol As Long
On Error GoTo ErrH
  
  Select Case KeyID
    Case 0 'Play
      Call PlayStream(frmListone.ListaMp3.ListIndex)
    Case 1 'Stop
      Call StopAll
    Case 2 'Back
      Call MoveInMp3(-1)
    Case 3 'Next
      Call MoveInMp3(1)
    Case 4 'Kill
      Call DeLoad(True)
    Case 5 'Hide
      Call MinimizzaXMP
    Case 6 'Show
        xmp.WindowState = vbNormal
        Call SetForegroundWindow(xmp.hWnd)
        xmp.Show
    Case 7 'Volume +
        If xmp.xmsVol.xValue > 0 Then
            xmp.xmsVol.xValue = xmp.xmsVol.xValue - 1
        End If
        Call ShowLedVol
    Case 8 'Volume -
        xmp.xmDVol.Visible = True
        If xmp.xmsVol.xValue < 100 Then
            xmp.xmsVol.xValue = xmp.xmsVol.xValue + 1
        End If
        Call ShowLedVol
    Case 9
        Call xmp.CommandImg_MouseDown(8, 1, 0, 0, 0)
        If Not xmp.Mute Then
            Call xmp.CommandImg_MouseUp(8, 1, 0, 0, 0)
        End If
    Case 10
        If mStreamIsActive Then
            mPauseStream
        Else
            mResumeStream
        End If
    Case 11
        Call SetXSound(Not frmMenu.mnuXSound.Checked)
  End Select
  
  Exit Sub
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modHotKeys.HotKeyRoutine" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub StartSubClassingHotkeys(hWnd As Long)
On Error GoTo ErrH
  ' Start Subclassing:
    If GetWinVersion >= 5 Then
        g_lngOldWindowProc = GetWindowLong(hWnd, GWL_WNDPROC)
        g_lngWinRet = GetWindowLong(hWnd, GWL_WNDPROC)
        'g_lngWinRet = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf CallBack)
        Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf CallBack)
        Call RegHotKeys(hWnd)
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modHotKeys.StartSubClassingHotkeys" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub StopSubClassingHotkeys(hWnd As Long)
On Error GoTo ErrH
    If GetWinVersion >= 5 Then
        Dim I As Integer
        
        For I = 0 To UBound(bloSuccess)
          If bloSuccess(I) Then
            UnregisterHotKey hWnd, I
          End If
        Next
        Pause
        ' VB puo' crashare / VB can crash
        SetWindowLong hWnd, GWL_WNDPROC, _
           g_lngOldWindowProc
        Pause 0.9
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modHotKeys.StopSubClassingHotkeys" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub RegHotKeys(hWnd As Long)
On Error GoTo ErrH
    
  Dim strKey As String
    
    ReDim bloSuccess(11)
    bloSuccess(0) = RegisterHotKey(hWnd, 0, MOD_ALT + MOD_CONTROL, vbKeyP)
    bloSuccess(1) = RegisterHotKey(hWnd, 1, MOD_ALT + MOD_CONTROL, vbKeyS)
    bloSuccess(2) = RegisterHotKey(hWnd, 2, MOD_ALT + MOD_CONTROL, vbKeyB)
    bloSuccess(3) = RegisterHotKey(hWnd, 3, MOD_ALT + MOD_CONTROL, vbKeyN)
    bloSuccess(4) = RegisterHotKey(hWnd, 4, MOD_ALT + MOD_CONTROL, vbKeyK)
    bloSuccess(5) = RegisterHotKey(hWnd, 5, MOD_ALT + MOD_CONTROL, vbKeyH)
    bloSuccess(6) = RegisterHotKey(hWnd, 6, MOD_ALT + MOD_CONTROL, vbKeyV)
    bloSuccess(7) = RegisterHotKey(hWnd, 7, MOD_ALT + MOD_CONTROL, vbKeyAdd)
    bloSuccess(8) = RegisterHotKey(hWnd, 8, MOD_ALT + MOD_CONTROL, vbKeySubtract)
    bloSuccess(9) = RegisterHotKey(hWnd, 9, MOD_ALT + MOD_CONTROL, vbKeyM)
    bloSuccess(10) = RegisterHotKey(hWnd, 10, MOD_ALT + MOD_CONTROL, vbKeyF)
    bloSuccess(10) = RegisterHotKey(hWnd, 11, MOD_ALT + MOD_CONTROL, vbKeyX)
  Exit Sub
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modHotKeys.RegHotKeys" & vbCrLf & Err.Description)
    End If
End Sub
