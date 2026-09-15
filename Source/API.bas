Attribute VB_Name = "modChiamateAPI"
Option Explicit

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
'Public Declare Sub ReleaseCapture Lib "user32" ()
'
'Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const HTCAPTION = 2

Private Const WM_MOUSEWHEEL = &H20A

Private Const MAXPNAMELEN = 32
Public Const WAVE_MAPPER = -1&

Public Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long

Public Enum StandardIconEnum
    IDI_ASTERISK = 32516&       ' like vbInformation
    IDI_EXCLAMATION = 32515&    ' like vbExlamation
    IDI_HAND = 32513&           ' like vbCritical
    IDI_QUESTION = 32514&       ' like vbQuestion
End Enum

'Public Declare Function GetInputState Lib "user32" () As Long

Public Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As StandardIconEnum) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long


Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

'Round Form
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'**************************************
'Windows API/Global Declarations for :Ho
'     w to detect if mousewheel scrolls
'**************************************


Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long


Public Type POINTAPI
    X As Long
    Y As Long
    End Type


Public Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lparam As Long
    time As Long
    pt As POINTAPI
    End Type
    
'''Preference'''
'Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 256&
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type WIN32_FIND_DATA
'   dwFileAttributes As Long
'   ftCreationTime As FILETIME
'   ftLastAccessTime As FILETIME
'   ftLastWriteTime As FILETIME
'   nFileSizeHigh As Long
'   nFileSizeLow As Long
'   dwReserved0 As Long
'   dwReserved1 As Long
'   cFileName As String * MAX_PATH
'   cAlternate As String * 14
'End Type
'
'Private Declare Function FindFirstFile Lib "kernel32" _
'   Alias "FindFirstFileA" _
'   (ByVal lpFileName As String, _
'   lpFindFileData As WIN32_FIND_DATA) As Long
'
'Private Declare Function FindClose Lib "kernel32" _
'   (ByVal hFindFile As Long) As Long
'Enum FlagConstants
'  cdlOFNAllowMultiselect = &H200 'The user can select more than one file atrun time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
'  cdlOFNCreatePrompt = &H2000 ' Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
'  cdlOFNExplorer = &H80000 ' Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
'  CdlOFNExtensionDifferent = &H400 ' Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
'  cdlOFNFileMustExist = &H1000 ' Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
'  cdlOFNHelpButton = &H10 ' Causes the dialog box to display the Help button.
'  cdlOFNHideReadOnly = &H4 'Hides the Read Onlycheck box.
'  cdlOFNLongNames = &H200000 ' Use long filenames.
'  cdlOFNNoChangeDir = &H8 'Forces the dialog box to set the current directory to what it was when the dialog box was opened.
'  CdlOFNNoDereferenceLinks = &H100000 ' Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
'  cdlOFNNoLongNames = &H40000 ' No long file names.
'  CdlOFNNoReadOnlyReturn = &H8000 ' Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
'  cdlOFNNoValidate = &H100 ' Specifies that the common dialog box allows invalid characters in the returned filename.
'  cdlOFNOverwritePrompt = &H2 'Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
'  cdlOFNPathMustExist = &H800 ' Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
'  cdlOFNReadOnly = &H1 'Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
'  cdlOFNShareAware = &H4000 ' Specifies that sharing violation errors will be ignored.
'End Enum

'*****************************************associate
Private Declare Function RegCreateKey& Lib "advapi32.dll" _
Alias "RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, lphKey&)

Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" _
(ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)

' Return codes from Registration functions.
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1
'************************************************

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_FAVORITES = &H1F

Private Declare Function SHAddToRecentDocs Lib "shell32.dll" (ByVal dwFlags As Long, ByVal dwData As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByValsPath As String) As Long
'''Preference'''

'PUT in The Tray
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Left-click
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

'Right-click
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

Private Declare Function Shell_NotifyIcon Lib "shell32" _
   Alias "Shell_NotifyIconA" _
   (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nid As NOTIFYICONDATA

'OPEN URL
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Form Trasparente
Private Const GWL_EXSTYLE = -20
Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const WS_EX_LAYERED = &H80000

'Rese Publiche xche usate in anche da fuori'
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'''
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal cKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long

'Noscosi le Scrolla Barre della listazza del listone
'Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, _
'    ByVal wBar As Long, ByVal bShow As Long) As Long
'
'Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, _
'    ByVal wSBflags As Long, ByVal wArrows As Long) As Long
'Private Const SB_HORZ = 0   ' horizontal scrollbar
'Private Const SB_VERT = 1   ' vertical scrollbar
'Private Const SB_CTL = 2    ' scollbar control
'Private Const SB_BOTH = 3   ' both horiz & vert scrollbars
'
'Private Const ESB_ENABLE_BOTH = &H0   ' enable both arrows
'Private Const ESB_DISABLE_LTUP = &H1  ' disable left/up arrows
'Private Const ESB_DISABLE_RTDN = &H2  ' disable right/down arrows
'Private Const ESB_DISABLE_BOTH = &H3  ' disable both arrows

'Sempre in primo piano
Private Const conHwndTopmost = -1
Private Const conHwndNoTopmost = -2
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40
'Larghezza ScrollBar

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXVSCROLL = 2
Private Const SM_CXHSCROLL = 21
'''
'Rese Publiche xche usate anche da fuori'
Public Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

'Public Const THREAD_BASE_PRIORITY_IDLE = -15
'Public Const THREAD_BASE_PRIORITY_LOWRT = 15
'Public Const THREAD_BASE_PRIORITY_MIN = -2
'Public Const THREAD_BASE_PRIORITY_MAX = 2
'Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
'Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
'Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
'Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
'Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
'Public Const THREAD_PRIORITY_NORMAL = 0
'Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
'Public Const HIGH_PRIORITY_CLASS = &H80
'Public Const IDLE_PRIORITY_CLASS = &H40
'Public Const ABOVE_NORMAL_PRIORITY = &H8000 '32768
'Public Const NORMAL_PRIORITY_CLASS = &H20
'Public Const REALTIME_PRIORITY_CLASS = &H100

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long

Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef sorg As Any, ByVal noByte As Long)

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
'Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
'''

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long


Public Type VersionInfo
    Major As Integer
    Minor As Integer
    Revision As Integer
End Type

Public Function DisplayVerInfo(FullFileName As String) As VersionInfo
    Dim rc As Long, lDummy As Long, sBuffer() As Byte
    Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
    Dim lVerbufferLen As Long
    Dim Store() As String
    
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
       DisplayVerInfo "No Version Info available!"
       Exit Function
    End If
    
    '**** Store info to udtVerBuffer struct ****
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    CopyMemory udtVerBuffer, ByVal lVerPointer, Len(udtVerBuffer)
    
    '**** Determine File Version number ****
    Store = Split(Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl), _
             ".")
        
    DisplayVerInfo.Major = Store(0)
    DisplayVerInfo.Minor = Store(1)
    DisplayVerInfo.Revision = Store(UBound(Store))
    
End Function

Public Sub chgTrasp(tVal As Integer, frm As Form)
On Error Resume Next
    If GetWinVersion >= 5 Then
        Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(frm.hWnd, 0, (255 * tVal) / 100, LWA_ALPHA)
    End If
Err.Clear
End Sub

Public Sub FadeIN(frm As Form)
On Error GoTo ErrH
    Dim I As Integer
    For I = 0 To 100
        Call Pause(, True)
        chgTrasp I, frm
    Next
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.FadeIN" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub ICONtray(frm As Form, ToolTip As String, Optional NIM_XX As Long = NIM_ADD)
On Error GoTo ErrH
         
         nid.cbSize = Len(nid)
         nid.hWnd = frm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = xmp.Icon
         
         nid.szTip = ToolTip & vbNullChar
         Shell_NotifyIcon NIM_XX, nid
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.ICONtray" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub DeICONtray()
On Error GoTo ErrH
    WriteINFO "Erase Tray Icon..."
    Shell_NotifyIcon NIM_DELETE, nid
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.DeICONtray" & vbCrLf & Err.Description)
    End If
End Sub

Public Function GetWinVersion() As Integer
On Error GoTo ErrH
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    GetWinVersion = Ver And &HFFFF&
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.GetWinVersion" & vbCrLf & Err.Description)
    End If
End Function


'Public Sub DeScrollList(hWndList As Long)
'On Error GoTo ErrH
'    EnableScrollBar hWndList, SB_VERT, ESB_DISABLE_BOTH
'    ShowScrollBar hWndList, SB_VERT, False
'ErrH:
'    If Err.Number <> 0 Then
'        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
'                vbCrLf & "ChiamateAPI.DeScrollList" & vbCrLf & Err.Description)
'    End If
'End Sub

Public Sub SempreInPrimoPiano(hWnd As Long, _
                              Largo As Long, _
                              Alto As Long, _
                              X As Long, _
                              Y As Long, _
                              Optional AOnTop As Boolean = True)
On Error GoTo ErrH
    X = X \ Screen.TwipsPerPixelX
    Y = Y \ Screen.TwipsPerPixelY
    Largo = Largo \ Screen.TwipsPerPixelY
    Alto = Alto \ Screen.TwipsPerPixelX
    If AOnTop Then
        SetWindowPos hWnd, conHwndTopmost, X, Y, Largo, Alto, _
                    conSwpNoActivate
                    'conSwpNoActivate Or conSwpShowWindow
    Else
        SetWindowPos hWnd, conHwndNoTopmost, X, Y, Largo, Alto, _
                    conSwpNoActivate
                    'conSwpNoActivate Or conSwpShowWindow
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.SempreInPrimoPiano" & vbCrLf & Err.Description)
    End If
End Sub

Public Function LarghezzaScrollBar() As Long
On Error Resume Next
    LarghezzaScrollBar = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelY
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.LarghezzaScrollBar" & vbCrLf & Err.Description)
    End If
End Function




'Da estrarre e mette nelle globali'
Public Sub CreateShortcut(CSIDL As Long, NomeShort As String, TargetPath As String, _
                          Optional Parametro As String, Optional StartFolder As String, _
                          Optional IconNum As Integer, Optional IconPath As String, _
                          Optional WindowMode As Integer)
On Error GoTo ErrH

' CSIDL_PROGRAMS  = Programs
' CSIDL_STARTUP = Startup
' CSIDL_RECENT = RecentDocs
' CSIDL_DESKTOP = Desktop
                       
Dim N0 As Integer
Dim L0 As Long
Dim Shortcut1 As String
Dim Shortcut0 As String
Dim N1 As Integer
Dim X1 As String * 1
Dim X0 As String * 1
Dim Y0 As String * 2
Dim L1 As Long

Dim Endete As Boolean
Dim P As Long
Dim I As Integer
Dim X As String
Dim ToShortCutter As String


If IsNumeric(CSIDL) Then
    ToShortCutter = GetSpecialfolder(CInt(CSIDL))
ElseIf Dir$(ToShortCutter, vbDirectory) = "" Then
    Exit Sub
End If

If SHAddToRecentDocs(2, TargetPath) > 0 Then


    Shortcut0 = GetSpecialfolder(8) & "\" & FileFolder(TargetPath) & ".lnk"

    Endete = False
    Do Until (Dir$(Shortcut0) <> "") Or Endete
        DoEvents
        Sleep 5000
        Endete = True
    Loop

    N0 = FreeFile()
    Open Shortcut0 For Binary Access Read As #N0
        Do Until LOF(N0) > 0
            DoEvents
        Loop
        L0 = LOF(N0)
    
        Shortcut1 = ToShortCutter & "\" & NomeShort & ".lnk"
        N1 = FreeFile()
        Open Shortcut1 For Binary Access Write As #N1
            P = (L0 - 4)
            Y0 = ""
            Do Until (P <= 0) Or (Y0 = vbNullChar & vbNullChar)
                Get #N0, P, Y0
                P = P - 1
            Loop
            L1 = P + 2
    
            For P = 1 To L1
                Get #N0, P, X0
        
                Select Case P
                    Case 21 'path for icon, startup, parameters
                        I = 3
                        If StartFolder <> "" Then
                            I = I + 16
                        End If
                        If Parametro <> "" Then
                            I = I + 32
                        End If
                        If (IconPath <> "") Or (IconNum > 0) Then
                            I = I + 64
                        End If
                        X1 = Chr$(I)
                    Case 57 'Icon index
                        X1 = Chr$(IconNum)
                    Case 61 'Window mode
                        X1 = Chr$(WindowMode)
                    Case Else
                        X1 = X0
                    End Select
                Put #N1, P, X1
            Next P
        Close #N0
        Call Kill(Shortcut0)
    
        X = ""
        If IsMissing(StartFolder) Then
            X = X & Chr$(Len(StartFolder)) & vbNullChar & StartFolder
        End If
        If IsMissing(Parametro) Then
            X = X & Chr$(Len(Parametro)) & vbNullChar & Parametro
        End If
        If IconPath = "" Then
            If IconNum > 0 Then
                X = X & Chr$(Len(TargetPath)) & vbNullChar & TargetPath
            End If
        Else
            X = X & Chr$(Len(IconPath)) & vbNullChar & IconPath
        End If
        X = X & String(4, vbNullChar)
        Put #N1, L1 + 1, X
    Close #N1

Else
    Call Err.Raise(vbError, , "Shotcut Not Create")
End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.CreateShortCut" & vbCrLf & Err.Description)
    End If
End Sub

Private Function GetSpecialfolder(CSIDL As Long) As String
Dim R     As Long
Dim F     As Long
Dim pidl  As Long
Dim sPath As String

On Error GoTo ErrH:

R = SHGetSpecialFolderLocation(frmPreference.hWnd, CSIDL, pidl)

If R = 0 Then

    sPath = Space$(260)
    R = SHGetPathFromIDList(ByVal pidl, ByVal sPath)
    If R Then
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If

End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.GetSpecialfolder" & vbCrLf & Err.Description)
    End If
End Function

Private Function FileFolder(FullPath As String) As String

Dim I As Integer

On Error GoTo ErrH:

FileFolder = FullPath
I = Len(FullPath)
Do Until I = 0
    If Mid$(FullPath, I, 1) = "\" Then
        FileFolder = Mid$(FullPath, I + 1)
        I = 0
    Else
        I = I - 1
    End If
Loop

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.FileFolder" & vbCrLf & Err.Description)
    End If
End Function


Public Function Associate(ByVal apPath As String, ByVal Ext As String) As Boolean

  Dim sKeyName As String 'Holds Key Name in registry.
  Dim sKeyValue As String 'Holds Key Value in registry.
  Dim Ret& 'Holds Error status If any from API calls.
  Dim lphKey& 'Holds created key handle from RegCreateKey.
  
  Dim apTitle As String
  
On Error GoTo ErrH:

    If Not FileExists(apPath) Then
        Exit Function
    End If
    
    apTitle = ParseName(apPath)
    
    If InStr(Ext, ".") = 0 Then
        Ext = "." & Ext
    End If
    
    'register .ext files as belonging to app
    sKeyName = Ext
    sKeyValue = apTitle
    Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    
    'set open command to path and filename
    sKeyName = apTitle
    sKeyValue = apPath & " %1"
    Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    Ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    
    'register app icon with .ext files
    sKeyValue = apPath
    Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    Ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
    If Ret& <> 0 Then
        GoTo ErrH
    End If
    
    Associate = True
    Exit Function
  
ErrH:
    Associate = False
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.Associate" & vbCrLf & Err.Description)
    End If
End Function

Public Function FileExists(sSource As String) As Boolean
On Error GoTo ErrH:
     
    If Dir$(sSource) <> "" Then
        FileExists = True
        Call Dir$
    Else
        FileExists = False
    End If
    
    Exit Function
ErrH:
    Err.Clear
    FileExists = False
End Function

Public Sub SetRoundForm(ByVal hWnd As Long, _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByVal X3 As Long, ByVal Y3 As Long)
On Error GoTo ErrH:
    
    SetWindowRgn hWnd, CreateRoundRectRgn(X1, Y1, X2, Y2, X3, Y3), True
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.SetRoundForm" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub OpenUrl(hWnd As Long, Optional Url As String = "http://www.zolnetwork.com/x-mad/")
On Error Resume Next
    ShellExecute hWnd, "open", Url, "", "", SW_SHOW Or SW_NORMAL
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.OpenUrl" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub OpenCMD(hWnd As Long, CMD As String, Optional Options As String = "")
On Error Resume Next
        ShellExecute 0&, "open", CMD, Options, 0&, SW_SHOW Or SW_NORMAL
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.OpenCMD" & vbCrLf & Err.Description)
    End If
End Sub


Private Function ParseName(ByVal sPath As String) As String
  Dim strX As String
  Dim intX As Integer
  
On Error GoTo ErrH
  
  intX = InStrRev(sPath, "\")
  
  strX = Trim(Right(sPath, Len(sPath) - intX))
  If Right(strX, 1) = Chr(0) Then
    ParseName = Left(strX, Len(strX) - 1)
  Else
    ParseName = strX
  End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.ParseName" & vbCrLf & Err.Description)
    End If
End Function

Public Function MouseWheel(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error GoTo ErrH

    If wMsg = WM_MOUSEWHEEL Then
        If wParam > 0 Then
            frmListone.ListaMp3.ScrollValue = frmListone.ListaMp3.ScrollValue - 1
        Else
            frmListone.ListaMp3.ScrollValue = frmListone.ListaMp3.ScrollValue + 1
        End If
    End If

    MouseWheel = CallWindowProc(gb_OldProcMouseWheel, hWnd, wMsg, wParam, lparam)
        
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modChiamateAPI.MouseWheel" & vbCrLf & Err.Description)
    End If
End Function

'Public Sub ListEvent()
'On Error GoTo errH
'        Dim amsg As Msg
'        Dim lResult As Long
'        lResult = GetMessage(amsg, frmListone.hWnd, 0, 0)
'        lResult = TranslateMessage(amsg)
'        lResult = DispatchMessage(amsg)
'
'        If Command$ <> "/debug" Then
'            If amsg.Message <> 275 Then
'                writeinfo CStr(amsg.Message)
'            End If
'        End If
'
'        Select Case amsg.Message
'            Case Is = 522
'                If amsg.wParam < 0 Then '-7864320
'                    frmListone.ListaMp3.ScrollValue = frmListone.ListaMp3.ScrollValue + 1
'                ElseIf amsg.wParam > 0 Then
'                    frmListone.ListaMp3.ScrollValue = frmListone.ListaMp3.ScrollValue - 1
'                End If
'            Case Is = 256
'                Call frmListone.ListaMp3.UserControl_KeyDown(CInt(amsg.wParam), 0)
'        End Select
'errH:
'    If Err.Number <> 0 Then
'        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
'                vbCrLf & "modFunzioniAccessorie.ListEvent(ex Rotella)" & vbCrLf & Err.Description)
'    End If
'End Sub

