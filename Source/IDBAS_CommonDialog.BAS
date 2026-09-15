Attribute VB_Name = "modCommonDialog"
Option Explicit
Option Base 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private hHook As Long

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 256

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public FileDialog As OPENFILENAME
Private ParenthWnd As Long

Public Function ShowOpen(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
On Error GoTo ErrH
    Dim Ret As Long
    Dim count As Integer
    Dim fileNameHolder As String
    Dim LastCharacter As Integer
    Dim NewCharacter As Integer
    Dim tempFiles() As String
    Dim hInst As Long
    Dim Thread As Long
    Dim I As Byte
    Dim Multiselect As Long
    Dim OpenBuffer As Long
    
        OpenBuffer = CLng(GetINI(cfgFile, "PREFERENCE", "OBUFFER", "8", 1, 32)) * 1024
    
        FileDialog.sFile = ""
        ParenthWnd = hWnd
        FileDialog.nStructSize = Len(FileDialog)
        FileDialog.nFileSize = 1024
        FileDialog.hwndOwner = hWnd
        FileDialog.sFileTitle = Space$(2048)
        FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
        FileDialog.sFile = FileDialog.sFile & Space$(OpenBuffer) & Chr$(0)
        FileDialog.nFileSize = Len(FileDialog.sFile)
        
        'If FileDialog.flags = 0 Then
            FileDialog.Flags = OFS_FILE_OPEN_FLAGS
        'End If
        
        'Set up the CBT hook
        hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        If centerForm = True Then
            hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
        Else
            hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
        End If
        
        Ret = GetOpenFileName(FileDialog)
        
        FileDialog.sFile = Trim$(FileDialog.sFile)
        For I = 1 To 2
            Multiselect = InStr(Multiselect + 1, FileDialog.sFile, vbNullChar)
        Next
        
         
        If Ret Then
            'Trim$(FileDialog.sFileTitle) = "" And
            If Mid(FileDialog.sFile, Multiselect + 1, 1) <> vbNullChar And _
               Trim$(Mid(FileDialog.sFile, Multiselect + 1, 1)) <> "" Then
                LastCharacter = 0
                count = -1
                While ShowOpen.nFilesSelected = 0
                    NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                    If count >= 0 Then
                        tempFiles(count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                    Else
                        ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                    End If
                    count = count + 1
                    ReDim Preserve tempFiles(count)
                    If InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar) = _
                       InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar & vbNullChar) Then
                        tempFiles(count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                        ShowOpen.nFilesSelected = count + 1
                    End If
                    
                    LastCharacter = NewCharacter
                Wend
                ReDim ShowOpen.sFiles(ShowOpen.nFilesSelected - 1)
                For count = 0 To ShowOpen.nFilesSelected - 1
                    ShowOpen.sFiles(count) = tempFiles(count)
                Next
            Else
                ReDim ShowOpen.sFiles(0)
                ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
                ShowOpen.nFilesSelected = 1
                ShowOpen.sFiles(0) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, _
                InStr(FileDialog.nFileOffset + 1, FileDialog.sFile, vbNullChar) - FileDialog.nFileOffset - 1)
            End If
            ShowOpen.bCanceled = False
            Exit Function
        Else
            ShowOpen.sLastDirectory = ""
            ShowOpen.nFilesSelected = 0
            ShowOpen.bCanceled = True
            Erase ShowOpen.sFiles
            Exit Function
        End If
        
        ParenthWnd = 0
        EraseFileDialog
        
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IDBAS_CommonDialog.ShowOpen" & vbCrLf & Err.Description)
    End If
End Function

Public Sub EraseFileDialog()
    Dim sEmpty As OPENFILENAME
    FileDialog = sEmpty
End Sub

Public Function ShowSave(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
On Error GoTo ErrH
Dim Ret As Long
Dim hInst As Long
Dim Thread As Long

    ParenthWnd = hWnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hWnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)

    If FileDialog.Flags = 0 Then
        FileDialog.Flags = OFS_FILE_SAVE_FLAGS
    End If

    'Set up the CBT hook
    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If

    Ret = GetSaveFileName(FileDialog)
    ReDim ShowSave.sFiles(1)

    If Ret Then
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IDBAS_CommonDialog.ShowSave" & vbCrLf & Err.Description)
    End If
End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error GoTo ErrH
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the DisplayError at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
        'Debug.Print "Screen " & Screen.Height / 2
        'Debug.Print "DisplayError " & (rectMsg.Right - rectMsg.Left) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IDBAS_CommonDialog.WinProcCenterScreen" & vbCrLf & Err.Description)
    End If
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error GoTo ErrH
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the DisplayError centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        'Position the DisplayError
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IDBAS_CommonDialog.WinProcCenterScreenForm" & vbCrLf & Err.Description)
    End If
End Function
