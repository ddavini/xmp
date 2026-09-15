Attribute VB_Name = "IniModule"
Option Explicit

'OPEN URL
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5

Private Const SM_CXVSCROLL = 2
Private Const SM_CXHSCROLL = 21

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SWW_HPARENT = -8

Public Const LB_FINDSTRING = &H18F
Public Const EM_UNDO = &HC7

Public Sub LoadFileINI(ByVal FileNamez As String, txtBox As TextBox)
On Error GoTo ErrH
Dim FF As Byte
Dim TextLine As String

FF = FreeFile
txtBox.Text = ""
Open FileNamez For Input As #FF
    Do Until EOF(FF)
        Line Input #FF, TextLine
        'If Left(textline, 1) = "[" Then lstTemp.AddItem textline
        txtBox.Text = txtBox.Text + TextLine + vbCrLf
    Loop
Close #FF
Exit Sub
ErrH:
    Call DisplayError(Err.Description, Dove:=Err.Source)
End Sub


Public Function SaveFileINI(ByVal FileNamez As String, txtBox As TextBox) As Boolean
Dim FF As Byte
Dim Response As VbMsgBoxResult
On Error GoTo ErrH
    FF = FreeFile
    If Trim$(txtBox.Text) = "" Then
        Response = MsgBox("The file is empty. Are you sure you want to save this file?", vbOKCancel)
        If Response = vbCancel Then
            SaveFileINI = False
            Exit Function
        End If
    End If
    Open FileNamez For Output As #FF
        Print #FF, txtBox.Text
    Close #FF
    SaveFileINI = True
ErrH:
    If Err.Number <> 0 Then
        MsgBox (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "DisplayError" & vbCrLf & Err.Description)
        SaveFileINI = False
    End If
End Function

Public Sub DisplayError(message As String, Optional Tipo As VbMsgBoxStyle = vbCritical, _
                        Optional Titolo As String = "Error", _
                        Optional Dove As String = "Error")
On Error GoTo ErrH
    Dim ErrorNum As Long
    MsgBox message & vbCrLf & vbCrLf & "Error Number: " _
    & ErrorNum & vbCrLf, vbCritical, Titolo
    Call ScriviLOG(Dove, "Error Number: " + CStr(ErrorNum), message)
ErrH:
    If Err.Number <> 0 Then
         MsgBox (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "DisplayError" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub ScriviLOG(ByVal where As String, ByVal Info As String, ErrorDes As String)

    Dim I As Integer
    Dim sLog As String
    Dim sEsito As String
    Dim Esiste As String
    Dim Filez As String
    Dim FileNumber As Integer
    
On Error Resume Next
        Filez = App.Path + "\Logs\"
        Esiste = Dir(Filez)
        If Esiste = "" Then
            MkDir (App.Path + "\Logs\")
        End If
On Error GoTo ErrH

        sLog = CStr(Time()) & " - Dove: " & where & " " & Info & "  Dettagli: " & ErrorDes

        Filez = App.Path + "\Logs\" + CStr(Date) + ".log"
        Filez = Replace$(Filez, "/", ".")
        Esiste = Dir(Filez)
        If Esiste = "" Then
            FileNumber = FreeFile
            Open Filez For Output As #FileNumber
                Print #FileNumber, "Error Log File."
                Print #FileNumber, sLog
            Close #FileNumber
        Else
            FileNumber = FreeFile
            Open Filez For Append As #FileNumber
                Print #FileNumber, sLog
            Close #FileNumber
        End If

ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ScriviLog" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub HandCursorOver(Obj As Object)
    If Obj.MousePointer <> vbCustom Then
        Obj.MousePointer = vbCustom
        Obj.MouseIcon = LoadResPicture("INETHAND", vbResCursor)
    End If
End Sub

Public Sub OpenUrl(hWnd As Long, Optional Url As String = "http://www.zolnetwork.com/x-mad/xmp")
On Error Resume Next
    ShellExecute hWnd, "open", Url, "", "", SW_SHOW Or SW_NORMAL
ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.xShell" & vbCrLf & Err.Description)
    End If
End Sub


Public Sub LoadINI(Filename As String)
Dim I As Integer
Dim Rand As Integer
    With frmMain
        Call LoadFileINI(Filename, .txtIni)
        
        .Saved = True
        '.OpenedFile = Filename
        
        'Add item to recent file list
        .mnuSep4.Visible = True
        
        .RefreshSectionList
        
        
        For I = 1 To .mnuRecent.count - 1
            If .mnuRecent(I).Caption = Filename Then Exit Sub
        Next I
        
        If .mnuRecent.count = 5 Then
            Rand = Int((4 * Rnd) + 1)
            .mnuRecent(Rand).Caption = Filename
            .mnuRecent(Rand).Visible = True
        Else
            Load .mnuRecent(.mnuRecent.count)
            .mnuRecent(.mnuRecent.count - 1).Caption = Filename
            .mnuRecent(.mnuRecent.count - 1).Visible = True
        End If
    End With
End Sub


Public Function HScrollBar() As Long
On Error Resume Next
    HScrollBar = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IniModule.HScrollBar" & vbCrLf & Err.Description)
    End If
End Function

Public Function VScrollBar() As Long
On Error Resume Next
    VScrollBar = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelY
ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "IniModule.VScrollBar" & vbCrLf & Err.Description)
    End If
End Function

Public Sub SaveBox()

    Call MsgBox("Current INI File Saved", vbInformation, "Kame...Saved")
    
End Sub
