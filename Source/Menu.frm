VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuXmPButton 
         Caption         =   "Option"
         Begin VB.Menu mnuXS 
            Caption         =   "XSound"
            Begin VB.Menu mnuXSound 
               Caption         =   "Surround"
            End
            Begin VB.Menu mnuXNorm 
               Caption         =   "Normalize"
            End
         End
         Begin VB.Menu mnuTools 
            Caption         =   "Tools"
            Begin VB.Menu mnuHotKey 
               Caption         =   "Hot Keys"
            End
            Begin VB.Menu mnuEQ 
               Caption         =   "EQ"
            End
            Begin VB.Menu nmuxmpMove 
               Caption         =   "Center XmP"
            End
            Begin VB.Menu mnuAggancia 
               Caption         =   "Link Mp3 List"
            End
         End
         Begin VB.Menu mnuLoop 
            Caption         =   "Loop"
         End
         Begin VB.Menu mnuRnd 
            Caption         =   "Random"
         End
         Begin VB.Menu mnuHideList 
            Caption         =   "Hide List"
         End
         Begin VB.Menu InPrimoPiano 
            Caption         =   "Always On Top"
         End
      End
      Begin VB.Menu mnuPreference 
         Caption         =   "Preference"
      End
      Begin VB.Menu S0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "Command"
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuBack 
            Caption         =   "Back"
         End
         Begin VB.Menu mnuNext 
            Caption         =   "Next"
         End
         Begin VB.Menu mnuPlay 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
      End
      Begin VB.Menu mnuHTrayMenu 
         Caption         =   "Hide Menu"
      End
      Begin VB.Menu S4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuIDtag 
      Caption         =   "IDtag"
      Begin VB.Menu mnuInfoTag 
         Caption         =   "Info Mp3 Tag"
      End
      Begin VB.Menu mnuEditTag 
         Caption         =   "Edit Mp3 Tag"
      End
   End
   Begin VB.Menu mnuPlayList 
      Caption         =   "PlayListMenu"
      Begin VB.Menu mnuQuickSave 
         Caption         =   "Quick Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuOpenList 
         Caption         =   "Open List"
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLstFileName 
         Caption         =   "List File Name"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InPrimoPianoFlag As Boolean

Public Sub InPrimoPiano_Click()
    InPrimoPiano.Checked = Not InPrimoPiano.Checked
    InPrimoPianoFlag = InPrimoPiano.Checked
    Call SempreInPrimoPiano(xmp.hWnd, xmp.Width, xmp.Height, _
                            xmp.Left, xmp.Top, InPrimoPiano.Checked)
    Call SempreInPrimoPiano(frmListone.hWnd, frmListone.Width, _
                            frmListone.Height, frmListone.Left, _
                            frmListone.Top, InPrimoPiano.Checked)
End Sub


Private Sub mnuAbout_Click()
    Call About
End Sub

Private Sub mnuAggancia_Click()
    Call AgganciaSub(True)
End Sub

Private Sub mnuBack_Click()
    Call MoveInMp3(-1)
End Sub

Private Sub mnuEditTag_Click()
On Error GoTo ErrH:
    Call xmpPensa(True)
    If frmListone.ListaMp3.ListIndex <> -1 And _
        Right$(GetFile(frmListone.ListaMp3.ListIndex), 4) <> ".WAV" Then
        frmTagEdit.NomeFile = GetFile(frmListone.ListaMp3.ListIndex)
        SempreInPrimoPiano frmTagEdit.hWnd, frmTagEdit.Width, frmTagEdit.Height, frmTagEdit.Left, frmTagEdit.Top, InPrimoPiano
        frmTagEdit.Show vbModal, frmListone
    End If
    Call xmpPensa(False)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Menu.mnuEditTag_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub mnuEQ_Click()
    Load frmEQ
    Call SempreInPrimoPiano(frmEQ.hWnd, xmp.Width, xmp.Height, _
                            xmp.Left, xmp.Top, InPrimoPianoFlag)
    frmEQ.Show vbModal
End Sub

Private Sub mnuEsci_Click()
    Call DeLoad(True)
End Sub

Public Sub mnuHideList_Click()
    Call SetHideListone(Not frmListone.Visible)
End Sub

Private Sub mnuHotKey_Click()
On Error GoTo ErrH

        frmInfo.Height = 35 * 200
        frmInfo.cmdChiudi.Top = 33 * 200 - frmInfo.cmdChiudi.Height
        With frmInfo.lstInfo
            frmInfo.lstInfo.Font = "Courier"
            frmInfo.lstInfo.Print ("System Level Hot-Keys (2k+ Only)")
            frmInfo.lstInfo.Print ("Play       : ALT+CTRL+P")
            frmInfo.lstInfo.Print ("Stop       : ALT+CTRL+S")
            frmInfo.lstInfo.Print ("Pause      : ALT+CTRL+F")
            frmInfo.lstInfo.Print ("Back       : ALT+CTRL+B")
            frmInfo.lstInfo.Print ("Next       : ALT+CTRL+N")
            frmInfo.lstInfo.Print ("Exit       : ALT+CTRL+K")
            frmInfo.lstInfo.Print ("Hide       : ALT+CTRL+H")
            frmInfo.lstInfo.Print ("Show       : ALT+CTRL+V")
            frmInfo.lstInfo.Print ("Stop       : ALT+CTRL+S")
            frmInfo.lstInfo.Print ("Volume Up  : ALT+CTRL++")
            frmInfo.lstInfo.Print ("Volume Down: ALT+CTRL+-")
            frmInfo.lstInfo.Print ("Volume Mute: ALT+CTRL+M")
            frmInfo.lstInfo.Print ("XSound.Srnd: ALT+CTRL+X")
            frmInfo.lstInfo.Print ("")
            frmInfo.lstInfo.Print ("Local Level Hot-Keys (On-Focus)")
            frmInfo.lstInfo.Print ("Play       : p")
            frmInfo.lstInfo.Print ("Stop       : s")
            frmInfo.lstInfo.Print ("Pause      : F")
            frmInfo.lstInfo.Print ("Back       : d")
            frmInfo.lstInfo.Print ("Next       : n")
            frmInfo.lstInfo.Print ("Exit       : k")
            frmInfo.lstInfo.Print ("Hide       : h")
            frmInfo.lstInfo.Print ("Show       : v")
            frmInfo.lstInfo.Print ("Stop       : s")
            frmInfo.lstInfo.Print ("Volume Up  : +")
            frmInfo.lstInfo.Print ("Volume Down: -")
            frmInfo.lstInfo.Print ("Volume Mute: m")
            frmInfo.lstInfo.Print ("XSound.Srnd: x")
            frmInfo.lstInfo.Print ("")
            frmInfo.lstInfo.Print ("On List Focus")
            frmInfo.lstInfo.Print ("Remove MP3     : d")
            frmInfo.lstInfo.Print ("Quick Seek MP3 : Shift+'First Letter'")
        End With
        frmInfo.tmrCPU.Interval = 1000
        SempreInPrimoPiano frmInfo.hWnd, frmInfo.Width, frmInfo.Height, frmInfo.Left, frmInfo.Top, frmMenu.InPrimoPiano
        Call frmInfo.Show(vbModal, frmListone)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Menu.mnuHotKey_Click" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mnuHTrayMenu_Click()
    DoEvents
End Sub

Private Sub mnuInfoTag_Click()
    Call xmpPensa(True)
    If frmListone.ListaMp3.ListIndex <> -1 Then
        Call InfoBoxShow(GetFile(frmListone.ListaMp3.ListIndex))
    End If
    Call xmpPensa(False)
End Sub

Public Sub mnuLoop_Click()
    Call SetLoop(Not xmp.g_Ripeti)
End Sub

Private Sub mnuNext_Click()
    Call MoveInMp3(1)
End Sub

Private Sub mnuOpenList_Click()
    Call xmp.CommandImg_MouseDown(5, 0, 0, 0, 0)
End Sub

Private Sub mnuPause_Click()
    If mStreamIsActive Then
        mPauseStream
    Else
        mResumeStream
    End If
End Sub

Private Sub mnuPlay_Click()
    Call PlayStream(frmListone.ListaMp3.ListIndex)
End Sub

Private Sub mnuPreference_Click()
On Error Resume Next
    xmp.Enabled = False
    frmListone.Enabled = False

    frmPreference.ExeName = App.ExeName + ".exe"
    frmPreference.Filez = App.Path & "\xmp.ini"
    
    frmPreference.Show vbModal, xmp
    
    
    xmp.Enabled = True
    frmListone.Enabled = True
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError("Preference Menu Can't Create Please refer to de Autor for This Problem!")
        Call ScriviLOG(Err.Source & ".Menu.mnuPreference_Click", Err.Number, Err.Description)
        xmp.Enabled = True
        frmListone.Enabled = True
        Err.Clear
    End If
End Sub

Private Sub mnuQuickSave_Click()
    Dim ListFile As String
    ListFile = GetIni(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
    Call ScriviListaMp3(ListFile)
    GestisciPosFrm False
End Sub

Public Sub mnuRnd_Click()
   Call SetRnd(Not xmp.Acaso)
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo ErrH:
    
    Dim sFile As SelectedFile
    
    FileDialog.sDefFileExt = "*.ls"
    FileDialog.sFilter = ""
    FileDialog.sFilter = "ls file (*.ls)" & Chr$(0) _
    & "*.ls"

    FileDialog.sInitDir = App.Path
    FileDialog.sDlgTitle = "XmP"
    'FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    FileDialog.Flags = OFS_FILE_SAVE_FLAGS
    
    sFile = ShowSave(frmListone.hWnd, True)
    If isVector(sFile.sFiles) Then
        Call ScriviListaMp3(sFile.sLastDirectory + sFile.sFiles(1))
    End If
    GestisciPosFrm False
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Menu.mnuSaveAs_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub mnuStop_Click()
    Call StopAll
End Sub

Private Sub mnuXNorm_Click()
    Call SetXNorm(Not mnuXNorm.Checked)
End Sub

Private Sub mnuXSound_Click()
    Call SetXSound(Not mnuXSound.Checked)
End Sub

Private Sub nmuxmpMove_Click()
On Error GoTo ErrH
    Dim X As Long, Y As Long
    Dim startTop As Long, startLeft As Long
    Dim centroTop As Long, centroLeft As Long
    Dim xStep As Long, yStep As Long
    
    xmp.WindowState = vbNormal
    Call SetForegroundWindow(xmp.hWnd)
    xmp.Show
    centroLeft = (Screen.Width \ 2) - (xmp.Width \ 2)
    centroTop = (Screen.Height \ 2) - (xmp.Height \ 2)
    startTop = xmp.Top
    startLeft = xmp.Left
    If startTop > centroTop Then
        yStep = -1
    Else
        yStep = 1
    End If
    If startLeft > centroLeft Then
        xStep = -1
    Else
        xStep = 1
    End If
    X = xmp.Left
    For Y = startTop To centroTop Step yStep
            DoEvents
            Call xmp.Move(X, Y)
    Next Y
    Y = xmp.Top
    For X = startLeft To centroLeft Step xStep
        DoEvents
        Call xmp.Move(X, Y)
    Next X
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Menu.mnuxmpMove_Click" & vbCrLf & Err.Description)
    End If
End Sub
