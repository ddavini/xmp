VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form xmp 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "X-MaD.Player.1"
   ClientHeight    =   4050
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   4950
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "xmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "xmp"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock ctrlWsk 
      Left            =   4320
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XmP1.xmDisplay xmDVol 
      Height          =   135
      Left            =   3045
      TabIndex        =   32
      Top             =   1375
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
   End
   Begin XmP1.xmDisplay BitRate 
      Height          =   195
      Left            =   3600
      TabIndex        =   31
      ToolTipText     =   "Frequency khz"
      Top             =   1375
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   344
   End
   Begin XmP1.xmDisplay lblFreq 
      Height          =   195
      Left            =   3990
      TabIndex        =   30
      ToolTipText     =   "Bit Rate (Kbit)"
      Top             =   1375
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   344
   End
   Begin XmP1.xmDisplay Durata 
      Height          =   135
      Left            =   3100
      TabIndex        =   29
      ToolTipText     =   "Time Left"
      Top             =   1030
      Width           =   550
      _ExtentX        =   873
      _ExtentY        =   238
   End
   Begin XmP1.xmDisplay lblMode 
      Height          =   135
      Left            =   3840
      TabIndex        =   28
      Top             =   1030
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Mute"
      Top             =   1080
      Width           =   240
   End
   Begin XmP1.xmSlide xmsVol 
      Height          =   735
      Left            =   4415
      TabIndex        =   26
      Top             =   915
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   1296
   End
   Begin XmP1.xmDisplay xmDisplay 
      Height          =   135
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   390
      Width           =   4170
      _ExtentX        =   2990
      _ExtentY        =   450
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   1440
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Pause"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Info"
      Top             =   1560
      Width           =   225
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   2150
      ScaleHeight     =   210
      ScaleWidth      =   360
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Show Volume"
      Top             =   1080
      Width           =   390
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Visual Mode"
      Top             =   1320
      Width           =   240
   End
   Begin VB.PictureBox LitePic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4335
      ScaleHeight     =   135
      ScaleWidth      =   150
      TabIndex        =   19
      ToolTipText     =   "Show Task"
      Top             =   120
      Width           =   180
   End
   Begin VB.PictureBox Logo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   18
      Top             =   120
      Width           =   195
   End
   Begin VB.PictureBox PicPosizione 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   165
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox gpeak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2160
      ScaleHeight     =   180
      ScaleWidth      =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   120
   End
   Begin VB.PictureBox analyzer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox gph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1920
      ScaleHeight     =   390
      ScaleWidth      =   45
      TabIndex        =   14
      Top             =   3600
      Width           =   45
   End
   Begin VB.PictureBox abuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   1575
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox picModo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   3675
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   9
      ToolTipText     =   "Mode"
      Top             =   1030
      Width           =   165
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   1200
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Open"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox MinimizzaPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4500
      ScaleHeight     =   135
      ScaleWidth      =   150
      TabIndex        =   5
      ToolTipText     =   "Minimize"
      Top             =   120
      Width           =   180
   End
   Begin VB.PictureBox MenuBarPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4680
      ScaleHeight     =   120
      ScaleWidth      =   150
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   180
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   960
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Next"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   720
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Stop"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Play"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Back"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Timer UnSec 
      Left            =   120
      Top             =   2520
   End
   Begin XmP1.SpectrumCtrl SpectrumSin 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   873
   End
   Begin XmP1.SpectrumCtrl SpectrumDes 
      Height          =   495
      Left            =   1000
      TabIndex        =   10
      Top             =   960
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   873
   End
   Begin VB.PictureBox picCardioSin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   350
      ScaleHeight     =   495
      ScaleWidth      =   600
      TabIndex        =   11
      ToolTipText     =   "CardioOSC"
      Top             =   960
      Width           =   600
   End
   Begin VB.PictureBox picCardioDes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1110
      ScaleHeight     =   495
      ScaleWidth      =   600
      TabIndex        =   12
      ToolTipText     =   "CardioOSC"
      Top             =   960
      Width           =   600
   End
   Begin VB.PictureBox picSfondoSlide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   855
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin XmP1.xmDisplay xmDisplay 
      Height          =   135
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   530
      Width           =   4170
      _ExtentX        =   2990
      _ExtentY        =   450
   End
   Begin VB.Image ImgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   240
      Top             =   900
      Width           =   1500
   End
   Begin VB.Line lnSposta2 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line lnPosizione 
      BorderColor     =   &H0000FF00&
      X1              =   165
      X2              =   4515
      Y1              =   785
      Y2              =   785
   End
   Begin VB.Image LowHigh 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   260
      Left            =   4580
      Stretch         =   -1  'True
      ToolTipText     =   "Shock"
      Top             =   400
      Width           =   250
   End
   Begin VB.Line lnSposta 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   195
      Y2              =   195
   End
End
Attribute VB_Name = "xmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private X1 As Integer, Y1 As Integer
Private Mouvement As Boolean
Private Button2 As Boolean
Private MouseUp As Boolean
Private VolumeVisibile As Boolean
Private WithEvents DragMP3 As DragEngine
Attribute DragMP3.VB_VarHelpID = -1
Private SlideMove As Boolean
Private XSlide1 As Long
Private xmpVisMod(6) As Boolean
Private LockButton As Boolean
Private CommandImgIdice As Byte
Private ModuloVolume As Byte
Private Storelbl As String
Private StoreSec As Long

Public Mute As Boolean
Public Acaso As Boolean
Public g_PlayDone As Boolean
Public g_Ripeti As Boolean

Private Sub analyzer_Click()
    ImgLogo.Visible = True
    analyzer.Visible = False
    SetCardioOSCVisibile False
End Sub


Public Sub CommandImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    
    Call IlluminaPulsante(CommandImg(Index), True)

    CommandImgIdice = Index
    MouseUp = False
    LockButton = False
    
    Select Case Index
        Case Is = 0 'Prev'
            MoveInMp3 (-1)
        Case Is = 1 'Play
            If mStreamIsPause Then
                mResumeStream
            Else
                Call PlayStream(frmListone.ListaMp3.ListIndex)
            End If
        Case Is = 2 'Stop Speciale (diverso da StopAll)
            gPosForSave.Mp3File = ""
            gPosForSave.SamplePos = 0
            mStopStream
            AzzeraPosizione
        Case Is = 3 'Next
            Call MoveInMp3(1)
        Case Is = 5 'Apri
        
            frmListone.Enabled = False
            Call OpenMp3dlg
            Call IlluminaPulsante(CommandImg(Index), False)
            Call InfoNmp3
            ChDir App.Path
            Pause
            frmListone.Enabled = True
    
        Case Is = 4  'Pausa
            If mStreamIsActive Then
                mPauseStream
            Else
                mResumeStream
            End If
        Case Is = 6  'Info dialog
            InfoBoxShow
            Call IlluminaPulsante(CommandImg(Index), False)
        Case Is = 7 'Volume visibile
            xmDVol.Visible = Not xmDVol.Visible
            VolumeVisibile = xmDVol.Visible
            If xmDVol.Visible Then
                frmVolume.tmrHide.Enabled = False
                Call ShowLedVol
            Else
                frmVolume.tmrHide.Enabled = True
                Unload frmVolume
            End If
        Case Is = 8 'Mute
            Mute = Not Mute
            LockButton = Mute
            If Mute Then
                ModuloVolume = mStreamVol
                xmsVol.xValue = xmsVol.xMax
                Call ShowLedVol
            Else
                xmsVol.xValue = xmsVol.xMax - ModuloVolume
                If VolumeVisibile = False Then
                    xmDVol.Visible = False
                    Unload frmVolume
                End If
            End If
        Case Is = 9 'xmp Vis Mode
            If analyzer.Visible = True Then
                GlobalSpectrumeMode = GlobalSpectrumeMode + 1
                If GlobalSpectrumeMode > StereoOSC Then
                    GlobalSpectrumeMode = PeakFalls
                End If
            End If
    End Select

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.CommandImg_MouseDown" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub CommandImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH

    CommandImgIdice = Index
    If Not LockButton Then
        Call IlluminaPulsante(CommandImg(Index), False)
    End If
    LockButton = False
    MouseUp = True
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.CommandImg_MouseUp" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub ctrlWsk_Connect()
On Error GoTo ErrH
        Select Case gb_MHACTIONS
            Case Is = LOGIN
                Call ctrlWsk.SendData("GET /default.mhd?action=login&appname=xmp" & vbCrLf & vbCrLf)
            Case Is = LOGOUT
                Call ctrlWsk.SendData("GET /default.mhd?action=logout&sessionid=" & gb_MHSESSIONID & vbCrLf & vbCrLf)
            Case Is = SNDDATA
                Call ctrlWsk.SendData("GET /default.mhd?action=senddata&sessionid=" & gb_MHSESSIONID & "&data=" & gb_MHDATA & "&datatype=" & gb_MHDATATYPE & vbCrLf & vbCrLf)
            Case Is = SHOWUI
                Call ctrlWsk.SendData("GET /default.mhd?action=sendcmd&sessionid=" & gb_MHSESSIONID & "&data=SHOWUI&datatype=0" & vbCrLf & vbCrLf)
            Case Is = HIDEUI
                Call ctrlWsk.SendData("GET /default.mhd?action=sendcmd&sessionid=" & gb_MHSESSIONID & "&data=HIDEUI&datatype=0" & vbCrLf & vbCrLf)
        End Select
ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.ctrlWsk_Connect" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub ctrlWsk_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrH
    Dim sBuffer As String
    
    If bytesTotal > 0 Then
        Call ctrlWsk.GetData(sBuffer, vbString)
        
        sBuffer = Mid$(sBuffer, InStr(sBuffer, vbCrLf & vbCrLf) + 4)
        gb_MHRESPONS = sBuffer
        
    End If
ErrH:
    If Err.Number <> 0 Then
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.ctrlWsk_DataArrival" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrH
    
    If KeyAscii = 27 Then
        Call DeLoad(True)
    Else
        Call frmHotKey(KeyAscii)
    End If
  
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.Form_KeyPress" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_Activate()
    Me.lnSposta2.BorderColor = gcActiveColor
    Me.lnSposta.BorderColor = gcActiveColor
End Sub

Private Sub Form_Deactivate()
    Me.lnSposta2.BorderColor = gcDeactiveColor
    Me.lnSposta.BorderColor = gcDeactiveColor
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim InfoLine As String
    Dim Flag As Boolean
                                   
If Command$ <> "/debug" Then
    Set DragMP3 = New DragEngine
    DragMP3.DragHwnd = Me.hWnd
    DragMP3.StartDrag
    Call StartSubClassingHotkeys(Me.hWnd)
End If
        
    xmDVol.Visible = False
    xmDVol.Center = True
    xmsVol.xMax = 100
    Durata.Center = True
    Durata.xCaption = "00:00"
    lblMode.Center = True
    lblMode.xCaption = "Mode"
    BitRate.Center = True
    lblFreq.Center = True
    
    'Display
    xmDisplay(0).Size = 5
    xmDisplay(1).Size = 5
    InfoLine = GetINI(cfgFile, "VISUALIZATION", "DISPSTR", "")
    If InfoLine = "" Then
        xmDisplay(1).xCaption = "*** v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " " + App.Comments & " ***"
    Else
        xmDisplay(1).xCaption = InfoLine
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Carica Res
    'LoadPicture(App.Path & "\img_new\xmp_.bmp")
    Me.Picture = LoadDataIntoFile("BKGXMPCPRS")
    ImgLogo.Stretch = True
    ImgLogo.Picture = LoadDataIntoFile("LOGOXMPCPRS")
    LitePic.Picture = LoadResPicture("PICLITE", vbResBitmap)
    MinimizzaPic.Picture = LoadResPicture("PICMIN", vbResBitmap)
    picModo.Picture = LoadResPicture("MODO0", vbResBitmap)
    MenuBarPic.Picture = LoadResPicture("EXIT", vbResBitmap)
    xmp.Icon = LoadResPicture("SKULL16", vbResIcon)
    'ImgZOL.Picture = LoadResPicture("LOGOZOL", vbResBitmap)
    LowHigh.Picture = LoadResPicture("LOGOJAP", vbResBitmap)
    Logo.Picture = LoadResPicture("PICCLOSE", vbResBitmap)
    gph.Picture = LoadResPicture("PICGPH", vbResBitmap)
    gpeak.Picture = LoadResPicture("PICPEAKS", vbResBitmap)
    
    For I = 0 To CommandImg.count - 1
        CommandImg(I).AutoSize = True
    Next I
      
'    CommandImg(0).Picture = LoadPicture(App.Path & "\img_new\back_.bmp") 'LoadResPicture("PICBACK", vbResBitmap)
'    CommandImg(1).Picture = LoadPicture(App.Path & "\img_new\play_.bmp") 'LoadResPicture("PICPLAY", vbResBitmap)
'    CommandImg(2).Picture = LoadPicture(App.Path & "\img_new\stop_.bmp") 'LoadResPicture("PICSTOP", vbResBitmap)
'    CommandImg(3).Picture = LoadPicture(App.Path & "\img_new\next_.bmp") 'LoadResPicture("PICNEXT", vbResBitmap)
'    CommandImg(4).Picture = LoadPicture(App.Path & "\img_new\pause_.bmp") 'LoadResPicture("PICPAUSE", vbResBitmap)
'    CommandImg(5).Picture = LoadPicture(App.Path & "\img_new\eject_.bmp") 'LoadResPicture("PICOPEN", vbResBitmap)
        
    CommandImg(0).Picture = LoadResPicture("PICBACK", vbResBitmap)
    CommandImg(1).Picture = LoadResPicture("PICPLAY", vbResBitmap)
    CommandImg(2).Picture = LoadResPicture("PICSTOP", vbResBitmap)
    CommandImg(3).Picture = LoadResPicture("PICNEXT", vbResBitmap)
    CommandImg(4).Picture = LoadResPicture("PICPAUSE", vbResBitmap)
    CommandImg(5).Picture = LoadResPicture("PICOPEN", vbResBitmap)
        
        
    For I = 1 To 5
        CommandImg(I).Left = CommandImg(I - 1).Left + CommandImg(I - 1).Width
    Next I
    
'    CommandImg(6).Picture = LoadPicture(App.Path & "\img_new\info_.bmp") 'LoadResPicture("PICINFO", vbResBitmap)
'    CommandImg(7).Picture = LoadPicture(App.Path & "\img_new\volume_.bmp") 'LoadResPicture("PICSHOWVOL", vbResBitmap)
'    CommandImg(8).Picture = LoadPicture(App.Path & "\img_new\mute_.bmp") 'LoadResPicture("PICMUTE", vbResBitmap)
'    CommandImg(9).Picture = LoadPicture(App.Path & "\img_new\specmode_.bmp") 'LoadResPicture("PICSPECMODE", vbResBitmap)
              
    CommandImg(6).Picture = LoadResPicture("PICINFO", vbResBitmap)
    CommandImg(7).Picture = LoadResPicture("PICSHOWVOL", vbResBitmap)
    CommandImg(8).Picture = LoadResPicture("PICMUTE", vbResBitmap)
    CommandImg(9).Picture = LoadResPicture("PICSPECMODE", vbResBitmap)
              
              
    Flag = CBool(GetINI(cfgFile, "VISUALIZATION", "SCALE", "TRUE"))
    If Flag Then
        analyzer.Picture = LoadDataIntoFile("SCALE")
        abuff.Height = analyzer.Height
        abuff.Width = analyzer.Width
        abuff.Picture = analyzer.Picture
        picCardioDes.Picture = analyzer.Picture
        picCardioSin.Picture = analyzer.Picture
    End If
    
    SpectrumDes.Modello gph.Picture
    SpectrumSin.Modello gph.Picture
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Load LedPic
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Slide
    lnPosizione.Visible = False
    picSfondoSlide.AutoRedraw = True
    picSfondoSlide.Left = lnPosizione.X1
    picSfondoSlide.Top = lnPosizione.Y1 - (PicPosizione.Height \ 2)
    picSfondoSlide.Height = PicPosizione.Height
    picSfondoSlide.Width = lnPosizione.X2 - lnPosizione.X1
    picSfondoSlide.BackColor = vbBlack
    
    'Disegna linea della morte sonica, no cioe' la linea della slide
    Call DrawOCenterLine(picSfondoSlide, RGB(206, 206, 206))
    PicPosizione.Width = 135 + 15
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Spectrum
    xmp.analyzer.Left = 200
    xmp.analyzer.Height = AltezzaSpectrumDisplay * Screen.TwipsPerPixelX
    abuff.Height = xmp.analyzer.Height
    xmp.analyzer.Width = AmpiezzaSpectrumDisplay * Screen.TwipsPerPixelY
    xmp.analyzer.Top = 1080
    abuff.Width = xmp.analyzer.Width
    analyzer.ScaleMode = vbPixels
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Cardio
    picCardioDes.Height = analyzer.Height
    picCardioSin.Height = analyzer.Height
    
    picCardioDes.Top = analyzer.Top
    picCardioSin.Top = analyzer.Top
        
    SpectrumDes.Height = analyzer.Height
    SpectrumSin.Height = analyzer.Height
    SpectrumDes.Top = analyzer.Top
    SpectrumSin.Top = analyzer.Top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'xmp
    xmp.Height = 2010
    Call RoundForm(Me)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    Acaso = False
    LockButton = False
    Mute = False
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    If Button = 2 Then
        Button2 = True
        PopupMenu frmMenu.mnuFile
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.CommandImg_MouseDown" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Dim lngReturnValue As Long
On Error Resume Next

    If Button = 0 Or Button2 Then
        Y1 = Y
        X1 = X
    End If
    
    If Y < 500 And X >= lnSposta.X1 And X <= lnSposta.X2 Or Mouvement Then
            If Button = 1 Then
                Mouvement = True
                Me.Left = Me.Left - (X1 - X)
                Me.Top = Me.Top - (Y1 - Y)
                If AgganciatoFlag Then
                    AgganciaSub (True)
                End If
            End If
    End If
    Button2 = False
    
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
       Case WM_LBUTTONDOWN
            Call SetForegroundWindow(frmMenu.hWnd)
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
            Call SetForegroundWindow(frmMenu.hWnd)
            Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hWnd)
            Me.Visible = True
       Case WM_RBUTTONDOWN
            Call SetForegroundWindow(frmMenu.hWnd)
            Pause
            PopupMenu frmMenu.mnuFile
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mouvement = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrH

If Command$ <> "/debug" Then
    DragMP3.StopDrag
    Pause
    Set DragMP3 = Nothing
    Pause
    Call StopSubClassingHotkeys(Me.hWnd)
    Pause
End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.Form_Unload" & vbCrLf & Err.Description)
    End If
End Sub




Private Sub LitePic_Click()
    Call ShowTaskBarIcon(Not frmShowTask.Visible)
End Sub

Private Sub LitePic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(LitePic, True)
End Sub

Private Sub LitePic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(LitePic, False)
End Sub

Private Sub Logo_Click()
    Call DeLoad(True)
End Sub

Private Sub Logo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(Logo, True)
End Sub

Private Sub Logo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(Logo, False)
End Sub

Private Sub LowHigh_Click()
    Dim Mode As Byte
    Mode = GetINI(cfgFile, "VISUALIZATION", "SHOCKMODE", 0)
    Select Case Mode
        Case Is = 0
            Randomize
            Mode = (Rnd * 2)
        Case Else
            Mode = Mode - 1
    End Select
    Call ScorriLabelNomeMp3(0, Mode, 0.03)
End Sub

Private Sub LowHigh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LowHigh.Appearance = 1
End Sub

Private Sub LowHigh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LowHigh.Appearance = 0
End Sub

Private Sub MenuBarPic_Click()
   Call DeLoad(True)
End Sub

Private Sub MenuBarPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, True)
End Sub

Private Sub MenuBarPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, False)
End Sub

Private Sub MinimizzaPic_Click()
    Call MinimizzaXMP
End Sub

Private Sub MinimizzaPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MinimizzaPic, True)
End Sub

Private Sub MinimizzaPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MinimizzaPic, False)
End Sub

Private Sub picCardioDes_Click()
    analyzer.Visible = True
    SetCardioOSCVisibile False
    ImgLogo.Visible = False
End Sub

Private Sub picCardioSin_Click()
    analyzer.Visible = True
    SetCardioOSCVisibile False
    ImgLogo.Visible = False
End Sub


Private Sub ImgLogo_Click()
    SetCardioOSCVisibile True
    analyzer.Visible = False
    ImgLogo.Visible = False
End Sub

Private Sub PicPosizione_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Disegna linea della morte sonica, no cioe' la linea della slide
    SlideMove = True
    Call DrawOCenterLine(picSfondoSlide, vbGreen)
End Sub

Private Sub PicPosizione_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    If Button = 0 Then
        XSlide1 = X
    End If
        
    If Button = 1 Or SlideMove Then
        Select Case PicPosizione.Left - (XSlide1 - X)
            Case Is > (lnPosizione.X2 - PicPosizione.Width)
                PicPosizione.Left = (lnPosizione.X2 - PicPosizione.Width)
            Case Is <= lnPosizione.X1
                PicPosizione.Left = lnPosizione.X1
            Case Else
                PicPosizione.Left = PicPosizione.Left - (XSlide1 - X)
        End Select
    End If
        
            
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.picPosizione_MouseMove" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub PicPosizione_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    SlideMove = False
    X = PicPosizione.Left - X
    'Disegna linea della morte sonica, no cioe' la linea della slide
    Call DrawOCenterLine(picSfondoSlide, RGB(206, 206, 206))
    
    gPosForSave.SamplePos = PosizioneSlide(X, picSfondoSlide.Width, mStreamLen)
    Call mStreamPos(gPosForSave.SamplePos)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.PicPosizione_MouseUp" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub picSfondoSlide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mStreamPos(PosizioneSlide(X, picSfondoSlide.Width, mStreamLen))
End Sub

Private Sub picSfondoSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSfondoSlide.MousePointer = vbCustom
    picSfondoSlide.MouseIcon = LoadResPicture("INETHAND", vbResCursor)
End Sub

Private Sub UnSec_Timer()
On Error GoTo ErrH
    Dim posizioneImg As Long
    Dim sWidth As Long
    Dim StrmPos As Long
    Dim StrmLen As Long
    Dim Sec As Long
   
    CommandImg(9).Enabled = analyzer.Visible
     
'    Select Case mPlayState
'        Case Is = Play_Status.STATE_PAUSE
'            WriteINFO "Pause."
'        Case Is = Play_Status.STATE_PLAY
'            WriteINFO "Play."
'        Case Is = Play_Status.STATE_SEEK
'            WriteINFO "Seek."
'        Case Is = Play_Status.STATE_STOP
'            WriteINFO "Stop."
'    End Select
           
    If mStreamIsActive And Not SlideMove Then
        ''' pias no '''
        If CBool(GetINI(cfgFile, "PREFERENCE", "SHOWRATE", "False")) Then
            BitRate.xCaption = mGetCurrentBitRate & "K"
        End If
        '''''''''''''''
        
        sWidth = lnPosizione.X2 - lnPosizione.X1 - PicPosizione.Width
        Durata.xCaption = DurataStream(mStreamLenInSeconds - mStreamPosInSeconds)
        StrmLen = mStreamLen
        StrmPos = mStreamPos
        gPosForSave.SamplePos = StrmPos
        posizioneImg = (StrmPos * (sWidth / StrmLen)) \ 1
        If posizioneImg + lnPosizione.X1 <= lnPosizione.X2 - PicPosizione.Width Then
            PicPosizione.Left = posizioneImg + lnPosizione.X1
        Else
            PicPosizione.Left = lnPosizione.X2 - PicPosizione.Width
            WriteINFO "Warning: Stream Len Error"
            gPosForSave.SamplePos = 0
        End If
    End If

    If g_PlayDone = True Then
        Call PlayDone
    End If
             
    Call ControllaStrToPrecIstance
       
    Call DisplayAll(GlobalSpectrumeMode)
        
    If mStreamPosInSeconds <> StoreSec Then
                StoreSec = mStreamPosInSeconds
                Call modMicroHandler.SendCommand(SNDDATA, "ROWS=" & xmp.xmDisplay(0).xCaption & "|- " & DurataStream(mStreamLenInSeconds - mStreamPosInSeconds))
    End If
    
'If Command$ = "/debug" Then
'    Dim sStr As String
'    sStr = Round(mXNormLevel * 100, 2)
'    sStr = "Normalization: " & sStr
'    frmListone.xmDInfo(1).xCaption = sStr
'End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmp.UnSec_Timer" & vbCrLf & Err.Description)
        Call DeLoad(True)
    End If
End Sub

Private Sub DragMP3_FilesDroped()
    Call GestioneExtDrag(DragMP3)
End Sub

Private Sub xmDisplay_Click(Index As Integer)
    Call ScorriLabelNomeMp3(55) ' 0, 0, 0.03
End Sub

Private Sub xmsVol_Change()
    Call mStreamVol(xmsVol.xMax - xmsVol.xValue, True)
    If CBool(GetINI(cfgFile, "VISUALIZATION", "SHOWVOSC", "FALSE")) Or xmDVol.Visible Then
        Call ShowLedVol
    End If
End Sub
