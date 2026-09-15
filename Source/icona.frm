VERSION 5.00
Begin VB.Form frmPreference 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preference"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9630
   ClipControls    =   0   'False
   ForeColor       =   &H0000FF00&
   Icon            =   "icona.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameVec 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Install"
      ForeColor       =   &H0000FF00&
      Height          =   4815
      Index           =   1
      Left            =   5280
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox chkAssocia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Associate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1680
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkAddDesk 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.CheckBox chkAddShort 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblAssociazione 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Associate With Following Extension"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2790
      End
      Begin VB.Label lblAddShort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Start Menu ShortCut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   2115
      End
      Begin VB.Label lblAddDesk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add DeskTop ShortCut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame frameVec 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "xmMP3"
      ForeColor       =   &H0000FF00&
      Height          =   4815
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtBuffer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Errore"
         Top             =   3480
         Width           =   855
      End
      Begin VB.VScrollBar vscBuffer 
         Height          =   285
         Left            =   1080
         Max             =   10
         TabIndex        =   16
         Top             =   3480
         Value           =   3
         Width           =   135
      End
      Begin VB.CheckBox chk16bit 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "16bit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chk8bit 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "8bit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox chkMono 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Mono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox chkStereo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Stereo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   975
      End
      Begin VB.ComboBox cmdSamRate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   990
         Width           =   1575
      End
      Begin VB.ComboBox cmbDevice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "icona.frx":000C
         Left            =   1560
         List            =   "icona.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2550
      End
      Begin VB.Label lblBuffer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buffer Len."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lblUnder 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "See XmP.INI for More Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   2940
      End
      Begin VB.Label lblMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Modo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblBit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Bit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblSamRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound Device"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.ListBox lstApplication 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      ItemData        =   "icona.frx":0010
      Left            =   120
      List            =   "icona.frx":0012
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPreference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Filez As String
Public ExeName As String

Private VecCfgApp() As String
Private VecCfgKeys() As String
Private VecCfgValue() As String
'Sempre in primo piano
Private Const conHwndTopmost = -1
Private Const conHwndNoTopmost = -2
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40

Private Sub chk16bit_Click()
    chk8bit.Value = Abs(chk16bit.Value - 1)
End Sub

Private Sub chk8bit_Click()
    chk16bit.Value = Abs(chk8bit.Value - 1)
End Sub

Private Sub chkAddDesk_Click()
On Error GoTo ErrH:
    If chkAddDesk.Value <> 0 Then
        chkAddDesk.Value = 0
        Call CreateShortcut(CSIDL_DESKTOP, "XmP", App.Path + "\" + ExeName)
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.chkAddDesk_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub chkAddShort_Click()
On Error GoTo ErrH:
    If chkAddShort.Value <> 0 Then
        chkAddShort.Value = 0
        Call CreateShortcut(CSIDL_PROGRAMS, "XmP", App.Path + "\" + ExeName)
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.chkAddShort_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub chkAssocia_Click()
On Error GoTo ErrH:
    If chkAssocia.Value <> 0 Then
        Dim I As Integer
        chkAssocia.Value = 0
        For I = 0 To lstExt.ListCount - 1
            If Not Associate(App.Path + "\" + ExeName, lstExt.List(I)) Then
                Call DisplayError("FAILED", , "Association")
            End If
        Next
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.chkAssocia_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub chkMono_Click()
    chkStereo.Value = Abs(chkMono.Value - 1)
End Sub

Private Sub chkStereo_Click()
    chkMono.Value = Abs(chkStereo.Value - 1)
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call SAVExmMP3
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    
    Dim I As Byte
    For I = 0 To frameVec.count - 1
        frameVec(I).Top = frameVec(0).Top
        frameVec(I).Left = frameVec(0).Left
        frameVec(I).Height = frameVec(0).Height
        frameVec(I).Width = frameVec(0).Width
    Next
    
    Call Elabora
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.Form_Load" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Elabora()
On Error GoTo ErrH
    Dim I As Byte
    Dim X As Long, Y As Long
    
    Me.Height = 5880
    Me.Width = 6820
    X = Screen.Width / Screen.TwipsPerPixelX / 2 - Me.Width / 2
    Y = Screen.Height / Screen.TwipsPerPixelY / 2 - Me.Height / 2
    
    SempreInPrimoPiano Me.hWnd, Me.Width, Me.Height, X, Y, frmMenu.InPrimoPiano.Checked
    Call SetForegroundWindow(Me.hWnd)
    'Me.Show vbModal, xmp
    VecCfgApp = GetIniAllApps(Filez)
    
    Call lstApplication.AddItem("0 - MPx Engine")
    Call lstApplication.AddItem("1 - Install")
    lstApplication.ListIndex = 0
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.Form_Load" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub lblUnder_Click()
On Error GoTo ErrH
    
    If Dir(App.Path + "\" + "xmINI.exe") <> "" Then
        Call OpenCMD(Me.hWnd, "xmINI.exe", "xmp.ini")
    Else
        Call OpenUrl(Me.hWnd, "xmp.ini")
    End If
   
ErrH:
If Err.Number <> 0 Then
    Call DisplayError(Err.Number & vbCrLf & Err.Source & _
            vbCrLf & "xmServicefrm.lblUnder_Click" & vbCrLf & Err.Description)
End If
End Sub

Private Sub lblUnder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call HandCursorOver(lblUnder)
    lblUnder.ToolTipText = "For Advanced User Only!"
End Sub

Private Sub lstApplication_Click()
On Error GoTo ErrH
    Erase VecCfgApp
    Erase VecCfgKeys
    Erase VecCfgValue
    
    Select Case lstApplication.ListIndex
        Case Is = 0
            Call xmMP3
        Case Is = 1
            Call INSTALL
        Case Else
            Dim I As Byte
            For I = 0 To frameVec.count - 1
                frameVec(I).Visible = False
            Next
    End Select
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.lstApplication_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub INSTALL()
On Error GoTo ErrH
    Dim Ext As String
    Dim Store As String
    Dim L As Integer
    Dim I As Integer
    Call EraseFrame
    frameVec(1).Visible = True
    lstApplication.SetFocus
    Ext = GetIni(Filez, "INSTALL", "EXT", "MP3")
    I = 1
    lstExt.Clear
    Do
        L = InStr(I, Ext, ";", vbTextCompare) - I
        Store = Mid$(Ext, I, L)
        Call lstExt.AddItem(Store)
        I = I + L + 1
    Loop Until I > Len(Ext)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.INSTALL" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub xmMP3()
On Error GoTo ErrH
    Const Appz As String = "xmMP3"
    Dim DefValue As String
    Dim I As Byte
    Dim J As Integer
    Dim DeviceNumber As Long
    Dim Capa As WAVEOUTCAPS
    Dim ATemp() As String
    
    cmbDevice.Clear
    cmdSamRate.Clear
    Erase VecCfgKeys
    VecCfgKeys = GetIniAppAllKeys(Filez, Appz)
    For I = 0 To UBound(VecCfgKeys)
        ReDim Preserve VecCfgValue(I)
        Select Case UCase(VecCfgKeys(I))
            Case Is = "DEVICE"
                DefValue = WAVE_MAPPER
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
                DeviceNumber = waveOutGetNumDevs - 1
                For J = -1 To DeviceNumber
                    Call waveOutGetDevCaps(J, Capa, Len(Capa))
                    ATemp = Split(Capa.szPname, vbNullChar)
                    Call cmbDevice.AddItem(ATemp(0))
                Next J
                Select Case VecCfgValue(I)
                    Case Is = WAVE_MAPPER
                        cmbDevice.ListIndex = 0
                    Case Else
                        cmbDevice.ListIndex = VecCfgValue(I) + 1
                End Select
            Case Is = "SAMPLERATE"
                DefValue = "44100"
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
                Call cmdSamRate.AddItem("44100")
                Call cmdSamRate.AddItem("22050")
                Call cmdSamRate.AddItem("11025")
                Select Case VecCfgValue(I)
                    Case Is = "44100"
                        cmdSamRate.ListIndex = 0
                    Case Is = "22050"
                        cmdSamRate.ListIndex = 1
                    Case Is = "11025"
                cmdSamRate.ListIndex = 2
                End Select
            Case Is = "8BIT"
                DefValue = "False"
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
                If CBool(VecCfgValue(I)) Then
                    chk8bit.Value = 1
                Else
                    chk16bit.Value = 1
                End If
            Case Is = "MONO"
                DefValue = "False"
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
                If CBool(VecCfgValue(I)) Then
                    chkMono.Value = 1
                Else
                    chkStereo.Value = 1
                End If
            Case Is = "BUFFER"
                DefValue = "1"
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
                txtBuffer.Text = VecCfgValue(I)
                vscBuffer.Value = txtBuffer.Text
            Case Else
                VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), DefValue)
        End Select
    Next
    Call EraseFrame
    frameVec(0).Visible = True
 '   cmbDevice.SetFocus
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.xmMP3" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub SAVExmMP3()
On Error GoTo ErrH
    Const Appz As String = "xmMP3"
    Dim I As Byte
    
    VecCfgKeys = GetIniAppAllKeys(Filez, Appz)
    For I = 0 To UBound(VecCfgValue)
        ReDim Preserve VecCfgValue(I)
         VecCfgValue(I) = GetIni(Filez, Appz, VecCfgKeys(I), "ERROR!")
    Next
    
    For I = 0 To UBound(VecCfgKeys)
        Select Case UCase(VecCfgKeys(I))
            Case Is = "DEVICE"
                Select Case cmbDevice.ListIndex
                    Case Is = 0
                        VecCfgValue(I) = -1
                    Case Else
                        VecCfgValue(I) = cmbDevice.ListIndex - 1
                End Select
            Case Is = "SAMPLERATE"
                Select Case cmdSamRate.ListIndex
                    Case Is = 0
                        VecCfgValue(I) = 44100
                    Case Is = 1
                        VecCfgValue(I) = 22050
                    Case Is = 2
                        VecCfgValue(I) = 11025
                    Case Else
                        VecCfgValue(I) = 44100
                End Select
            Case Is = "8BIT"
                If chk8bit.Value = 1 Then
                    VecCfgValue(I) = "True"
                Else
                    VecCfgValue(I) = "False"
                End If
            Case Is = "MONO"
                If chkMono.Value = 1 Then
                    VecCfgValue(I) = "True"
                Else
                    VecCfgValue(I) = "False"
                End If
            Case Is = "BUFFER"
                    If IsNumeric(txtBuffer.Text) Then
                        VecCfgValue(I) = txtBuffer.Text
                    Else
                        VecCfgValue(I) = 1
                    End If
            Case Else
        End Select
    Next
       
    For I = 0 To UBound(VecCfgValue)
        Call SetIni(Filez, Appz, VecCfgKeys(I), VecCfgValue(I))
    Next
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.SAVExmMP3" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub vscBuffer_Change()
    txtBuffer.Text = vscBuffer.Value
End Sub
Private Sub EraseFrame()
On Error GoTo ErrH
    Dim I As Byte
    For I = 0 To frameVec.count - 1
        frameVec(I).Visible = False
    Next
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmServicefrm.EraseFrame" & vbCrLf & Err.Description)
    End If
End Sub
