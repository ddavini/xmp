VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BASS plugin test"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Slider sldPosition 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "click here to open a file..."
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Frame frmPlugins 
      Caption         =   "Loaded plugins"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ListBox lstPlugins 
         Height          =   1035
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   480
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All files|*.*"
   End
   Begin VB.Timer tmrPosition 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   2400
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BASS plugin test, copyright (c) 2005 Sebastian Andersson, MaresWEB, www.maresweb.net
' Based on "plugins.c", by Ian Luck.

Option Explicit

Private chan As Long ' the channel
Private blnSeek As Boolean

Private Sub ErrorM(ByVal es As String)
    MsgBox es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")", vbOKOnly, "Error"
End Sub

' translate a CTYPE value to text
Private Function GetCTypeString(ByVal ctype As Long) As String
    If ctype = BASS_CTYPE_SAMPLE Then
        GetCTypeString = "sample"
        Exit Function
    End If
    If ctype = BASS_CTYPE_RECORD Then
        GetCTypeString = "recording"
        Exit Function
    End If
    If (ctype And BASS_CTYPE_MUSIC_MOD) Then ' MOD music type...
        If (ctype And BASS_CTYPE_MUSIC_MO3) Then
            GetCTypeString = "MO3"
            Exit Function
        End If
        If ctype = BASS_CTYPE_MUSIC_MTM Then
            GetCTypeString = "MTM"
            Exit Function
        End If
        If ctype = BASS_CTYPE_MUSIC_S3M Then
            GetCTypeString = "S3M"
            Exit Function
        End If
        If ctype = BASS_CTYPE_MUSIC_XM Then
            GetCTypeString = "XM"
            Exit Function
        End If
        If ctype = BASS_CTYPE_MUSIC_IT Then
            GetCTypeString = "IT"
            Exit Function
        End If
        GetCTypeString = "MOD"
        Exit Function
    End If
    If (ctype And BASS_CTYPE_STREAM) Then
        If ctype = BASS_CTYPE_STREAM Then
            GetCTypeString = "custom stream"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_WAV Then
            GetCTypeString = "WAV"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_OGG Then
            GetCTypeString = "OGG"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_MP1 Then
            GetCTypeString = "MP1"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_MP2 Then
            GetCTypeString = "MP2"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_MP3 Then
            GetCTypeString = "MP3"
            Exit Function
        End If
        If ctype = BASS_CTYPE_STREAM_AIFF Then
            GetCTypeString = "AIFF"
            Exit Function
        End If
        ' check add-ons...
        If ctype = &H10200 Then
            GetCTypeString = "CDA" ' BASS_CTYPE_STREAM_CD
            Exit Function
        End If
        If ctype = &H10300 Then
            GetCTypeString = "WMA" ' BASS_CTYPE_STREAM_WMA
            Exit Function
        End If
        If ctype = &H10900 Then
            GetCTypeString = "FLAC" ' BASS_CTYPE_STREAM_FLAC
            Exit Function
        End If
        If ctype >= &H10500 And ctype <= &H10503 Then
            GetCTypeString = "Wavpack" ' BASS_CTYPE_STREAM_WV/BASS_CTYPE_STREAM_WV_LH
            Exit Function
        End If
        If ctype = &H10600 Then
            GetCTypeString = "OFR" ' BASS_CTYPE_STREAM_OFR
            Exit Function
        End If
        If ctype = &H10700 Then
            GetCTypeString = "APE" ' BASS_CTYPE_STREAM_APE
            Exit Function
        End If
        If ctype = &H10A00 Then
            GetCTypeString = "MPC" ' BASS_CTYPE_STREAM_MPC
            Exit Function
        End If
        If ctype = &H10B00 Then
            GetCTypeString = "AAC" ' BASS_CTYPE_STREAM_AAC
            Exit Function
        End If
        If ctype = &H10B01 Then
            GetCTypeString = "MP4" ' BASS_CTYPE_STREAM_MP4
            Exit Function
        End If
        If ctype = &H10C00 Then
            GetCTypeString = "SPX" ' BASS_CTYPE_STREAM_SPX
            Exit Function
        End If
        If ctype = &H10E00 Then
            GetCTypeString = "ALAC" ' BASS_CTYPE_STREAM_ALAC
            Exit Function
        End If
        If ctype = &H10F00 Then
            GetCTypeString = "TTA" ' BASS_CTYPE_STREAM_TTA
            Exit Function
        End If
        If ctype = &H11000 Then
            GetCTypeString = "AC3" ' BASS_CTYPE_STREAM_AC3
            Exit Function
        End If
        GetCTypeString = "unknown add-on"
        Exit Function
    End If
    GetCTypeString = "?"
End Function

Private Sub cmdOpen_Click()
    cdOpen.CancelError = True
    cdOpen.flags = cdlOFNHideReadOnly Or cdlOFNExplorer
    
    On Error GoTo foo
    cdOpen.ShowOpen
    On Error GoTo 0
    
    If cdOpen.FileName = vbNullString Then
        Exit Sub
    End If
    
    BASS_StreamFree chan
    chan = BASS_StreamCreateFile(BASSFALSE, cdOpen.FileName, 0, 0, BASS_SAMPLE_LOOP)
    If chan = 0 Then
        cmdOpen.Caption = "click here to open a file..."
        ErrorM "Can't play the file"
        Exit Sub
    End If
    
    Dim info As BASS_CHANNELINFO
    BASS_ChannelGetInfo chan, info
    cmdOpen.Caption = cdOpen.FileName
    lblType.Caption = "channel type = " & LCase$(Hex(info.ctype)) & " (" & GetCTypeString(info.ctype) & ")"
    
    Dim time As Long
    time = CLng(BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetLength(chan)))
    sldPosition.max = time
    tmrPosition.Enabled = True
    BASS_ChannelPlay chan, BASSFALSE
foo:
End Sub

Private Sub Form_Load()
    If (Not BASS_GetVersion() = MakeLong(2, 2)) Then
        MsgBox "BASS version 2.2 was not loaded", vbOKOnly, "Incorrect BASS.DLL"
        End
    End If

    If BASS_Init(-1, 44100, 0, Me.hWnd, 0) = 0 Then
        ErrorM "Can't initialize device"
        End
    End If
    
    Dim fh As String
    fh = Dir$("bass*.dll")
    Do
        If BASS_PluginLoad(fh) Then
            lstPlugins.AddItem fh
        End If
        fh = Dir$()
    Loop While (Not fh = vbNullString)
    
    If lstPlugins.ListCount = 0 Then
        lstPlugins.AddItem "no plugins - visit the BASS webpage to get some"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BASS_Free
    BASS_PluginFree 0

    Set frmMain = Nothing
End Sub

Private Sub sldPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        blnSeek = True
    End If
End Sub

Private Sub sldPosition_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        blnSeek = False
    End If
End Sub

Private Sub sldPosition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        blnSeek = True
    End If
End Sub

Private Sub sldPosition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        blnSeek = False
    End If
End Sub

Private Sub sldPosition_Scroll()
    If blnSeek = True Then
        BASS_ChannelSetPosition chan, BASS_ChannelSeconds2Bytes(chan, sldPosition.value)
    End If
End Sub

Private Sub tmrPosition_Timer()
    sldPosition.value = CLng(BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan)))
End Sub
