VERSION 5.00
Begin VB.Form frmBassEngine 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BASSEngine"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBass 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4200
      Top             =   120
   End
   Begin VB.Frame frameStream 
      Caption         =   "Stream"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.CommandButton cmdWriteLs 
         Caption         =   "LSGen"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStreamStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStreamPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label SONG 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBassEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private StreamHandle As Long

Private Type State
    Play As Boolean
    Pause As Boolean
End Type

Private StreamState As State

Public Sub cmdNext_Click()
    If CBool(GetIni("", "MAIN", "RANDOM", "False")) Then
        Call RNDGen
    Else
        LSIndex = LSIndex + 1
    End If
    If LSIndex > NumeroMp3 Then
        LSIndex = 0
    End If
    Stream_GO
End Sub

Public Sub cmdPause_Click()
    StreamState.Pause = Not StreamState.Pause
    If StreamState.Pause Then
        Call BASS_Pause
    Else
        Call BASS_Start
    End If
End Sub

Public Sub cmdStreamPlay_Click()
    Call Stream_GO
End Sub

Public Sub cmdStreamStop_Click()
    Call StreamStop
End Sub

Public Sub cmdWriteLs_Click()
    Call CollectMP3(GetIni("", "MAIN", "MP3ROOT", "D:\MP3\"), App.path & "\mp3.ls")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    Else
        Call frmHotKey(KeyAscii)
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrH

    LSIndex = 0
    
    If App.PrevInstance Then
        End
    End If
    
    Call DisplayError("Bass engine Start", "Information", "Form_Load")
    
    'change and set the current path
    'so it won't ever tell you that bass.dll is not found
    ChDrive App.path
    ChDir App.path
        
    If CBool(GetIni("", "MAIN", "RANDOM", "False")) Then
        Call InitRND
        Call LoadRNDDat(NumeroMp3)
    Else
        Call KillRNDFile
        NumeroMp3 = ReadMp3Number(App.path & "\mp3.ls")
    End If
    
    LSIndex = ReadMp3Last(App.path & "\mp3.ls")

    'check if 'bass.dll' is exists
    If Not FileExists(RPP(App.path) & "bass.dll") Then
        Call DisplayError("BASS.DLL does not exists", "Critical", "BASS.DLL")
        End
    End If

    'Check that BASS 2.2 was loaded
    If BASS_GetVersion <> MakeLong(2, 2) Then
        Call DisplayError("BASS version 2.2 was not loaded", "Critical", "Incorrect BASS.DLL")
        End
    End If

    'Initialize output - default device, 44100hz, stereo, 16 bits
    If BASS_Init(-1, 44100, 0, Me.hWnd, 0) = BASSFALSE Then
        Call DisplayError("BASS Init Faild", "Critical", "")
        End
    End If

    'Start the timer
    tmrBass.Enabled = True
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & "::" & Err.Source & _
                "::" & "frmBassEngine.Form_load" & "::" & Err.Description)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'stop timer
    tmrBass.Enabled = False

    'Close sound system and release everything
    Call BASS_StreamFree(StreamHandle)
    Call BASS_Free
    
    Call DisplayError("Bass engine Stop", "Information", "Form_Unload")
End Sub


Private Sub Stream_GO()
On Error GoTo ErrH
    Dim FileToPlay() As String
    
    Call BASS_Start
    Call StreamStop
    StreamState.Pause = False
    
    Call ScriviPosizioneListaMp3(App.path & "\mp3.ls")
    
    If CBool(GetIni("", "MAIN", "RANDOM", "False")) Then
        Call WriteRNDFile(LSIndex)
    Else
        Call KillRNDFile
    End If
    
    FileToPlay = Split(ReadMp3List(App.path & "\mp3.ls", LSIndex), ";")
    
    If UCase$(GetIni("", "MAIN", "LOG", "")) = "VERBOSE" Then
        Call DisplayError("** " & FileToPlay(0) & " ** file will be be played! [Stream_GO]")
    End If
    
    StreamHandle = BASS_StreamCreateFile(BASSFALSE, FileToPlay(0), 0, 0, 0)

    If StreamHandle = 0 Then
        StreamHandle = BASS_MusicLoad(BASSFALSE, FileToPlay(0), 0, 0, BASS_MUSIC_RAMP, 0)
    End If

    If StreamHandle = 0 Then
        Call DisplayError("** " & FileToPlay(0) & " ** file couldn't be played! [BASS_MusicLoad]")
        'StreamState.Play = False
    End If
    
    If BASS_ChannelPlay(StreamHandle, BASSFALSE) = 1 Then
        'StreamState.Play = True
    Else
        Call DisplayError("** " & FileToPlay(0) & " ** file couldn't be played! [BASS_ChannelPlay]")
        'StreamState.Play = False
    End If
    
    StreamState.Play = True
    
   ' Call Sleep(5000)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & "::" & Err.Description, "Runtime Error", "frmBassEngine.Stream_GO")
    End If
End Sub

Private Sub StreamStop()
    Call BASS_ChannelStop(StreamHandle)
    StreamState.Play = False
End Sub

Private Sub tmrBass_Timer()
    If BASS_ChannelIsActive(StreamHandle) = 0 Then
        If StreamState.Play And Not StreamState.Pause Then
            If Not LSRestart Then
                If CBool(GetIni("", "MAIN", "RANDOM", "False")) Then
                    Call RNDGen
                Else
                    LSIndex = LSIndex + 1
                End If
            Else
                LSIndex = 0
                LSRestart = False
            End If
            Stream_GO
        Else
            cmdWriteLs.Enabled = True
        End If
    Else
        cmdWriteLs.Enabled = False
    End If
End Sub

Private Sub RNDGen()
     ''' RANDOM '''
    
     Randomize
     
     Call InitRND
     
     LSIndex = Int((NumeroMp3 * Rnd))
     
     Dim Forward As Boolean: Forward = True
     Dim IndexStore As Long: IndexStore = LSIndex
     
     While vRND(LSIndex)
         If LSIndex = NumeroMp3 Then
             Forward = False
             LSIndex = IndexStore
         End If
         
         If Forward Then
             LSIndex = LSIndex + 1
         Else
             LSIndex = LSIndex - 1
         End If
         
         If LSIndex <= 0 Then
             Call CreateRNDFile(NumeroMp3)
         End If
     Wend
     
     'Call WriteRNDFile(LSIndex)
     
     ''' RANDOM '''
End Sub


