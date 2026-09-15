VERSION 5.00
Begin VB.Form frmNetRadio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BASS internet radio tuner"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framePresents 
      Caption         =   "Presents"
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkSave 
         Caption         =   "Save local copy"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "5"
         Height          =   350
         Index           =   9
         Left            =   3360
         TabIndex        =   9
         Top             =   630
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "4"
         Height          =   350
         Index           =   8
         Left            =   2790
         TabIndex        =   7
         Top             =   630
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "3"
         Height          =   350
         Index           =   7
         Left            =   2220
         TabIndex        =   5
         Top             =   630
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "2"
         Height          =   350
         Index           =   6
         Left            =   1650
         TabIndex        =   3
         Top             =   630
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "1"
         Height          =   350
         Index           =   5
         Left            =   1080
         TabIndex        =   1
         Top             =   630
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "5"
         Height          =   350
         Index           =   4
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "4"
         Height          =   350
         Index           =   3
         Left            =   2790
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "3"
         Height          =   350
         Index           =   2
         Left            =   2220
         TabIndex        =   4
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "2"
         Height          =   350
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton btnPresents 
         Caption         =   "1"
         Height          =   350
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblModem 
         AutoSize        =   -1  'True
         Caption         =   "Modem"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblBroadband 
         AutoSize        =   -1  'True
         Caption         =   "Broadband"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame framePlaying 
      Caption         =   "Currently playing"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   3975
      Begin VB.Label lblBPS 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1200
         Width           =   3795
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Caption         =   "not playing"
         Height          =   435
         Left            =   105
         TabIndex        =   16
         Top             =   720
         Width           =   3765
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSong 
         Alignment       =   2  'Center
         Height          =   435
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   3765
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmNetRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////////
' frmNetRadio.frm - Copyright (c) 2002-2005 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                         [http://www.jobnik.org]
'                                                         [  jobnik@jobnik.org  ]
'
' * Save local copy is added by: Peter Hebels @ http://www.phsoft.nl
'                                             e-mail: info@phsoft.nl
'
' Other sources: modNetRadio.bas & clsFileIo.cls
'
' BASS Internet radio example
' Originally translated from - netradio.c - Example of Ian Luck
'/////////////////////////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Sub Form_Load()
    'change and set the current path
    'so it won't ever tell you that "bass.dll" isn't found
    ChDrive App.Path
    ChDir App.Path

    'check if "bass.dll" is exists
    If (Not FileExists(RPP(App.Path) & "bass.dll")) Then
        Call MsgBox("BASS.DLL does not exists", vbCritical, "BASS.DLL")
        End
    End If

    'Check that BASS 2.2 was loaded
    If (BASS_GetVersion <> MakeLong(2, 2)) Then
        Call MsgBox("BASS version 2.2 was not loaded", vbCritical, "Incorrect BASS.DLL")
        End
    End If

    'setup output device
    If (BASS_Init(-1, 44100, 0, Me.hwnd, 0) = 0) Then
        Call Error_("Can't initialize device")
        End
    End If

    Call BASS_SetConfig(BASS_CONFIG_NET_PREBUF, 0) 'minimize automatic pre-buffering, so we can do it (and display it) instead

    'Stream URLs
    url = Array("http://64.236.34.196/stream/1048", "http://205.188.234.129:8024", _
                "http://64.236.34.97/stream/1006", "http://206.98.167.99:8406", _
                "http://160.79.1.141:8000", "http://streams.riverhosting.net:8012", _
                "http://205.188.234.4:8016", "http://205.188.234.4:8014", _
                "http://207.200.96.225:8010", "http://64.202.98.91:8082")

    Set WriteFile = New clsFileIo
    cthread = 0
End Sub

'this function will check if you're running in IDE or EXE modes
'VB will crash if you're closing the app while (cthread<>0) in IDE,
'but won't crash if in EXE mode
Public Function isIDEmode() As Boolean
    Dim sFileName As String, lCount As Long

    sFileName = String(255, 0)
    lCount = GetModuleFileName(App.hInstance, sFileName, 255)
    sFileName = UCase(GetFileName(Mid(sFileName, 1, lCount)))

    isIDEmode = (sFileName = "VB6.EXE")
End Function

Private Sub Form_Unload(Cancel As Integer)
    If (isIDEmode And cthread) Then
        'IDE Version
        Cancel = True   'disable closing app to avoid crash
    Else
        'Compiled Version or (cthread = 0) close app is available
        Call BASS_Free
    End If
End Sub

Private Sub btnPresents_Click(Index As Integer)
    If (cthread) Then   'already connecting
        Call Beep
    Else
        'open URL in a new thread (so that main thread is free)
        Dim threadid As Long
        cthread = CreateThread(ByVal 0&, 0, AddressOf OpenURL, Index, 0, threadid)   'threadid param required on win9x
    End If
End Sub

Private Sub chkSave_Click()
    If chkSave.value = vbChecked Then
        DoDownload = True
    Else
        DoDownload = False
    End If
End Sub

'--------------------------
' some useful functions :)
'--------------------------

'check if any file exists
Public Function FileExists(ByVal fp As String) As Boolean
    FileExists = (Dir(fp) <> "")
End Function

'RPP = Return Proper Path
Function RPP(ByVal fp As String) As String
    RPP = IIf(Mid(fp, Len(fp), 1) <> "\", fp & "\", fp)
End Function

'get file name from file path
Public Function GetFileName(ByVal fp As String) As String
    GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function

