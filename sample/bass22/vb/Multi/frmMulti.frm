VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMulti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BASS multiple output example"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameMulti 
      Caption         =   " device 1 "
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdOpen 
         Caption         =   "click here to open a file..."
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameMulti 
      Caption         =   " device 2 "
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      Begin VB.CommandButton cmdOpen 
         Caption         =   "click here to open a file..."
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
' frmMulti.frm - Copyright (c) 2003-2005 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                      [http://www.jobnik.org]
'                                                      [  jobnik@jobnik.org  ]
' Other source: frmDevice.frm
'
' BASS Multiple output example
' Originally translated from - multi.c - Example of Ian Luck
'//////////////////////////////////////////////////////////////////////////////
 
Option Explicit

Dim outdev(2) As Long   'output devices
Dim chan(2) As Long     'the streams

'display error messages
Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & vbCrLf & "Error Code : " & BASS_ErrorGetCode, vbExclamation, "Error")
End Sub

Private Sub Form_Load()
    'change and set the current path
    'so VB won't ever tell you, that "bass.dll" isn't found
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

    'Let the user choose the output devices
    With frmDevice
        .SelectDevice 1
        .Show vbModal, Me
        outdev(0) = .device
        .SelectDevice 2
        .Show vbModal, Me
        outdev(1) = .device
    End With

    'setup output devices
    If (BASS_Init(outdev(0), 44100, 0, Me.hWnd, 0) = 0) Then
        Call Error_("Can't initialize device 1")
        Unload Me
    End If

    If (BASS_Init(outdev(1), 44100, 0, Me.hWnd, 0) = 0) Then
        Call Error_("Can't initialize device 2")
        Unload Me
    End If

    frameMulti(0).Caption = " " & VBStrFromAnsiPtr(BASS_GetDeviceDescription(outdev(0))) & " "
    frameMulti(1).Caption = " " & VBStrFromAnsiPtr(BASS_GetDeviceDescription(outdev(1))) & " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'release both devices
    Call BASS_SetDevice(outdev(0))
    Call BASS_Free
    Call BASS_SetDevice(outdev(1))
    Call BASS_Free
    End
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    On Local Error Resume Next    'if Cancel pressed...

    'open a file to play on selected device
    cmd.CancelError = True
    cmd.flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    cmd.DialogTitle = "Open"
    cmd.Filter = "streamable files|*.mp3;*.mp2;*.mp1;*.ogg;*.wav;*.aif|All files|*.*"
    cmd.ShowOpen

    'if cancel was pressed, exit the procedure
    If Err.Number = 32755 Then Exit Sub

    Call BASS_StreamFree(chan(Index))
    Call BASS_SetDevice(outdev(Index)) 'set the device to create stream on

    chan(Index) = BASS_StreamCreateFile(BASSFALSE, cmd.FileName, 0, 0, BASS_SAMPLE_LOOP)

    If (chan(Index) = 0) Then
        cmdOpen(Index).Caption = "click here to open a file..."
        Call Error_("Can't play the file")
        Exit Sub
    End If

    cmdOpen(Index).Caption = cmd.FileName
    Call BASS_ChannelPlay(chan(Index), BASSFALSE)
End Sub

'--------------------------
' some useful functions :)
'--------------------------

'check if any file exists
Public Function FileExists(ByVal fp As String) As Boolean
  FileExists = (Dir(fp) <> "")
End Function

' RPP = Return Proper Path
Function RPP(ByVal fp As String) As String
    RPP = IIf(Mid(fp, Len(fp), 1) <> "\", fp & "\", fp)
End Function
