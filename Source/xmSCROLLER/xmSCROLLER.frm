VERSION 5.00
Begin VB.Form xmSCROLLER 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "xmSCROfa"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5775
   Icon            =   "xmSCROLLER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xmSCROFe.ctrCredits ctrCredits1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
   End
End
Attribute VB_Name = "xmSCROLLER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim v_dx As DirectX7
Dim v_dmp As DirectMusicPerformance
Dim v_dml As DirectMusicLoader
Dim v_dms As DirectMusicSegment
Dim v_dmss As DirectMusicSegmentState

Private Sub Form_Load()
    Dim CreditFile As String
    Dim ConfigFile As String
    Dim CMD As String
    
    Me.ScaleMode = vbPixels
    
    CMD = Command$
    If Trim$(CMD) = "" Then
        CMD = App.EXEName
    End If
    
    ConfigFile = App.Path & "\" & CMD & ".ini"
    Me.Width = GetIni(ConfigFile, "MAIN", "WIDTH", 640) * Screen.TwipsPerPixelX
    Me.Height = GetIni(ConfigFile, "MAIN", "HEIGHT", 480) * Screen.TwipsPerPixelY
    
    ctrCredits1.Top = 0
    ctrCredits1.Left = 0
    ctrCredits1.Height = Me.ScaleHeight
    ctrCredits1.Width = Me.ScaleWidth
    

    CreditFile = GetIni(ConfigFile, "MAIN", "FILE", "GPL_E.TXT")
    If Dir(CreditFile) <> "" Then
        ctrCredits1.CreditFile = CreditFile
        Me.Caption = GetIni(ConfigFile, "MAIN", "CAPTION", "xmSCROFe")
        ctrCredits1.Intervallo = GetIni(ConfigFile, "MAIN", "INTERVAL", 1)
        Call PlayMid(GetIni(ConfigFile, "MAIN", "MID", "DOOMSONG.MID"))
        ctrCredits1.Start
    Else
        Call MsgBox("Credit File not found", vbCritical)
        End
    End If
    

End Sub

Private Sub PlayMid(MidFile As String)
On Error Resume Next
    
    Set v_dx = New DxVBLib.DirectX7
    
    Set v_dml = v_dx.DirectMusicLoaderCreate
    Set v_dmp = v_dx.DirectMusicPerformanceCreate
    
    Call v_dmp.Init(Nothing, hWnd)
    Call v_dmp.SetPort(-1, 1)
    
   
    Set v_dms = v_dml.LoadSegment(MidFile)
    
    v_dms.SetStandardMidiFile

    
    Call v_dmp.SetMasterAutoDownload(True)
    Call v_dms.Download(v_dmp)
    
    Set v_dmss = v_dmp.PlaySegment(v_dms, 0, 0)
    Call v_dmp.SetMasterVolume(10)

End Sub

