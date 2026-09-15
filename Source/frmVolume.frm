VERSION 5.00
Begin VB.Form frmVolume 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4155
   Icon            =   "frmVolume.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox LED 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Timer tmrHide 
      Left            =   1200
      Top             =   240
   End
   Begin VB.PictureBox PercValLed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    If xmp.WindowState = vbNormal Then
        xmp.SetFocus
    End If
End Sub

Private Sub Form_Load()
Dim I As Byte
On Error GoTo errH

    For I = 1 To 16
        Load LED(I)
    Next I
    
    Call LoadGFX(LED)
    
    Me.Height = 360
    Me.Width = 720
    
    PercValLed.Top = (Me.Height \ 2) - (PercValLed.Height \ 2)
    PercValLed.Left = 0
    
    Me.BackColor = vbBlack
    
    Me.Top = GetIni(cfgFile, "VISUALIZATION", "YPOSITION", "100")
    Me.Left = GetIni(cfgFile, "VISUALIZATION", "XPOSITION", "100")
    Me.Visible = GetIni(cfgFile, "VISUALIZATION", "VVENABLE", "False")
    Call RoundForm(Me)
    Call SempreInPrimoPiano(Me.hWnd, Me.Width, Me.Height, Me.Top, Me.Left)
    
    tmrHide.Interval = GetIni(cfgFile, "VISUALIZATION", "HIDEIDLE", "5000")
    
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmVolume.Load" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub tmrHide_Timer()
    xmp.xmDVol.Visible = False
    Unload Me
End Sub
