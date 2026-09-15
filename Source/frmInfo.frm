VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Information"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6510
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox lstInfo 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
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
      Height          =   3510
      Left            =   0
      ScaleHeight     =   3450
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   0
      Width           =   5895
   End
   Begin VB.PictureBox picLeft 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   3510
      Left            =   5895
      ScaleHeight     =   3480
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   0
      Width           =   615
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   50
         TabIndex        =   1
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblCPU 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   590
      End
   End
   Begin VB.Timer tmrCPU 
      Left            =   4080
      Top             =   120
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CPU As New CPULoad

Private Sub cmdChiudi_Click()
    frmInfo.tmrCPU.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    frmInfo.tmrCPU.Interval = 1000
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmrCPU.Interval = 0
End Sub

Private Sub tmrCPU_Timer()
On Error GoTo errH
Dim iCPUl As Long
Dim iCPU As Long

   m_CPU.CollectCPUData
    
    For iCPU = 0 To m_CPU.GetCPUCount - 1
        iCPUl = m_CPU.GetCPUUsage(iCPU + 1)
        If iCPU > 0 Then
            If lblCPU.count < m_CPU.GetCPUCount Then
                Load lblCPU(iCPU)
                lblCPU(iCPU).Top = lblCPU(iCPU - 1).Top + lblCPU(iCPU - 1).Height
                lblCPU(iCPU).Visible = True
            End If
        End If
        lblCPU(iCPU).Caption = Format(iCPUl, "0") & "%"
        lblCPU(iCPU).ToolTipText = "CPU" & iCPU + 1
    Next
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmInfo.tmrCPU_Timer" & vbCrLf & Err.Description)
        tmrCPU.Enabled = False
    End If
End Sub
