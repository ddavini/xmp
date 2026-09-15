VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         Height          =   2820
         Left            =   1320
         Picture         =   "frmAbout.frx":000C
         ScaleHeight     =   2760
         ScaleWidth      =   1860
         TabIndex        =   4
         Top             =   720
         Width           =   1920
      End
      Begin VB.Label lblMy 
         Alignment       =   2  'Center
         Caption         =   "xmINI Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Based on DMIni v1.0 Sample © 2001, DMSoftware"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   3840
         Width           =   3630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3000
         TabIndex        =   1
         Top             =   1320
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblMy.Caption = lblMy.Caption + " " + CStr(App.Major) + "." + CStr(App.Minor) _
                    + "." + CStr(App.Revision) + " " + App.Comments
    picLogo.Left = (Me.Width / 2) - (picLogo.Width / 2)
    lblMy.Left = (Me.Width / 2) - (lblMy.Width / 2)
End Sub

Private Sub Label4_Click()
ShellExecute hWnd, "open", "mailto:meelkertje@hotmail.com", vbNull, vbNull, SW_SHOWNORMAL
End Sub

Private Sub picLogo_Click()
    Call OpenUrl(Me.hWnd)
End Sub

Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call HandCursorOver(picLogo)
    picLogo.ToolTipText = "http://www.zolnetwork.com/x-mad/"
End Sub
