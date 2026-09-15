VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LCD232"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm ComCtrl 
      Left            =   3960
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtData2 
      Height          =   1440
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain.frx":0000
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtLCD 
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4125
      Width           =   4215
   End
   Begin VB.TextBox txtLCD 
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3750
      Width           =   4215
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   225
      Width           =   4215
   End
   Begin VB.ListBox lstApps 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   150
      TabIndex        =   0
      Top             =   2475
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Me.Show
End Sub
