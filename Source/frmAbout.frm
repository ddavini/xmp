VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About XmP"
   ClientHeight    =   5145
   ClientLeft      =   2040
   ClientTop       =   1890
   ClientWidth     =   8010
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XmP1.ctrCredits ctrCredits 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
   End
   Begin VB.PictureBox picLblCotainer 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Image ImgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   1
         Left            =   6480
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image ImgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   2
         Left            =   5880
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image ImgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   0
         Left            =   5160
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image ImgIcon 
         Height          =   1335
         Left            =   3600
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00000000&
         Caption         =   "Version"
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
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3285
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00000000&
         Caption         =   "Application Title"
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
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00000000&
         Caption         =   "App Description"
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
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   3195
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   8040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   7920
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private X1 As Integer, Y1 As Integer
Private Mouvement As Boolean

Private Sub Form_Load()
    Dim FileBkg As String
    
    'Carica Res
    ImgLogo(0).Picture = xmp.ImgLogo.Picture
    ImgLogo(1).Picture = LoadDataIntoFile("LOGOAPL")
    ImgLogo(2).Picture = LoadDataIntoFile("LOGOREGFREE")
         
    ImgIcon.Picture = LoadDataIntoFile("LOGOGSSJ2CPRS")
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = lblVersion.Caption + " " + App.Comments
    lblTitle.Caption = App.Title
    lblDescription.Caption = "X-MaD.Player The S Player coded by MaD" + _
                vbCrLf + "Tested with Project Prometeo MP3s" + vbCrLf
    Call ReadFile
    ctrCredits.Intervallo = 60
    ctrCredits.Start
End Sub

Private Sub ReadFile(Optional FileName As String = "CREDITS.TXT")

    Dim StrStore As String
    Dim Filez As String
    Dim FileNumber As Integer

On Error Resume Next
        ChDir App.Path
        Filez = ".\" + FileName
        ctrCredits.CreditFile = Filez

End Sub

Private Sub ImgIcon_Click()
    Call OpenUrl(Me.hWnd, "mailto:x-mad@zolnetwork.com")
End Sub

Private Sub ImgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call HandCursorOver(ImgIcon)
    ImgIcon.ToolTipText = "Mail Me"
End Sub

Private Sub ImgLogo_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            Call OpenUrl(Me.hWnd)
        Case Is = 1
            Call OpenUrl(Me.hWnd, "http://www.zolnetwork.com/x-mad/xmBinary/default.asp?File=APLRuntime.rar")
    End Select
End Sub

Private Sub ImgLogo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case Is = 0
            Call HandCursorOver(ImgLogo(Index))
            ImgLogo(Index).ToolTipText = "http://www.zolnetwork.com/x-mad/"
        Case Is = 1
            Call HandCursorOver(ImgLogo(Index))
            ImgLogo(Index).ToolTipText = "AP&L 2.0 Compilant - Click to Download Last AP&L Pack"
        Case Is = 2
            ImgLogo(Index).ToolTipText = "This Program is Registry Free"
    End Select
End Sub
