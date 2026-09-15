VERSION 5.00
Begin VB.Form frmTagEdit 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tag Edit"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6060
   Icon            =   "frmTagEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      ToolTipText     =   "Update the ID3v1 tag"
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Frame fraID3v1 
      BackColor       =   &H00000000&
      Caption         =   "ID3v1 Tag"
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
      Height          =   3195
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   6045
      Begin VB.TextBox txtTrack 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   15
         Top             =   210
         Width           =   560
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2160
         Width           =   4185
      End
      Begin VB.TextBox txtYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1770
         Width           =   4185
      End
      Begin VB.TextBox txtAlbum 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   3
         Top             =   990
         Width           =   4185
      End
      Begin VB.TextBox txtArtist 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   2
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtSongName 
         Height          =   315
         Left            =   1770
         MaxLength       =   30
         TabIndex        =   1
         Top             =   210
         Width           =   2865
      End
      Begin VB.CommandButton cmdUpdateID3v1tag 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Write Tag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         ToolTipText     =   "Update the ID3v1 tag"
         Top             =   2640
         Width           =   1125
      End
      Begin VB.ComboBox cboGenreName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1380
         Width           =   4185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Track #:"
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Song:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Genre Name:"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.CheckBox chkv1 
      BackColor       =   &H00000000&
      Caption         =   "Has ID3v1"
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
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Checked = contains ID3v1 info"
      Top             =   0
      Width           =   1605
   End
End
Attribute VB_Name = "frmTagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_GetMP3Info As modInfoMp3.Mp3Info
Public NomeFile As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUpdateID3v1tag_Click()
        Call WriteTag(NomeFile, txtSongName, txtArtist, txtAlbum, txtYear, txtComment, cboGenreName.ListIndex)
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
Dim I As Integer

        For I = 0 To 146
            cboGenreName.AddItem GenreText(I)
        Next
        
        m_GetMP3Info = ReadMP3(NomeFile, True, True, gb_Data)
        If m_GetMP3Info.HasTag Then
            chkv1.Value = 1
            With m_GetMP3Info
                txtSongName.Text = .SongName
                txtYear.Text = .Year
                txtAlbum.Text = .Album
                If .Genre <> 255 Then
                    cboGenreName.ListIndex = .Genre
                End If
                txtComment.Text = .Comment
                txtArtist.Text = .Artist
                txtTrack = .track
            End With
        End If
        
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                "frmEditTag.Form_load" & vbCrLf & Err.Description)
    End If
End Sub
