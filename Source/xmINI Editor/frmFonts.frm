VERSION 5.00
Begin VB.Form frmFonts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox picSection 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3225
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkSectionItalic 
         Caption         =   "&Italic"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkSectionBold 
         Caption         =   "&Bold"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Example"
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   5535
         Begin VB.Label lblSectionsTest 
            Caption         =   "Kame......ame.....ame..........aaaaa!!!!!!!!!!"
            Height          =   975
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.TextBox txtSectionSize 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbSectionFont 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Font &Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Font name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEditBold_Click()
UpdateTest
End Sub

Private Sub chkEditItalic_Click()
UpdateTest
End Sub

Private Sub chkSectionBold_Click()
UpdateTest
End Sub

Private Sub chkSectionItalic_Click()
UpdateTest
End Sub

Private Sub cmbEditFont_Click()
UpdateTest
End Sub

Private Sub cmbSectionFont_Click()
UpdateTest
End Sub

Private Sub Command1_Click()
With frmMain.lstSections.Font
    .Name = lblSectionsTest.FontName
    .Size = lblSectionsTest.FontSize
    .Bold = lblSectionsTest.FontBold
    .Italic = lblSectionsTest.FontItalic
End With

With frmMain.txtIni.Font
    .Name = lblSectionsTest.FontName
    .Size = lblSectionsTest.FontSize
    .Bold = lblSectionsTest.FontBold
    .Italic = lblSectionsTest.FontItalic
End With

frmMain.ResizeControls

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 0 To Screen.FontCount - 1
    cmbSectionFont.AddItem Screen.Fonts(I)
Next I

cmbSectionFont.Text = frmMain.lstSections.FontName

txtSectionSize.Text = frmMain.lstSections.FontSize

chkSectionBold.Value = IIf(frmMain.lstSections.FontBold, 1, 0)
chkSectionItalic.Value = IIf(frmMain.lstSections.FontItalic, 1, 0)

UpdateTest
End Sub

Private Sub txtEditSize_Change()
UpdateTest
End Sub

Private Sub txtSectionSize_Change()
UpdateTest
End Sub

Public Function UpdateTest()
On Error Resume Next
With lblSectionsTest
    .FontName = cmbSectionFont.Text
    .FontSize = txtSectionSize.Text
    .FontBold = chkSectionBold.Value
    .FontItalic = chkSectionItalic.Value
End With
End Function
