VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F&ind Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Case sensitive"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtTextToFind 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Find what:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastPos As Variant
Dim LastSearchedText As String

Private Sub Command1_Click()
Dim FoundPos
If Not txtTextToFind.Text = LastSearchedText Then
    LastSearchedText = txtTextToFind.Text
    LastPos = 1
End If

If Check1.Value = 1 Then
    FoundPos = InStr(LastPos, frmMain.txtIni.Text, txtTextToFind.Text)
Else
    FoundPos = InStr(LastPos, frmMain.txtIni.Text, txtTextToFind.Text, vbTextCompare)
End If

If FoundPos = 0 Then
    MsgBox "Sorry, string not found", vbExclamation, Me.Caption
    Exit Sub
End If
LastPos = FoundPos + Len(txtTextToFind.Text) - 1
frmMain.txtIni.SelStart = FoundPos - 1
frmMain.txtIni.SelLength = Len(txtTextToFind.Text)
frmMain.txtIni.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetWindowWord frmFind.hWnd, SWW_HPARENT, frmMain.hWnd
End Sub
