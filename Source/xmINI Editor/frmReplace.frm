VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "R&eplace"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtReplaceWith 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtTextToFind 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Case sensitive"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F&ind Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Replace With:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Find what:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmReplace"
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

Private Sub Command3_Click()
frmMain.txtIni.SelText = txtReplaceWith.Text
End Sub

Private Sub Command4_Click()
Dim FoundPos
Dim Counter As Integer

If Not txtTextToFind.Text = LastSearchedText Then
    LastSearchedText = txtTextToFind.Text
    LastPos = 1
End If

Counter = 0

Do

If Check1.Value = 1 Then
    FoundPos = InStr(LastPos, frmMain.txtIni.Text, txtTextToFind.Text)
Else
    FoundPos = InStr(LastPos, frmMain.txtIni.Text, txtTextToFind.Text, vbTextCompare)
End If

If FoundPos = 0 Then Exit Do

LastPos = FoundPos + Len(txtTextToFind.Text) - 1
frmMain.txtIni.SelStart = FoundPos - 1
frmMain.txtIni.SelLength = Len(txtTextToFind.Text)
frmMain.txtIni.SelText = txtReplaceWith.Text

Counter = Counter + 1

Loop Until FoundPos = 0

frmMain.txtIni.SetFocus

MsgBox "Done replacing, " & Counter & " replacements made", vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
SetWindowWord frmReplace.hWnd, SWW_HPARENT, frmMain.hWnd
txtTextToFind.Text = frmMain.txtIni.SelText
End Sub

