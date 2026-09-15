VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ECG (Secure PRNG for VB)"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " Random String: "
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   3615
      Begin VB.CommandButton cmdGenerate2 
         Caption         =   "&Generate"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRandomString 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox numLength 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Length:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Random Number: "
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox numLower 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox numUpper 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Text            =   "10000"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox numOut 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper:"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ECG As New clsECG

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub cmdGenerate_Click()
    If IsNumeric(numLower) = True And IsNumeric(numUpper) = True Then numOut = ECG.RandomNumber(numLower, numUpper)
End Sub

Private Sub cmdGenerate2_Click()
    If IsNumeric(numLength) = True Then txtRandomString.Text = ECG.RandomString(numLength)
End Sub

Private Sub Form_Load()
lngHwnd = GetWindowLong(numLower.hwnd, (-16))
lngReturn = SetWindowLong(numLower.hwnd, (-16), lngHwnd Or &H2000)
lngHwnd = GetWindowLong(numUpper.hwnd, (-16))
lngReturn = SetWindowLong(numUpper.hwnd, (-16), lngHwnd Or &H2000)
lngHwnd = GetWindowLong(numLength.hwnd, (-16))
lngReturn = SetWindowLong(numLength.hwnd, (-16), lngHwnd Or &H2000)
End Sub

Private Sub mnuAbout_Click()
    MsgBox "ECG (Ethereal Chaos Generator) is a strong, secure and" + Chr$(10) + "unpredictable pseudo-random number generator for Visual Basic." + Chr$(10) + Chr$(10) + "By: David Midkiff <mdj2023@hotmail.com>", vbOKOnly, "About ECG"
End Sub
Private Sub mnuExit_Click()
    End
End Sub

Private Sub numLength_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then If IsNumeric(numLength) = True Then txtRandomString.Text = ECG.RandomString(numLength)
End Sub
Private Sub numLower_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then If IsNumeric(numLower) = True And IsNumeric(numUpper) = True Then numOut = ECG.RandomNumber(numLower, numUpper)
End Sub
Private Sub numUpper_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then If IsNumeric(numLower) = True And IsNumeric(numUpper) = True Then numOut = ECG.RandomNumber(numLower, numUpper)
End Sub


