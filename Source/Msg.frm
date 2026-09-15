VERSION 5.00
Begin VB.Form frmMessage 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4305
   ClipControls    =   0   'False
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
   Icon            =   "Msg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txError 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Msg.frx":000C
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   0
      Width           =   0
   End
   Begin VB.Label cmdButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label cmdButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label cmdButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Based on Sample of:
' Andreas Schwarz, FutureProjects Development
' e-mail: andi@futureprojects.de , http://www.futureprojects.de
Option Explicit

Private ButtonResult As MsgResult

Public Enum MsgType
    MsgError = 16
    MsgAsk = 32
    MsgWarning = 48
    MsgInfo = 64
    
    MsgOkOnly = 0
    MsgOkCancel = 1
    MsgOkIgnore = 2
    MsgYesNo = 3
    MsgYesNoCancel = 4
End Enum
    
Public Enum MsgResult
    MsgOK = vbOK
    MsgCancel = vbCancel
    MsgIgnore = vbIgnore
    MsgYes = vbYes
    MsgNo = vbNo
End Enum


Function dwMessageBox(Prompt As String, Optional Flags As MsgType = 0, Optional PromptCaption As String, Optional iForm As Form) As MsgResult
        
    Dim hIcon As Long
    
    'Display Icon
    If Flags < 16 Then 'NoIcon
            
    ElseIf Flags < 32 Then 'Error-Icon
           
        hIcon = LoadStandardIcon(0&, IDI_HAND)
        Call DrawIcon(Me.hdc, 5&, 5&, hIcon)
        Flags = Flags - 16
        
    ElseIf Flags < 48 Then 'Ask-Icon
        
        hIcon = LoadStandardIcon(0&, IDI_QUESTION)
        Call DrawIcon(Me.hdc, 5&, 5&, hIcon)
        Flags = Flags - 32
        
    ElseIf Flags < 64 Then 'Warning-Icon
    
        hIcon = LoadStandardIcon(0&, IDI_EXCLAMATION)
        Call DrawIcon(Me.hdc, 5&, 5&, hIcon)
        Flags = Flags - 48
        
    ElseIf Flags < 80 Then 'Info-Icon
    
        hIcon = LoadStandardIcon(0&, IDI_ASTERISK)
        Call DrawIcon(Me.hdc, 5&, 5&, hIcon)
        Flags = Flags - 64
        
    End If
    
    'Displaytext
    txError.Text = Prompt
    txError.SelStart = Len(txError.Text)
    Caption = PromptCaption
           
    'Initialize Buttons
    
    Select Case Flags
        Case 0
            SetButton 0, "OK"
        Case 1
            SetButton 0, "OK"
            SetButton 1, "Apply"
        Case 2
            SetButton 0, "OK"
            SetButton 1, "Ignore"
        Case 3
            SetButton 0, "Yes"
            SetButton 1, "No"
        Case 4
            SetButton 0, "Yes"
            SetButton 1, "No"
            SetButton 2, "Cancel"
    End Select

   
    SempreInPrimoPiano frmMessage.hWnd, frmMessage.Width, frmMessage.Height, _
                frmMessage.Left, frmMessage.Top, frmMenu.InPrimoPiano
    frmMessage.Show 1, iForm
    
    dwMessageBox = ButtonResult
    
    Unload Me
   
End Function
Private Sub SetButton(ButtonIndex As Integer, PromptText As String)
    cmdButton(ButtonIndex).Caption = PromptText
    cmdButton(ButtonIndex).Visible = True
End Sub

Private Sub cmdButton_Click(Index As Integer)
Select Case cmdButton(Index).Caption
    Case "Ok"
        ButtonResult = MsgOK
    Case "Cancel"
        ButtonResult = MsgCancel
    Case "Ignore"
        ButtonResult = MsgIgnore
    Case "Yes"
        ButtonResult = MsgYes
    Case "No"
        ButtonResult = MsgNo
End Select
Me.Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
               
End Sub

Private Sub Form_Load()
    Dim T As Integer
    For T = 0 To cmdButton.count - 1
        cmdButton(T).Visible = False
    Next T
End Sub

