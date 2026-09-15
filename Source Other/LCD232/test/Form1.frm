VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "login"
      Height          =   315
      Left            =   2550
      TabIndex        =   6
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "logout"
      Height          =   315
      Left            =   2550
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "senddata"
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "show UI"
      Height          =   285
      Left            =   1650
      TabIndex        =   3
      Top             =   75
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "senddata"
      Height          =   315
      Left            =   375
      TabIndex        =   2
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "logout"
      Height          =   315
      Left            =   375
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   315
      Left            =   375
      TabIndex        =   0
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Response:"
      Height          =   240
      Left            =   150
      TabIndex        =   10
      Top             =   1875
      Width           =   840
   End
   Begin VB.Label lblResp 
      Caption         =   "Label3"
      Height          =   240
      Left            =   1125
      TabIndex        =   9
      Top             =   1875
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3825
      TabIndex        =   8
      Top             =   750
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   7
      Top             =   750
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public obj As Object


Private Sub Command1_Click()
    Call obj.login(1, "testapp1", 120, "1122334455667788")
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command2_Click()
    Call obj.logout(1)
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command3_Click()
    Call obj.senddata(1, "ROWS=ciao ciao", 0)
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command4_Click()
    Call obj.sendcommand(1, "SHOWUI", 0)
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command5_Click()
    Call obj.senddata(2, "ROWS=bye bye|2  end.", 0)
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command6_Click()
    Call obj.logout(2)
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Command7_Click()
    Call obj.login(2, "testapp2", 60, "8877665544332211")
    
    lblResp.Caption = obj.response.sresponse
End Sub

Private Sub Form_Load()
    Set obj = CreateObject("LCD232.Interface")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set obj = Nothing
End Sub
