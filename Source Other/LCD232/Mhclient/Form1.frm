VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPORT 
      Height          =   285
      Left            =   3900
      TabIndex        =   9
      Text            =   "8080"
      Top             =   75
      Width           =   690
   End
   Begin VB.TextBox txtHOST 
      Height          =   285
      Left            =   2625
      TabIndex        =   8
      Text            =   "127.0.0.1"
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "hide UI"
      Height          =   315
      Left            =   2550
      TabIndex        =   7
      Top             =   900
      Width           =   1290
   End
   Begin VB.TextBox txtLOG 
      Height          =   1740
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   2100
      Width           =   4440
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      Height          =   315
      Left            =   1050
      TabIndex        =   5
      Top             =   1650
      Width           =   1065
   End
   Begin VB.TextBox txtDataType 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Text            =   "0"
      Top             =   1275
      Width           =   390
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "send data"
      Height          =   315
      Left            =   3450
      TabIndex        =   3
      Top             =   1275
      Width           =   1065
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   1275
      Width           =   1515
   End
   Begin VB.CommandButton cmdUI 
      Caption         =   "show UI"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   900
      Width           =   1290
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   525
      Width           =   1065
   End
   Begin MSWinsockLib.Winsock Winsoz 
      Left            =   150
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_Action As String
Public g_SESSIONID As Long


Private Sub cmdHide_Click()
    g_Action = "HIDEUI"
    Call Winsoz.Connect(txtHOST.Text, txtPORT.Text)
End Sub

Private Sub cmdLogin_Click()
    g_Action = "LOGIN"
    Call Winsoz.Connect(txtHOST.Text, txtPORT.Text)
End Sub

Private Sub cmdLogout_Click()
    g_Action = "LOGOUT"
    Call Winsoz.Connect(txtHOST.Text, txtPORT.Text)
End Sub

Private Sub cmdSend_Click()
    g_Action = "SENDDATA"
    Call Winsoz.Connect(txtHOST.Text, txtPORT.Text)
End Sub

Private Sub cmdUI_Click()
    g_Action = "SHOWUI"
    Call Winsoz.Connect(txtHOST.Text, txtPORT.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Winsoz_Close()
    Winsoz.Close
End Sub

Private Sub Winsoz_Connect()
    Select Case g_Action
        Case "LOGIN":
            Call Winsoz.SendData("GET /default.mhd?action=login&appname=MH test" & vbCrLf & vbCrLf)
            
        Case "SHOWUI":
            Call Winsoz.SendData("GET /default.mhd?action=sendcmd&sessionid=" & g_SESSIONID & "&data=SHOWUI&datatype=0" & vbCrLf & vbCrLf)
        
        Case "HIDEUI":
            Call Winsoz.SendData("GET /default.mhd?action=sendcmd&sessionid=" & g_SESSIONID & "&data=HIDEUI&datatype=0" & vbCrLf & vbCrLf)
            
        Case "SENDDATA":
            Call Winsoz.SendData("GET /default.mhd?action=senddata&sessionid=" & g_SESSIONID & "&data=" & txtData.Text & "&datatype=" & txtDataType.Text & vbCrLf & vbCrLf)
            
        Case "LOGOUT":
            Call Winsoz.SendData("GET /default.mhd?action=logout&sessionid=" & g_SESSIONID & vbCrLf & vbCrLf)
    End Select
End Sub

Private Sub Winsoz_DataArrival(ByVal bytesTotal As Long)
    Dim buff As String
    Dim vStart As Long, vEnd As Long
    
    If bytesTotal > 0 Then
        Call Winsoz.GetData(buff, vbString)
        
        buff = Mid$(buff, InStr(buff, vbCrLf & vbCrLf) + 4)
        txtLOG.Text = buff
        
        Select Case True
            Case InStr(1, buff, "Logged in"):
                vStart = InStr(1, buff, "sessionid=") + 10
                vEnd = InStr(vStart, buff, vbCrLf)
                g_SESSIONID = Mid$(buff, vStart, vEnd - vStart)
                
                Me.Cls
                Me.Print g_SESSIONID
        End Select
    End If
End Sub

