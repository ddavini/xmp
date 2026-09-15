VERSION 5.00
Begin VB.Form frmShowTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "X-MaD.Player"
   ClientHeight    =   150
   ClientLeft      =   -960
   ClientTop       =   -675
   ClientWidth     =   4680
   Icon            =   "ShowTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   150
   ScaleWidth      =   4680
End
Attribute VB_Name = "frmShowTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ShowMe As Boolean

Private Sub Form_Load()
    Me.Icon = xmp.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call DeLoad(True)
    End If
End Sub
