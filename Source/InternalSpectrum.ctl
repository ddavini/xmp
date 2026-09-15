VERSION 5.00
Begin VB.UserControl SpectrumCtrl 
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   FillColor       =   &H00FFFFFF&
   LockControls    =   -1  'True
   ScaleHeight     =   2036.364
   ScaleMode       =   0  'User
   ScaleWidth      =   3690
   Begin VB.PictureBox ScancellatorePic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   619
      Left            =   120
      ScaleHeight     =   496.97
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox SfondoPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   619
      Left            =   0
      ScaleHeight     =   315.385
      ScaleMode       =   0  'User
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   60
   End
End
Attribute VB_Name = "SpectrumCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    
    SfondoPic.Top = 0
    SfondoPic.Left = 0
    SfondoPic.Height = UserControl.Height
    SfondoPic.Width = UserControl.Width
    ScancellatorePic.Height = SfondoPic.Height
    ScancellatorePic.Width = SfondoPic.Width
    ScancellatorePic.Top = SfondoPic.Top
    ScancellatorePic.Left = SfondoPic.Left
    
End Sub

Public Sub Modello(ByVal Pic As IPictureDisp)
    SfondoPic.Picture = Pic
End Sub

Public Sub WriteSpec(Valore As Integer, Max As Byte)
    Dim Appoggio As Integer
On Error GoTo ErrH
    Appoggio = Abs(UserControl.ScaleHeight - ((UserControl.ScaleHeight / Max) * Valore))
    If Appoggio > (UserControl.ScaleHeight - 10) And Valore >= 1 Then
        Appoggio = UserControl.ScaleHeight - 10
    End If
    If Appoggio <= 0 Or Valore >= Max Then
        Appoggio = 0
    End If
    ScancellatorePic.Height = Appoggio
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Spectrum.WriteSpec" & vbCrLf & Err.Description)
    End If
End Sub
