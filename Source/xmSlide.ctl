VERSION 5.00
Begin VB.UserControl xmSlide 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ForwardFocus    =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2130
   ScaleWidth      =   1275
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   240
   End
   Begin VB.VScrollBar vsGhost 
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picSlide 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox picSfondoSlide 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line lnSlide2 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   1200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line lnSlide 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   1200
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "xmSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private YSlide1 As Integer
Private MouvementSlide As Boolean
Private MouseUp As Boolean

'Event Change(Index As Integer)
Event Change()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Private Sub CommandImg_Click(Index As Integer)
'    If MouseUp Then
'        Select Case Index
'            Case Is = 1 'Giu'
'                If vsGhost.Value < vsGhost.Max Then
'                    vsGhost.Value = vsGhost.Value + 1
'                End If
'            Case Is = 0 'Su
'                If vsGhost.Value > vsGhost.Min Then
'                    vsGhost.Value = vsGhost.Value - 1
'                End If
'        End Select
'    End If
'End Sub

Private Sub CommandImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P As Single
        Call xIlluminaPulsante(CommandImg(Index), True)
        MouseUp = False
        Select Case Index
            Case Is = 1 'Giu'
                While Not MouseUp
                    If vsGhost.Value < vsGhost.Max Then
                        vsGhost.Value = vsGhost.Value + 1
                    Else
                        vsGhost.Value = vsGhost.Max
                        MouseUp = True
                    End If
                    P = P + (GetIni(cfgFile, "PREFERENCE", "SLIDEACCEL", "1") / 1000)
                    Call Pause(0.2 - P, True)
                Wend
            Case Is = 0 'Su
                Do
                    If vsGhost.Value > vsGhost.Min Then
                        vsGhost.Value = vsGhost.Value - 1
                    Else
                        vsGhost.Value = vsGhost.Min
                        MouseUp = True
                    End If
                    P = P + (GetIni(cfgFile, "PREFERENCE", "SLIDEACCEL", "1") / 1000)
                    Call Pause(0.2 - P, True)
                Loop Until MouseUp
        End Select
End Sub

Private Sub CommandImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Pause(0.3, True)
    Call xIlluminaPulsante(CommandImg(Index), False)
    MouseUp = True
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub xIlluminaPulsante(CommandImg As PictureBox, _
                            Optional LockButton As Boolean = True, _
                            Optional Colore As Long = vbGreen)
On Error GoTo ErrH
Dim sX As Integer
Dim sY As Integer
Dim sXTwips As Integer
Dim sYTwips As Integer
    If LockButton = False Then
         Colore = UserControl.BackColor
         sXTwips = Screen.TwipsPerPixelX
         sYTwips = Screen.TwipsPerPixelY
         sX = CommandImg.Width - (sXTwips * 3)
         sY = CommandImg.Height - (sYTwips * 3)
         CommandImg.AutoRedraw = True
         CommandImg.Line (0, 0)-(0, sY), Colore
         CommandImg.Line -(sX, sY), Colore
         CommandImg.Line -(sX, 0), Colore
         CommandImg.Line -(0, 0), Colore
         CommandImg.PSet (sXTwips, sYTwips), Colore
         CommandImg.PSet (sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sYTwips), Colore
         CommandImg.AutoRedraw = False
     Else
         sXTwips = Screen.TwipsPerPixelX
         sYTwips = Screen.TwipsPerPixelY
         sX = CommandImg.Width - (sXTwips * 3)
         sY = CommandImg.Height - (sYTwips * 3)
         CommandImg.AutoRedraw = True
         CommandImg.Line (0, 0)-(0, sY), Colore
         CommandImg.Line -(sX, sY), Colore
         CommandImg.Line -(sX, 0), Colore
         CommandImg.Line -(0, 0), Colore
         CommandImg.PSet (sXTwips, sYTwips), Colore
         CommandImg.PSet (sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sYTwips), Colore
         CommandImg.AutoRedraw = False
     End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.IlluminaPulsante" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub picSlide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lnSlide.BorderColor = vbGreen
    lnSlide2.BorderColor = vbGreen
End Sub

Private Sub picSlide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouvementSlide = False
    lnSlide.BorderColor = RGB(206, 206, 206)
    lnSlide2.BorderColor = RGB(206, 206, 206)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    UserControl.BackColor = vbBlack
    
    vsGhost.Left = UserControl.Width + vsGhost.Width + 100
       
   ' CommandImg(0).Width = UserControl.Width
    CommandImg(0).Align = vbAlignTop
    CommandImg(0).AutoRedraw = True
    'CommandImg(0).BackColor = vbWhite
    Call CommandImg(0).PaintPicture(LoadResPicture("FRECCIAUP", vbResBitmap), UserControl.Width / 2 - _
                                                   (LoadResPicture("FRECCIAUP", vbResBitmap).Width * 0.6) / 2, 0)
    CommandImg(0).AutoRedraw = False
    
    'CommandImg(1).Width = UserControl.Width
    CommandImg(1).Align = vbAlignBottom
    CommandImg(1).AutoRedraw = True
    Call CommandImg(1).PaintPicture(LoadResPicture("FRECCIADWN", vbResBitmap), UserControl.Width / 2 - _
                                                   (LoadResPicture("FRECCIADWN", vbResBitmap).Width * 0.6) / 2, 0)
    CommandImg(1).AutoRedraw = False
    
    picSfondoSlide.Top = CommandImg(0).Height
    picSfondoSlide.Height = CommandImg(1).Top - picSfondoSlide.Top
    picSfondoSlide.Width = UserControl.Width
    picSfondoSlide.Left = 0
    picSfondoSlide.Visible = False
    
    picSlide.Height = 135
    picSlide.Width = 135 + 15
    picSlide.Left = (picSfondoSlide.Width / 2) - (picSlide.Width / 2)
    picSlide.Top = picSfondoSlide.Top
    picSlide.BackColor = vbGreen
    
    lnSlide.Y1 = picSfondoSlide.Top
    lnSlide.Y2 = lnSlide.Y1 + picSfondoSlide.Height
    lnSlide.X1 = picSlide.Left + (picSlide.Width / 2) - 10
    lnSlide.X2 = lnSlide.X1
    
    lnSlide2.Y1 = picSfondoSlide.Top
    lnSlide2.Y2 = lnSlide2.Y1 + picSfondoSlide.Height
    lnSlide2.X1 = picSlide.Left + (picSlide.Width / 2)
    lnSlide2.X2 = lnSlide2.X1

End Sub

Private Sub UserControl_Resize()
    Dim IPIC As IPictureDisp
On Error Resume Next
    vsGhost.Left = UserControl.Width + vsGhost.Width + 100
    
    CommandImg(0).AutoRedraw = True
    Call CommandImg(0).Cls
    
    'Set IPIC = LoadPicture(App.Path & "\img_new\aup_.bmp")
    Set IPIC = LoadResPicture("FRECCIAUP", vbResBitmap)
    Call CommandImg(0).PaintPicture(IPIC, UserControl.Width / 2 - ((IPIC.Width * 0.6) / 2) - 10, 0)
    
    CommandImg(0).AutoRedraw = False
    
    CommandImg(1).AutoRedraw = True
    Call CommandImg(1).Cls
    
    'Set IPIC = LoadPicture(App.Path & "\img_new\adown_.bmp")
    Set IPIC = LoadResPicture("FRECCIADWN", vbResBitmap)
    Call CommandImg(1).PaintPicture(IPIC, UserControl.Width / 2 - ((IPIC.Width * 0.6) / 2) - 10, 0)
    
    CommandImg(1).AutoRedraw = False
    
    picSfondoSlide.Top = CommandImg(0).Height
    If CommandImg(1).Top >= picSfondoSlide.Top Then
        picSfondoSlide.Height = CommandImg(1).Top - picSfondoSlide.Top
    End If
    picSfondoSlide.Width = UserControl.Width
    picSfondoSlide.Left = 0
    
    picSlide.Height = 135
    picSlide.Width = 135 + 15
    picSlide.Left = (picSfondoSlide.Width / 2) - (picSlide.Width / 2)
    picSlide.Top = picSfondoSlide.Top
    picSlide.BackColor = vbGreen
    
    lnSlide.Y1 = picSfondoSlide.Top
    lnSlide.Y2 = lnSlide.Y1 + picSfondoSlide.Height
    lnSlide.X1 = picSlide.Left + (picSlide.Width / 2) - 10
    lnSlide.X2 = lnSlide.X1
    lnSlide.BorderColor = RGB(206, 206, 206)
    
    lnSlide2.Y1 = picSfondoSlide.Top
    lnSlide2.Y2 = lnSlide2.Y1 + picSfondoSlide.Height
    lnSlide2.X1 = picSlide.Left + (picSlide.Width / 2)
    lnSlide2.X2 = lnSlide2.X1
    lnSlide2.BorderColor = RGB(206, 206, 206)
    
End Sub

Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
        Dim Value As Long
        If Button = 0 Then
            YSlide1 = Y
        End If
           If Button = 1 Or MouvementSlide Then
                MouvementSlide = True
                Select Case picSlide.Top - (YSlide1 - Y)
                    Case Is > (picSfondoSlide.Top + picSfondoSlide.Height - picSlide.Height)
                        picSlide.Top = (picSfondoSlide.Top + picSfondoSlide.Height - picSlide.Height)
                    Case Is <= picSfondoSlide.Top
                        picSlide.Top = picSfondoSlide.Top
                    Case Else
                        picSlide.Top = picSlide.Top - (YSlide1 - Y)
                End Select

                Value = (((vsGhost.Max - vsGhost.Min) / (picSfondoSlide.Height - picSlide.Height)) * (picSlide.Top - picSfondoSlide.Top)) \ 1
                vsGhost.Value = vsGhost.Min + Value
            End If
            
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "picSlide_MouseMove" & vbCrLf & Err.Description)
    End If
End Sub

Public Property Get xValue() As Long

    xValue = vsGhost.Value
    
End Property

Public Property Let xValue(ByVal Value As Long)
    
    If Value >= vsGhost.Min And Value <= vsGhost.Max Then
        If Value = 0 Then
            Call vsGhost_Change
        End If
        vsGhost.Value = Value
    End If

End Property

Public Property Get xMax() As Long

    xMax = vsGhost.Max
    
End Property

Public Property Let xMax(ByVal Value As Long)

    vsGhost.Max = Value

End Property

Public Property Get xMin() As Long

    xMin = vsGhost.Min
    
End Property

Public Property Let xMin(ByVal Value As Long)

    vsGhost.Min = Value

End Property

Private Sub vsGhost_Change()
Dim VeraAltezza As Integer
Dim VeroMassimo As Integer
Dim VeroValore As Integer

VeroMassimo = vsGhost.Max - vsGhost.Min
If VeroMassimo = 0 Then
    VeroMassimo = 1
End If
VeraAltezza = picSfondoSlide.Height - picSlide.Height
VeroValore = Abs(vsGhost.Min) + vsGhost.Value

    If MouvementSlide = False Then
        picSlide.Top = picSfondoSlide.Top + ((VeraAltezza / VeroMassimo) * VeroValore)
    End If
    RaiseEvent Change
End Sub
