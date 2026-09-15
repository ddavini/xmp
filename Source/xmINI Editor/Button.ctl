VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ScaleHeight     =   2925
   ScaleWidth      =   4650
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox ButtonPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   0
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   0
      X2              =   2400
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Event Declarations:
Event Click() 'MappingInfo=pic,pic,-1,Click
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseUp


Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
pic.Cls
Draw3d pic, True
DrawPicture True
pic.Refresh
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
    With pic  'Change this to the name of the control

       If Button = 0 Then

          If (x < 0) Or (Y < 0) Or (x > .Width) Or (Y > .Height) Then
              'Mouse pointer is outside button, so let other controls receive
              'mouseevents too:
              ReleaseCapture

              ' Do your 'mouse-exit' stuff here
              .Cls
              DrawPicture False
              pic.Refresh
              .Refresh
          Else
              ' Mouse pointer is over button, so we'll capture it, thus
              ' we'll receive mouse messages even if the mouse pointer is
              ' not over the button
              SetCapture .hWnd

              ' Do your 'mouse-enter' stuff here
              pic.Cls
              Draw3d pic, False
              DrawPicture False
              pic.Refresh
          End If
       End If

    End With
End Sub
Private Sub DrawPictureWithoutMask(Source As PictureBox, Dest As PictureBox, MaskColor As OLE_COLOR)
Dim Height As Integer
Dim Width As Integer
Dim I As Integer
Dim A As Integer
Dim FoundColor As Long

Height = (Source.Height / Screen.TwipsPerPixelY) - 1
Width = (Source.Width / Screen.TwipsPerPixelX) - 1

Dest.Cls
For I = 0 To Height
    For A = 0 To Width
        FoundColor = GetPixel(Source.hdc, A, I)
        If Not FoundColor = MaskColor Then
            SetPixel Dest.hdc, A, I, FoundColor
        End If
    Next A
Next I
Dest.Refresh
End Sub
Private Sub pic_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
pic.Cls
Draw3d pic, False
DrawPicture False
pic.Refresh
End Sub

Private Function Draw3d(Dest As Control, Pushed As Boolean)
Dim Height As Integer
Dim Width As Integer

Width = Dest.Width - Screen.TwipsPerPixelX ') ' - 20
Height = Dest.Height - Screen.TwipsPerPixelY ') ' - 20

'Dest.ForeColor = vbWhite

If Pushed = False Then
    Dest.Line (0, 0)-(0, Height), Line1.BorderColor
    Dest.Line (0, 0)-(Width, 0), Line1.BorderColor
    
    Dest.Line (Width, Height)-(0, Height), Line2.BorderColor
    Dest.Line (Width, Height)-(Width, 0), Line2.BorderColor
Else
    Dest.Line (0, 0)-(0, Height), Line2.BorderColor
    Dest.Line (0, 0)-(Width, 0), Line2.BorderColor
    
    Dest.Line (Width, Height)-(0, Height), Line1.BorderColor
    Dest.Line (Width, Height)-(Width, 0), Line1.BorderColor
End If


End Function

Private Sub UserControl_Initialize()
pic.Cls
DrawPicture False
pic.Refresh
End Sub

Private Sub UserControl_Resize()
pic.Width = UserControl.ScaleWidth
pic.Height = UserControl.ScaleHeight

pic.Cls
DrawPicture False
pic.Refresh
End Sub

Private Sub DrawPicture(Down As Boolean)
'On Error Resume Next
Dim DestX As Integer
Dim DestY As Integer

If Down = True Then
    DestX = ((pic.Width / 2) - (ButtonPic.Width / 2)) + Screen.TwipsPerPixelX
    DestY = ((pic.Height / 2) - (ButtonPic.Height / 2)) + Screen.TwipsPerPixelY
Else
    DestX = (pic.Width / 2) - (ButtonPic.Width / 2)
    DestY = (pic.Height / 2) - (ButtonPic.Height / 2)
End If
pic.PaintPicture tmpPic.Image, DestX, DestY, ButtonPic.Width, ButtonPic.Height, 0, 0, ButtonPic.Width, ButtonPic.Height

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = pic.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    pic.BackColor() = New_BackColor
    tmpPic.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = ButtonPic.BackColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    Dim New_Mask As OLE_COLOR
    ButtonPic.BackColor() = New_Mask
    DrawPictureWithoutMask ButtonPic, tmpPic, ButtonPic.BackColor
    PropertyChanged "MaskColor"
End Property

Private Sub pic_Click()
    RaiseEvent Click
    pic.Cls
    DrawPicture False
    pic.Refresh
    pic.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ButtonPic,ButtonPic,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = ButtonPic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set ButtonPic.Picture = New_Picture
    PropertyChanged "Picture"
    
    tmpPic.Width = ButtonPic.Width
    tmpPic.Height = ButtonPic.Height
    
    DrawPictureWithoutMask ButtonPic, tmpPic, ButtonPic.BackColor
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Line1,Line1,-1,BorderColor
'Public Property Get LineColor1() As Long
'    LineColor1 = Line1.BorderColor
'End Property
'
'Public Property Let LineColor1(ByVal New_LineColor1 As Long)
'    Line1.BorderColor() = New_LineColor1
'    PropertyChanged "LineColor1"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Line2,Line2,-1,BorderColor
'Public Property Get LineColor2() As Long
'    LineColor2 = Line2.BorderColor
'End Property
'
'Public Property Let LineColor2(ByVal New_LineColor2 As Long)
'    Line2.BorderColor() = New_LineColor2
'    PropertyChanged "LineColor2"
'End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    pic.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Line1.BorderColor = PropBag.ReadProperty("LineColor1", 16777215)
    Line2.BorderColor = PropBag.ReadProperty("LineColor2", 8421504)
    Line1.BorderColor = PropBag.ReadProperty("LineColor1", 16777215)
    Line2.BorderColor = PropBag.ReadProperty("LineColor2", 16777215)
    ButtonPic.BackColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
End Sub

Private Sub UserControl_Show()
pic.Cls
DrawPicture False
pic.Refresh
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", pic.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("LineColor1", Line1.BorderColor, 16777215)
    Call PropBag.WriteProperty("LineColor2", Line2.BorderColor, 8421504)
    Call PropBag.WriteProperty("LineColor1", Line1.BorderColor, 16777215)
    Call PropBag.WriteProperty("LineColor2", Line2.BorderColor, 16777215)
    Call PropBag.WriteProperty("MaskColor", ButtonPic.BackColor, &HFF00FF)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get LineColor1() As OLE_COLOR
    LineColor1 = Line1.BorderColor
End Property

Public Property Let LineColor1(ByVal New_LineColor1 As OLE_COLOR)
    Line1.BorderColor() = New_LineColor1
    PropertyChanged "LineColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderColor
Public Property Get LineColor2() As OLE_COLOR
    LineColor2 = Line2.BorderColor
End Property

Public Property Let LineColor2(ByVal New_LineColor2 As OLE_COLOR)
    Line2.BorderColor() = New_LineColor2
    PropertyChanged "LineColor2"
End Property


