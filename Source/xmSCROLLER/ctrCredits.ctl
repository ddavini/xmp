VERSION 5.00
Begin VB.UserControl ctrCredits 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   2385
      Width           =   4800
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   120
   End
End
Attribute VB_Name = "ctrCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum Direz
    Su = -1
    Giu = 1
End Enum

Private CreditText() As String
Private NumLines As Integer

Private p_Intervallo As Long
Private p_Y As Long
Private Filez As String
Private Mouvement As Boolean
Private Direzione As Boolean
Private X1 As Integer, Y1 As Integer
 
Private MaxLineLen As Long

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCCOPY = &HCC0020


Public Function Start() As Boolean
    Dim Temp As String
    Dim I As Long
    Dim FileNumber As Integer
    Dim MaxLineIndex As Long
On Error GoTo errH
    p_Y = picBuffer.ScaleHeight
    Timer1.Interval = p_Intervallo
    FileNumber = FreeFile
    I = 0
    Open Filez For Input Shared As #FileNumber
        While Not EOF(FileNumber)
            ReDim Preserve CreditText(I)
            Line Input #FileNumber, CreditText(I)
            I = I + 1
        Wend
    Close #FileNumber
    
    NumLines = UBound(CreditText)
    
    For I = 0 To NumLines
        CreditText(I) = Replace$(RTrim$(CreditText(I)), vbTab, Space(8))
        If MaxLineLen < picBuffer.TextWidth(CreditText(I)) Then
            MaxLineLen = picBuffer.TextWidth(CreditText(MaxLineIndex))
            MaxLineIndex = I
        End If
    Next
    While picBuffer.TextWidth(CreditText(MaxLineIndex)) > UserControl.ScaleWidth
            picBuffer.FontSize = picBuffer.FontSize - 1
    Wend
    
    MaxLineLen = picBuffer.TextWidth(CreditText(MaxLineIndex))
    Start = True
    Exit Function
    
errH:
    Err.Clear
    Start = False
End Function


Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mouvement Then
        If Y1 - Y < 0 Then
            Direzione = True
        Else
            Direzione = False
        End If
    End If
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Y1 = Y
    Mouvement = True
End Sub


Private Sub picScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mouvement = False
End Sub


Private Sub UserControl_Initialize()
    picBuffer.Font = "COURIER NEW"
    picBuffer.FontSize = 8
    picScreen.Font = picBuffer.Font
    picScreen.FontSize = picBuffer.FontSize
    picBuffer.Visible = False
    picBuffer.ForeColor = vbGreen
    picScreen.BackColor = RGB(0, 0, 0)
    picBuffer.BackColor = RGB(0, 0, 0)
    
    picBuffer.Top = 0
    picBuffer.Left = 0
    picBuffer.Height = UserControl.ScaleHeight
    picBuffer.Width = UserControl.ScaleWidth
    
    picScreen.Top = 0
    picScreen.Left = 0
    picScreen.Height = UserControl.ScaleHeight
    picScreen.Width = UserControl.ScaleWidth
End Sub

Private Sub Timer1_Timer()
If Not Mouvement Then
    If Not Direzione Then
        Call Moto(Su)
    Else
        Call Moto(Giu)
    End If
End If
End Sub

Private Sub UserControl_Resize()
    picBuffer.Top = 0
    picBuffer.Left = 0
    picBuffer.Height = UserControl.ScaleHeight
    picBuffer.Width = UserControl.ScaleWidth
    
    picScreen.Top = 0
    picScreen.Left = 0
    picScreen.Height = UserControl.ScaleHeight
    picScreen.Width = UserControl.ScaleWidth
    
End Sub

Private Sub UserControl_Terminate()
    Timer1.Interval = 0
End Sub

Private Sub Moto(Dove As Direz)
    Dim AltezzaCTR As Long
    Dim LarghezzaCTR As Long
    Dim TopLBL As Long
    Dim I As Long
    Dim XX As Long
    
    LarghezzaCTR = picBuffer.ScaleWidth
    AltezzaCTR = picBuffer.ScaleHeight
    If Dove = Su Then
        If (p_Y + picBuffer.TextHeight(CreditText(0)) * NumLines) < 0 Then
            p_Y = picBuffer.ScaleHeight
        End If
    Else
        If p_Y > picBuffer.ScaleHeight Then
            p_Y = -(picBuffer.TextHeight(CreditText(0)) * NumLines)
        End If
    End If
        
        picBuffer.CurrentY = p_Y + Dove
        p_Y = picBuffer.CurrentY
        XX = picBuffer.ScaleWidth
        
        For I = 0 To NumLines
            DoEvents
            picBuffer.CurrentX = (XX * 0.5) - _
            (MaxLineLen * 0.5)
            picBuffer.Print CreditText(I)
        Next

        Call BitBlt(picScreen.hDC, 0, picScreen.ScaleTop, picScreen.ScaleWidth, picScreen.ScaleHeight, picBuffer.hDC, 0, 0, SRCCOPY)
        picBuffer.Cls
        

End Sub

Public Property Get Intervallo() As Long

    Intervallo = p_Intervallo
    
End Property

Public Property Let Intervallo(ByVal Value As Long)

    p_Intervallo = Value

End Property

Public Property Get CreditFile() As String

    CreditFile = Filez
    
End Property

Public Property Let CreditFile(ByVal Value As String)

    Filez = Value

End Property

