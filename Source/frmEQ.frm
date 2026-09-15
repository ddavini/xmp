VERSION 5.00
Begin VB.Form frmEQ 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2790
   ClientLeft      =   9420
   ClientTop       =   6330
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEQ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin XmP1.xmDisplay xmDFreq 
      Height          =   135
      Index           =   0
      Left            =   840
      TabIndex        =   16
      Top             =   1680
      Width           =   240
      _extentx        =   4233
      _extenty        =   238
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   9
      Left            =   3000
      TabIndex        =   15
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin VB.PictureBox MenuBarPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4680
      ScaleHeight     =   120
      ScaleWidth      =   150
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   180
   End
   Begin VB.OptionButton optEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Treble Boost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1300
   End
   Begin VB.OptionButton optEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1300
   End
   Begin VB.OptionButton optEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Bass Boost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   3
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1300
   End
   Begin VB.OptionButton optEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Pop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1300
   End
   Begin VB.OptionButton optEQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Rock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1300
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   6
      Left            =   2280
      TabIndex        =   12
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   7
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmSlide vsGraphic 
      Height          =   1335
      Index           =   8
      Left            =   2760
      TabIndex        =   14
      Top             =   360
      Width           =   240
      _extentx        =   423
      _extenty        =   2355
   End
   Begin XmP1.xmDisplay xmDFreq 
      Height          =   135
      Index           =   10
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   400
      _extentx        =   714
      _extenty        =   238
   End
   Begin XmP1.xmDisplay xmDFreq 
      Height          =   135
      Index           =   11
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   400
      _extentx        =   714
      _extenty        =   238
   End
   Begin XmP1.xmDisplay xmDFreq 
      Height          =   135
      Index           =   12
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   400
      _extentx        =   714
      _extenty        =   238
   End
   Begin VB.Line lnSposta 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line lnSposta2 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private X1 As Integer, Y1 As Integer
Private Mouvement As Boolean
Private pmeVal(9) As Long

Private Sub Form_Activate()
    Me.lnSposta2.BorderColor = gcActiveColor
    Me.lnSposta.BorderColor = gcActiveColor
End Sub

Private Sub Form_Deactivate()
    Me.lnSposta2.BorderColor = gcDeactiveColor
    Me.lnSposta.BorderColor = gcDeactiveColor
End Sub

Private Sub Form_Load()
On Error GoTo errH:
    Dim I As Integer
    Dim aFreq() As Variant
    
    aFreq = Array("60", "170", "310", "600", "1K", "3K", _
                  "6K", "12K", "14K", "16K")
            
    'Carica Res
    MenuBarPic.Picture = LoadResPicture("EXIT", vbResBitmap)
    
    'LoadPicture(App.Path & "\img_new\bkg_.bmp")
    Me.Picture = LoadDataIntoFile("BKGCPRS")
    '''''
    
    Me.Top = xmp.Top
    Me.Left = xmp.Left
    
    Me.Height = xmp.Height
    Me.Width = xmp.Width
    
    xmDFreq(0).xCaption = aFreq(0)
    For I = 1 To 9
        Load xmDFreq(I)
        xmDFreq(I).Visible = True
        xmDFreq(I).Center = True
        xmDFreq(I).Left = xmDFreq(I - 1).Left + xmDFreq(I - 1).Width
        xmDFreq(I).xCaption = aFreq(I)
    Next I
    
    aFreq = Array("+12db", "0db", "-12db")
    
    For I = 10 To 12
        xmDFreq(I).Center = True
        xmDFreq(I).xCaption = aFreq(I - 10)
    Next I
    
    For I = 0 To 9
        vsGraphic(I).xMax = 127
        vsGraphic(I).xMin = -127
        pmeVal(I) = GetIni(cfgFile, "EQ", CStr(I), "0")
        vsGraphic(I).xValue = pmeVal(I)
        'DoEvents
    Next I
    I = 0
    optEQ(I).Top = lnSposta2.Y1 + 150
    optEQ(I).Left = optEQ(I).Left - 150
    optEQ(I).Height = optEQ(I).Height - 30
    For I = 1 To optEQ.count - 1
        optEQ(I).Top = optEQ(I - 1).Top + optEQ(I - 1).Height + 15
        optEQ(I).Left = optEQ(I - 1).Left
        optEQ(I).Height = optEQ(I - 1).Height
        optEQ(I).Value = False
        optEQ(I).Font = frmListone.ListaMp3.Font
        optEQ(I).FontSize = 8
    Next I
    optEQ(0).Value = False
    
    Call RoundForm(Me)
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmEQ.Form_Load" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH:
    If Button = 0 Then
        Y1 = Y
        X1 = X
    End If
    If Y < 500 And X >= lnSposta.X1 And X <= lnSposta.X2 Or Mouvement Then
        If Button = 1 Then
            Mouvement = True
            Me.Left = Me.Left - (X1 - X)
            Me.Top = Me.Top - (Y1 - Y)
        End If
    End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmEQ.Form_Load" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveEQ
End Sub

Private Sub MenuBarPic_Click()
    Unload Me
End Sub

Private Sub MenuBarPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, True)
End Sub

Private Sub MenuBarPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, True)
End Sub

Private Sub optEQ_Click(Index As Integer)
On Error GoTo errH:
    Dim tmp As Variant
    Dim I As Integer
    
    Select Case Index
    Case 0
        tmp = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    Case 1
        tmp = Array(60, 40, 20, 0, -20, -20, 0, 20, 40, 60)
    Case 2
        tmp = Array(20, 30, 40, 60, 60, 40, 30, 20, 0, 0)
    Case 3
        tmp = Array(60, 80, 40, 20, 0, 0, 0, 0, 0, 0)
    Case 4
        tmp = Array(0, 0, 0, 0, 20, 30, 40, 60, 60, 60)
    End Select
    
    For I = 0 To 9
        vsGraphic(I).xValue = tmp(I)
    Next I
    optEQ(Index).Value = False
    Call mSetEQ(pmeVal(0))
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmEQ.optEQ_Click" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub vsGraphic_Change(Index2 As Integer)
    pmeVal(Index2) = vsGraphic(Index2).xValue
End Sub

Private Sub vsGraphic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call xmMP3_setEqualizer(pmeVal(0))
        Call xmMP3_reload
End Sub

Private Sub SaveEQ()
On Error GoTo errH:
    Dim I As Byte
    For I = 0 To 9
        Call SetIni(cfgFile, "EQ", CStr(I), CStr(pmeVal(I)))
    Next
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmEQ.SaveEQ" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub LoadEQ()
On Error GoTo errH:
    Dim I As Byte
    For I = 0 To 9
        pmeVal(I) = GetIni(cfgFile, "EQ", CStr(I), "0")
    Next I
    optEQ(Index).Value = False
    Call mSetEQ(pmeVal(0))
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "frmEQ.LoadEQ" & vbCrLf & Err.Description)
    End If
End Sub

