VERSION 5.00
Begin VB.Form frmListone 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Listone.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer frmListoneTimer 
      Left            =   240
      Top             =   2640
   End
   Begin XmP1.xmDisplay xmDInfo 
      Height          =   135
      Index           =   0
      Left            =   2185
      TabIndex        =   9
      Top             =   1560
      Width           =   2205
      _extentx        =   3889
      _extenty        =   450
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
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   180
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   1440
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Seek Active Stream"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox Ripeti 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   5
      ToolTipText     =   "Loop Active"
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   720
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Manage List"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   960
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Mp3 Up"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1200
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mp3 Down"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Delete Item"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox CommandImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Clear"
      Top             =   1560
      Width           =   240
   End
   Begin XmP1.xmListBox ListaMp3 
      Height          =   390
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   2295
      _extentx        =   4048
      _extenty        =   688
      selectmodeappearance=   1
      enabled         =   -1  'True
      fontcolor       =   65280
      backcolor       =   0
      borderstyle     =   0
      font            =   "Listone.frx":000C
      font            =   "Listone.frx":0034
   End
   Begin XmP1.xmDisplay xmDInfo 
      Height          =   135
      Index           =   1
      Left            =   2185
      TabIndex        =   10
      Top             =   1695
      Width           =   2205
      _extentx        =   3889
      _extenty        =   450
   End
   Begin VB.Line lnSposta2 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line lnSposta 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   4320
      Y1              =   195
      Y2              =   195
   End
End
Attribute VB_Name = "frmListone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X1 As Integer, Y1 As Integer
Private Mouvement As Boolean
Private MouseUp As Boolean
Private WithEvents DragMP3 As DragEngine
Attribute DragMP3.VB_VarHelpID = -1

Public ScancellatoDaXmp As Boolean
Public MouseD As Boolean

Const MovimentoUpDwn = 100

Private Sub Clear_Click()
On Error GoTo ErrH
    Call StopAll
    ListaMp3.Clear
    Call KillTmpDat(False)
    ReDim file(0)
    xmp.xmDisplay(0).xCaption = "Nope"
    xmp.Durata = "00:00"
    xmp.BitRate = "0"
    Call InfoNmp3
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Clear_Click" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub CommandImg_Click(Index As Integer)
On Error GoTo ErrH
    Dim ListFile As String
    Select Case Index
        Case Is = 0 'Clear
             Call PulisciTutto(True)
        Case Is = 1 'Delete
            Call ListaMp3_KeyPress(Asc("d"))
        Case Is = 2 'UP List
            Call SpostaItem(ListaMp3.ListIndex, Su)
            Call SpostaItemList(ListaMp3.ListIndex, ListaMp3, Su)
        Case Is = 3 'DWN List
            Call SpostaItem(ListaMp3.ListIndex, Giu)
            Call SpostaItemList(ListaMp3.ListIndex, ListaMp3, Giu)
        Case Is = 4 'Quick Save
            ListFile = GetINI(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
            frmMenu.mnuLstFileName.Caption = RetFileName(ListFile)
            frmMenu.mnuLstFileName.Enabled = False
            
            PopupMenu frmMenu.mnuPlayList

        Case Is = 5 'Seek
            ListaMp3.ListIndex = -1
            ListaMp3.ListIndex = IndiceGlobalissimo
    End Select
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.CommandImg_Click" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub CommandImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    Call IlluminaPulsante(CommandImg(Index), True)
                    
    MouseUp = False
    Select Case Index
        Case Is = 6 'Lista Giu'
            While Not MouseUp
                If ListaMp3.ScrollValue < ListaMp3.ScrollMax Then
                    ListaMp3.ScrollValue = ListaMp3.ScrollValue + 1
                End If
                Call Pause(0.3, True)
            Wend
        Case Is = 7 'Lista Su
            While Not MouseUp
                If ListaMp3.ScrollValue > 0 Then
                    ListaMp3.ScrollValue = ListaMp3.ScrollValue - 1
                End If
                Call Pause(0.3, True)
            Wend
    End Select

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.CommandImg_MouseDown" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub CommandImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseUp = True

    Call IlluminaPulsante(CommandImg(Index), False)
End Sub

Private Sub Form_Activate()
    Me.lnSposta2.BorderColor = gcActiveColor
    Me.lnSposta.BorderColor = gcActiveColor
End Sub

Private Sub Form_Deactivate()
    Me.lnSposta2.BorderColor = gcDeactiveColor
    Me.lnSposta.BorderColor = gcDeactiveColor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrH
    
    If KeyAscii = 27 Then
        Call DeLoad(True)
    Else
        Call SearchList(KeyAscii, ListaMp3)
        Call frmHotKey(KeyAscii)
    End If
 
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Form_KeyPress" & vbCrLf & Err.Description)
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrH
    
    'Me.KeyPreview = False

If Command$ <> "/debug" Then
    Set DragMP3 = New DragEngine
    DragMP3.DragHwnd = Me.hWnd
    DragMP3.StartDrag
    
    gb_OldProcMouseWheel = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf MouseWheel
End If
       
    'Carica Res'
'    CommandImg(3).Picture = LoadPicture(App.Path & "\img_new\aup_.bmp") 'LoadResPicture("FRECCIAUP", vbResBitmap)
'    CommandImg(2).Picture = LoadPicture(App.Path & "\img_new\adown_.bmp") 'LoadResPicture("FRECCIADWN", vbResBitmap)
'    CommandImg(0).Picture = LoadPicture(App.Path & "\img_new\clear.bmp") 'LoadResPicture("PICCLEAR", vbResBitmap)
'    CommandImg(1).Picture = LoadPicture(App.Path & "\img_new\delete.bmp") 'LoadResPicture("PICDELETE", vbResBitmap)
'    CommandImg(4).Picture = LoadPicture(App.Path & "\img_new\salvaquick.bmp") 'LoadResPicture("PICSAVE", vbResBitmap)
'    CommandImg(5).Picture = LoadPicture(App.Path & "\img_new\seek.bmp") 'LoadResPicture("PICSEEK", vbResBitmap)
'    Ripeti.Picture = LoadPicture(App.Path & "\img_new\loopattivo.bmp") ' LoadResPicture("PICLOOPACTIVE", vbResBitmap)
    
    CommandImg(3).Picture = LoadResPicture("FRECCIAUP", vbResBitmap)
    CommandImg(2).Picture = LoadResPicture("FRECCIADWN", vbResBitmap)
    CommandImg(0).Picture = LoadResPicture("PICCLEAR", vbResBitmap)
    CommandImg(1).Picture = LoadResPicture("PICDELETE", vbResBitmap)
    CommandImg(4).Picture = LoadResPicture("PICSAVE", vbResBitmap)
    CommandImg(5).Picture = LoadResPicture("PICSEEK", vbResBitmap)
    Ripeti.Picture = LoadResPicture("PICLOOPACTIVE", vbResBitmap)
    MenuBarPic.Picture = LoadResPicture("EXIT", vbResBitmap)
    
    'LoadPicture(App.Path & "\img_new\bkg_.bmp")
    Me.Picture = LoadDataIntoFile("BKGCPRS")
    '''
    
    Me.Width = xmp.Width
    Me.Top = xmp.Top + xmp.Height
    Me.Left = xmp.Left
    Me.Height = xmp.Height
    
    ListaMp3.ScrollBarWidth = CommandImg(5).Width
    ListaMp3.Left = 150
    ListaMp3.Top = 360
    ListaMp3.Height = Abs(Me.Height - 855)
    ListaMp3.Width = Abs(Me.Width - 250)
           
    Call RoundForm(Me)
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Form_Load" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    If Button = 2 Then
      ListaMp3.ListIndex = -1
      ListaMp3.ListIndex = IndiceGlobalissimo
    End If
    MouseD = True
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Form_MouseDown" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
        If Button = 0 Then
            Y1 = Y
            X1 = X
        End If
        If Y < Me.Height - 500 Then
            If Y < 1000 And X >= lnSposta.X1 And X <= lnSposta.X2 Or Mouvement = True Then
                If Button = 1 Then
                    Mouvement = True
                    Me.Left = Me.Left - (X1 - X)
                    Me.Top = Me.Top - (Y1 - Y)
                End If
            End If
        End If
        
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Form_MouseMove" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseD = False
    Mouvement = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrH
If Command$ <> "/debug" Then
    DragMP3.StopDrag
    Pause
    Set DragMP3 = Nothing
    Pause
End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.Form_Unload" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub frmListoneTimer_Timer()
On Error GoTo ErrH
    'Aggancia Lista Mp3
    If xmp.WindowState = 0 And frmListone.ScancellatoDaXmp = False And _
    frmListone.Visible = False Then
        frmListone.Show
        If frmShowTask.ShowMe Then
            Call ShowTaskBarIcon(True)
        End If
    End If
    If frmListone.MouseD = False Then
        Call AgganciaSub
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.frmListoneTimer_Timer" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub ListaMp3_Click()
On Error GoTo ErrH
    If mStreamIsActive = False And ListaMp3.ListIndex <> -1 Then
        IndiceGlobalissimo = ListaMp3.ListIndex
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.ListaMp3_Click" & vbCrLf & Err.Description)
    End If

End Sub

Private Sub ListaMp3_DblClick()
    Call PlayStream(ListaMp3.ListIndex)
End Sub

Private Sub ListaMp3_GotFocus()
On Error GoTo ErrH
    If ListaMp3.ListCount = 0 Then
        ListaMp3.ListIndex = -1
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.ListaMp3_GotFocus" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub ListaMp3_KeyPress(KeyAscii As Integer)
On Error GoTo ErrH
    Dim I As Integer
    Dim Suiccessivo
        If KeyAscii = Asc("d") Then
            If ListaMp3.SelectedCount > 0 Then
                Select Case ListaMp3.ListCount
                    Case Is > 1
                        If ListaMp3.IsSelected(IndiceGlobalissimo) Then
                            StopAll
                        End If
                        If ListaMp3.ListIndex + 1 = ListaMp3.ListCount Then
                            'ReDim Preserve Directory(UBound(Directory) - 1)
                            Call DelFile(NFile)
                            I = ListaMp3.ListIndex - 1
                            ListaMp3.RemoveItem (ListaMp3.ListIndex)
                            ListaMp3.ListIndex = I
                        Else
                            Call xmpPensa(True)
                            Call SetFile(ListaMp3.ListIndex, GetFile(ListaMp3.ListIndex + 1))
                            For I = ListaMp3.ListIndex + 1 To NFile - 1
                                Call SetFile(I, GetFile(I + 1))
                            Next
                            'ReDim Preserve Directory(UBound(Directory) - 1)
                            Call DelFile(NFile)
                            I = ListaMp3.ListIndex
                            ListaMp3.RemoveItem (ListaMp3.ListIndex)
                            ListaMp3.ListIndex = I
                            Call xmpPensa(False)
                        End If
                        If IndiceGlobalissimo = NFile + 1 Then
                            IndiceGlobalissimo = IndiceGlobalissimo - 1
                        End If
                    Case Is = 1
                        Call KillTmpDat(False)
                        ListaMp3.Clear
                        IndiceGlobalissimo = -1
                        Call PulisciTutto
                End Select
                Call InfoNmp3
            Else
                PulisciTutto
            End If
        End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.ListaMp3_KeyPress" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub ListaMp3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrH
    If ListaMp3.ListCount = 0 Then
        ListaMp3.ListIndex = -1
    End If
    If Button = 2 Then
        PopupMenu frmMenu.mnuIDtag
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Listone.ListaMp3_MouseDown" & vbCrLf & Err.Description)
    End If
End Sub

Private Sub DragMP3_FilesDroped()
    Call GestioneExtDrag(DragMP3)
End Sub

Private Sub MenuBarPic_Click()
    frmListone.Visible = Not frmListone.Visible
    frmListone.ScancellatoDaXmp = Not frmListone.Visible
End Sub

Private Sub MenuBarPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, True)
End Sub

Private Sub MenuBarPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call IlluminaPulsante(MenuBarPic, True)
End Sub


Private Sub xmDInfo_Click(Index As Integer)
    Call InfoNmp3
End Sub

