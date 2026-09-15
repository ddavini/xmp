VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Sub FFT Lib "xmFFT.dll" (ByVal Sample As Long, _
    ByRef DataReal As Single, ByRef DataImg As Single, ByVal Barre As Long)

Private Declare Sub InitxmFFT Lib "xmFFT.dll" Alias "Init" ()

Const FFTSAMPLE = 1024
Const HH = 100
Const WW = 100

Private Sub Command1_Click()
Do
    Dim BufferD() As Single
    Dim BufferE() As Single
    Dim WaveL() As Long
    Dim WaveR() As Long
    Dim S As Single
    Dim I As Integer
    Dim Step As Integer

    ReDim BufferD(FFTSAMPLE)
    ReDim BufferE(FFTSAMPLE)
    

    
    'Ricordarsi che se pWave e' meno di 256 va in guru
    ReDim WaveR(FFTSAMPLE)
    ReDim WaveL(FFTSAMPLE)
    Call xmMP3_getWave(WaveL(0), WaveR(0))
    For I = 0 To FFTSAMPLE
            BufferD(I) = WaveL(I) + WaveR(I)
    Next
    Call FFT(FFTSAMPLE, BufferD(0), BufferE(0), FFTSAMPLE)
    
    For I = 0 To FFTSAMPLE
            BufferD(I) = BufferD(I)
    Next
    Call Sleep(1)
    DoEvents
    Me.Cls
    
    Step = FFTSAMPLE / 256
'
'    For I = 0 To WW
'        Me.Line -(I, HH - (BufferD(I) * HH)), vbBlack
'    Next I
    For I = 1 To WW Step Step
        Me.Line -(I - 1, HH - (BufferD(I) * HH)), vbBlack
    Next I
Loop Until False
End Sub

Private Sub Form_Load()

    Dim Opt As DEC_OPTION
    Dim Data As InputInfo

    Opt.convert = 0
    Opt.reduction = 0
    Opt.freqLimit = 44000
    ChDir App.Path

    Call InitxmFFT
    Call xmMP3_init
    Call xmMP3_setDecodeOption(Opt)
    If xmMP3_open("1.mp3", Data) = False Then
        MsgBox "Can't create stream", vbCritical, "Open Stream", "mOpenStream"
    End If

    Call xmMP3_play

    Me.ScaleMode = vbPixels
    Me.AutoRedraw = True
    Me.Show
    If Me.ScaleWidth < WW Then
        Me.Width = WW * Screen.TwipsPerPixelX
    End If

    Timer1.Interval = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call xmMP3_stop
    End
End Sub

Private Sub Timer1_Timer()
    Call Command1_Click
End Sub
