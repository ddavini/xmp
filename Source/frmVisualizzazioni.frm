VERSION 5.00
Begin VB.Form frmVisualizzazioni 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmVisualizzazioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmVisualizzazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents xmServiceObjScroll As xmService
Attribute xmServiceObjScroll.VB_VarHelpID = -1
Private xmServiceObjScrollActive As Boolean
Public Shock As Boolean

Private Sub Form_Load()
     Shock = False
     Timer.Interval = 35
End Sub
Private Sub Start()
        
        Do

            Call DisplayAll(GlobalSpectrumeMode)
            DoEvents
            Call SleepEx(1, True)
'            Call ListEvent

        Loop Until frmVisualizzazioniAttivo = False
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmVisualizzazioniAttivo = False
End Sub

Private Sub Timer_Timer()
    
    frmVisualizzazioniAttivo = True
    Timer.Interval = 0
    Start
    
End Sub

Private Sub xmServiceObjScroll_EndElab()
    Set xmServiceObjScroll = Nothing
    xmServiceObjScrollActive = False
End Sub
            
Public Sub ScorriLabelNomeMp3(MaxChar As Integer, _
                       Optional Modo As Byte = 0, _
                       Optional LungPausa As Double = 0.08)
On Error Resume Next
    
    If xmServiceObjScrollActive = False Then
        xmServiceObjScrollActive = True
        Set xmServiceObjScroll = New xmService
        
        xmServiceObjScroll.ScorrilblCaption = xmp.xmDisplay(0).xCaption
        xmServiceObjScroll.ScorrilblLungPausa = LungPausa
        xmServiceObjScroll.ScorrilblMaxChar = MaxChar
        xmServiceObjScroll.ScorrilblModo = Modo
        Call xmServiceObjScroll.SeviceStart(Scrol)
    End If
            
If Err.Number <> 0 Then
    ScriviLOG Err.Source & " - frmVisualizzazioni.ScorriLabelNomeMp3", Err.Number, Err.Description
    If Shock Then
        SettaNomeMp3
    End If
End If

Err.Clear
            
End Sub

Private Sub xmServiceObjScroll_ScorriLabelNomeMp3Event(lblNomeMP3 As String)
    
    xmp.xmDisplay(0).xCaption = lblNomeMP3
    
End Sub
