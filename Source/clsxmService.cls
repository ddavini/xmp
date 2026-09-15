VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ServiceName
    Swirl = 0
    Scrol = 1
End Enum

'Public Event lblSwirlEvent(Colore As Long)
Public Event ScorriLabelNomeMp3Event(lblNomeMP3 As String)
Public Event EndElab()


Public pServiceName As ServiceName
Public inDebug As Boolean
'Public lblColor As Long
Public ScorrilblModo As Byte
Public ScorrilblMaxChar As Integer
Public ScorrilblLungPausa As Double
Public ScorrilblCaption As String

Public Function SeviceStart(NomeService As ServiceName)
    Call Start(Me)
    pServiceName = NomeService
End Function

Friend Sub Elabora()
    Select Case pServiceName
        Case Is = Scrol
            Call ScorriLabelNomeMp3(ScorrilblMaxChar, ScorrilblCaption, ScorrilblModo, ScorrilblLungPausa)
    End Select
    RaiseEvent EndElab
End Sub

Public Sub ScorriLabelNomeMp3(MaxChar As Integer, _
                       NomeMp3 As String, _
                       Optional Modo As Byte = 0, _
                       Optional LungPausa As Double = 0.08)
On Error GoTo ErrH
    Dim Salva As String
    Dim Elabo As String
    Dim I As Integer
        Salva = NomeMp3
        If Len(Salva) > MaxChar Then
            Select Case Modo
                Case Is = 0 'Indietro Avanti / Backward Forward
                    MaxChar = MaxChar \ 2
                    For I = 0 To Len(Salva) - MaxChar
                        Elabo = Right$(Salva, Len(Salva) - I)
                        RaiseEvent ScorriLabelNomeMp3Event(Elabo)
                        Pause LungPausa, True
                    Next
                    For I = Len(Salva) - MaxChar To 0 Step -1
                        RaiseEvent ScorriLabelNomeMp3Event(Right$(Salva, Len(Salva) - I))
                        Pause LungPausa, True
                    Next
                Case Is = 1 'Scrittura una lettera alla volta / Writing one letter at a time
                    For I = Len(Salva) To 0 Step -1
                        RaiseEvent ScorriLabelNomeMp3Event(Left$(Salva, Len(Salva) - I))
                        Pause LungPausa, True
                    Next
                Case Is = 2 'Scrittura una lettera alla volta avanti indre' / Writing one letter at a time, forward and backward
                    For I = 0 To Len(Salva)
                        RaiseEvent ScorriLabelNomeMp3Event(Left$(Salva, Len(Salva) - I))
                        Pause LungPausa, True
                    Next
                    For I = Len(Salva) To 0 Step -1
                        RaiseEvent ScorriLabelNomeMp3Event(Left$(Salva, Len(Salva) - I))
                        Pause LungPausa, True
                    Next
                End Select
        End If
        Pause LungPausa, True
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmService.ScorriLabel" & vbCrLf & Err.Description)
    End If
End Sub

Friend Sub RaiseEndElab()
    RaiseEvent EndElab
End Sub
