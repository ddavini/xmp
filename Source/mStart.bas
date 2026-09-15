Attribute VB_Name = "mStart"
Option Explicit

'Da spostare nel Mod API'
Private Declare Function CoLockObjectExternal Lib "ole32" ( _
    ByVal pUnk As IUnknown, ByVal fLock As Long, _
    ByVal fLastUnlockReleases As Long) As Long


Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) _
    As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long

Private m_colRunnables As Collection

Private m_lTimerID As Long

Private Sub TimerProc(ByVal lHwnd As Long, ByVal lMsg As Long, _
    ByVal lTimerID As Long, ByVal lTime As Long)

    Dim this As xmService
   
    With m_colRunnables
       Do While .count > 0
          Set this = .Item(1)
          .Remove 1
          this.Elabora
          CoLockObjectExternal this, 0, 1
       Loop
    End With
    KillTimer 0, lTimerID
    m_lTimerID = 0
    
End Sub

Public Sub Start(this As xmService)
   CoLockObjectExternal this, 1, 1
   If m_colRunnables Is Nothing Then
      Set m_colRunnables = New Collection
   End If
   m_colRunnables.Add this
   If Not m_lTimerID Then
      m_lTimerID = SetTimer(0, 0, 1, AddressOf TimerProc)
   End If
End Sub

