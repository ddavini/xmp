Attribute VB_Name = "mdlGVars"
Option Explicit

Private Const conHwndTopmost = -1
Private Const conHwndNoTopmost = -2
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40

Public Const ININAME As String = "LCD232.INI"

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type tApplication
    SessionID As Long
    SessionTimeout As Long
    Name As String
    Key As String
End Type

Public g_Applications() As tApplication



Public Sub FillLstApps(l As ListBox, g() As tApplication)
    Dim I As Long
    
    l.Clear
    
    If g(0).SessionID = -1 Then Exit Sub
    For I = LBound(g) To UBound(g)
        l.AddItem "(" & Format$(g(I).SessionID, "@@@@@@") & ") " & g(I).Name, 0
        l.ItemData(l.NewIndex) = g(I).SessionID
    Next
End Sub

Public Sub AlwaysOnTop(hWnd As Long, _
                              Width As Long, _
                              Height As Long, _
                              X As Long, _
                              Y As Long, _
                              Optional AOnTop As Boolean = True)
On Error GoTo ErrH
    X = X \ Screen.TwipsPerPixelX
    Y = Y \ Screen.TwipsPerPixelY
    Width = Width \ Screen.TwipsPerPixelY
    Height = Height \ Screen.TwipsPerPixelX
    If AOnTop Then
        SetWindowPos hWnd, conHwndTopmost, X, Y, Width, Height, _
                    conSwpNoActivate
                    'conSwpNoActivate Or conSwpShowWindow
    Else
        SetWindowPos hWnd, conHwndNoTopmost, X, Y, Width, Height, _
                    conSwpNoActivate
                    'conSwpNoActivate Or conSwpShowWindow
    End If
ErrH:
    If Err.Number <> 0 Then
        MsgBox (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "ChiamateAPI.SempreInPrimoPiano" & vbCrLf & Err.Description)
    End If
End Sub

