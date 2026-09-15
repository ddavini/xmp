VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type tResponse
    sResponse As String
    lResponse As Long
    bResponse As Boolean
End Type
Private p_Response As tResponse


Public Property Get Response() As tResponse
    Response = p_Response
End Property


Public Sub Login(ByVal SessionID As Long, ByVal AppName As String, ByVal SessionTimeout As Long, ByRef Key As String)
    Dim I As Long, J As Long
    Dim da_aggiungere As Boolean
    Dim sLine(1) As String
    
    frmMain.txtData2.Text = _
        "***LOGIN***" & vbCrLf & _
        "sessionid=" & SessionID & vbCrLf & _
        "appname=" & AppName & vbCrLf & _
        "sessionimeout=" & SessionTimeout & vbCrLf & _
        "key=" & Key

    frmMain.txtLCD(0).Text = AppName
    frmMain.txtLCD(1).Text = SessionID
       
    sLine(0) = frmMain.txtLCD(0).Text
    sLine(1) = frmMain.txtLCD(1).Text
    Call SendString(sLine)
    
    da_aggiungere = True
    If g_Applications(0).SessionID = -1 Then
        J = 0
        da_aggiungere = False
    Else
        For I = LBound(g_Applications) To UBound(g_Applications)
            If g_Applications(I).SessionID = SessionID Then
                ' esiste gia' -> rimpiazzo
                
                J = I
                
                da_aggiungere = False
                Exit For
            End If
        Next
    End If
    
    If da_aggiungere Then
        J = UBound(g_Applications) + 1
        ReDim Preserve g_Applications(J)
    End If
    With g_Applications(J)
        .SessionID = SessionID
        .SessionTimeout = SessionTimeout
        .Name = AppName
        .Key = Key
    End With
    
    With p_Response
        .bResponse = True
        .lResponse = -1
        .sResponse = "Logged in"
    End With
    
    'MsgBox "login: (" & SessionID & "," & appname & "," & SessionTimeout & "," & Key & ")"
    Call FillLstApps(frmMain.lstApps, g_Applications)
    
End Sub

Public Sub Logout(ByVal SessionID As Long)
    Dim I As Long, J As Long
    Dim sLine(1) As String
    
    frmMain.txtData2.Text = _
        "***LOGOUT***" & vbCrLf & _
        "sessionid=" & SessionID
    
    frmMain.txtLCD(0).Text = "LOGOUT " & SessionID
    frmMain.txtLCD(1).Text = ""
    
    sLine(0) = frmMain.txtLCD(0).Text
    sLine(1) = frmMain.txtLCD(1).Text
    
    Call SendString(sLine)
    
    For I = LBound(g_Applications) To UBound(g_Applications)
        If g_Applications(I).SessionID = SessionID Then
            J = I
            Exit For
        End If
    Next
    
    For I = J + 1 To UBound(g_Applications)
        g_Applications(I - 1) = g_Applications(I)
    Next
    
    If UBound(g_Applications) > 0 Then
        ReDim Preserve g_Applications(UBound(g_Applications) - 1)
    Else
        g_Applications(0).SessionID = -1
    End If
    
    With p_Response
        .bResponse = True
        .lResponse = -1
        .sResponse = "Logged out"
    End With
    
    'MsgBox "logout: (" & SessionID & ")"
    Call FillLstApps(frmMain.lstApps, g_Applications)
End Sub

Public Sub SendData(ByVal SessionID As Long, ByVal data As String, ByVal datatype As Long)
    Dim I As Long
    Dim S(1) As String
    Dim sLine(1) As String
    
    frmMain.txtData2.Text = _
        "***SENDCOMMAND***" & vbCrLf & _
        "sessionid=" & SessionID & vbCrLf & _
        "data=" & data & vbCrLf & _
        "datatype=" & datatype
    
    If SessionID = g_Applications(UBound(g_Applications)).SessionID Then
        Select Case True
            Case UCase$(Left$(data, 5)) = "ROWS="
                data = Mid$(data, 6)
                I = InStr(1, data, "|")
                If I = 0 Then
                    S(0) = data
                Else
                    S(0) = Left$(data, I - 1)
                    S(1) = Mid$(data, I + 1)
                End If
        
        End Select
        
        frmMain.txtLCD(0).Text = S(0)
        
        frmMain.txtLCD(1).Text = S(1)
        
        sLine(1) = frmMain.txtLCD(1).Text
        sLine(0) = frmMain.txtLCD(0).Text
        Call SendString(sLine)
        
        With p_Response
            .bResponse = True
            .lResponse = -1
            .sResponse = "Data sent"
        End With
    Else
        With p_Response
            .bResponse = False
            .lResponse = 0
            .sResponse = "Data not sent (not logged in)"
        End With
    End If

        
    'MsgBox "data (" & SessionID & "," & data & "," & datatype & ")"
End Sub

Public Sub SendCommand(ByVal SessionID As Long, ByVal cmd As String, ByVal cmdtype As Long)
    Dim unKnown As Boolean
    
    frmMain.txtData2.Text = _
        "***SENDCOMMAND***" & vbCrLf & _
        "sessionid=" & SessionID & vbCrLf & _
        "cmd=" & cmd & vbCrLf & _
        "cmdtype=" & cmdtype
    
    unKnown = False
    Select Case cmd
        Case "SHOWUI":
            If Not frmMain.Visible Then
                Call AlwaysOnTop(frmMain.hWnd, frmMain.Width, frmMain.Height, frmMain.Top, frmMain.Left, True)
                frmMain.Show
            End If
            
        Case "HIDEUI":
            If frmMain.Visible Then frmMain.Hide
        
        Case Else
            unKnown = True
    End Select
    
    With p_Response
        If unKnown Then
            .bResponse = False
            .lResponse = 0
            .sResponse = "Unknown command received"
        Else
            .bResponse = True
            .lResponse = -1
            .sResponse = "Command processed"
        End If
    End With
    
    
    'MsgBox "data (" & SessionID & "," & data & "," & datatype & ")"
End Sub

Private Sub Class_Initialize()
    
    Dim sSettings As String
    ReDim g_Applications(0)
    g_Applications(0).SessionID = -1
    
    sSettings = GetIni(App.Path & "\" & ININAME, "COM", "BAUDRATE", "9600") & "," & _
                GetIni(App.Path & "\" & ININAME, "COM", "PARITY", "N") & "," & _
                GetIni(App.Path & "\" & ININAME, "COM", "DATABIT", "8") & "," & _
                GetIni(App.Path & "\" & ININAME, "COM", "STOPBIT", "1")
    
    If frmMain.ComCtrl.PortOpen = False Then
        frmMain.ComCtrl.CommPort = CInt(GetIni(App.Path & "\" & ININAME, "COM", "COMNUMBER", "1"))
        frmMain.ComCtrl.Settings = sSettings
        frmMain.ComCtrl.InputLen = 0
        frmMain.ComCtrl.PortOpen = True
    End If
End Sub

Private Sub Class_Terminate()
    frmMain.ComCtrl.PortOpen = False
End Sub
