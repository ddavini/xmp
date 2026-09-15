VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form FrmMicroweb 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   420
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Microweb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin NTService.NTService NTService 
      Left            =   480
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "x.m Microweb Server"
      ServiceName     =   "Microweb"
      StartMode       =   3
   End
   Begin MSWinsockLib.Winsock Principale 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMicroweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Si ringrazia Roberto Negro per il suo HTTP Virtual Server demo project (C) 1999
'Da cui ha tratto spunto Microweb server
'
'Nota: per utilizzare il Progetto in debug asteriscare la Chiamata alla Sub NTSvc
'Oppure usare come command line -debug
'Nota1: La presente vesione di Microweb supporta il solo comando GET
Public Pila As Integer
Public PilaLock As Boolean

Private Sub Form_Load()
On Error GoTo Errore
    NTSvc
           
    Call StartUP_Mweb
    
    Principale.LocalPort = GetIni(ConfigFileName, "MAIN", "PORT", "80")
    Principale.Bind Principale.LocalPort, GetIni(ConfigFileName, "MAIN", "BINDING", "0.0.0.0")
    Principale.Listen
    
    ChDir App.path
    AllocCmd
    Clrscr
    
    Writeln "x.m MicroWeb version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " (" + StableStr + ") by www.zolnetwork.com"
    Writeln "Listen On: " + CStr(Principale.LocalIP) + " Prort: " + CStr(Principale.LocalPort)
    Writeln
    
    Exit Sub
Errore:
    Call ErrorGST(Err.Description + " - " + "Microweb")
    Unload Me
    End
End Sub

Private Sub StartUP_Mweb()
Dim Esiste As String
Dim MWeb_Temp As String
Dim MWeb_Logs As String
Dim MWeb_Root As String
Const IniFileName As String = "\config.ini"

On Error GoTo Errore

    ConfigFileName = App.path & IniFileName

    MWeb_Temp = App.path & GetIni(ConfigFileName, "MAIN", "TMPFOLDER", "\Tmp")
    MWeb_Logs = App.path & GetIni(ConfigFileName, "MAIN", "LOGFOLDER", "\Log")
    MWeb_Root = App.path & GetIni(ConfigFileName, "MAIN", "WEBFOLDER", "\WebRoot")
    
    Esiste = Dir(MWeb_Temp, vbDirectory)
    If Esiste = "" Then
        Call MkDir(MWeb_Temp)
    Else
        Esiste = Dir(MWeb_Temp & "\*.$", vbNormal)
        If Esiste <> "" Then
            Call FileSystem.Kill(MWeb_Temp & "\*.$")
        End If
    End If
    
    Esiste = Dir(MWeb_Logs, vbDirectory)
    If Esiste = "" Then
        Call MkDir(MWeb_Logs)
    End If
    
    Esiste = Dir(MWeb_Logs, vbDirectory)
    If Esiste = "" Then
        Call Err.Raise(666, Description:="No Root Directory!")
    End If
    
    LogsDir = MWeb_Logs
    TempDir = MWeb_Temp
    RootDir = MWeb_Root
    
    Exit Sub
    
Errore:
    Call ErrorGST(Err.Description + " - " + "StartUP_Mweb")
    Unload Me
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FreeCmd
End Sub

Private Sub Principale_ConnectionRequest(ByVal requestID As Long)
On Error GoTo Errore:
    
    Dim ServerSND As FrmSender
    Dim TimeOut As Long
    Dim MaxReq As Integer
    Dim strLog As String
    
    MaxReq = CLng(GetIni(ConfigFileName, "MAIN", "MAXQREQ", "25"))
    If Pila < MaxReq Then
                
        Set ServerSND = New FrmSender
        
        strLog = CStr(Date) + SepTag + CStr(Time) + SepTag + Principale.RemoteHostIP
        strLog = strLog + SepTag + Principale.RemoteHost + SepTag + Principale.LocalHostName
        strLog = strLog + SepTag + Principale.LocalIP
        Writeln "Connection From: " + Principale.RemoteHostIP + " For: " + Principale.LocalHostName
        
        'Debug.Print "Principale :" & Principale.State
        If BannedIP(Principale.RemoteHostIP) Then
            Writeln "[Deny Ip Banned]"
            ScriviLOG (strLog & " " & "Deny Ip Banned")
            Set ServerSND = Nothing
        Else
            ServerSND.strLog = strLog
            ServerSND.FrmSenderCreate = True
            ServerSND.WsSender.Accept requestID
        End If
    Else
        Principale.Close
        If CBool(GetIni(ConfigFileName, "MAIN", "EVLOG", "False")) Then
            Call App.LogEvent("Too Many Request - " & Pila)
        End If
        Writeln "Too Many Request - " & Pila
    End If

Exit Sub
Errore:
    Call ErrorGST(Err.Description + " - " + "Princpale.ConnectioReq")
End Sub

'''''''''''''''''''''''NT SERVICE ''''''''''''''''''''''''''''
Private Sub NTSvc()
    Dim strDisplayName As String
    Dim bStarted As Boolean
    
On Error GoTo Errore
    
    strDisplayName = "MicroWeb"
       
    If Command = "-install" Then
        NTService.Interactive = False
        
        If NTService.Install Then
            Call NTService.SaveSetting("Parameters", "TimerInterval", "1000")
            Writeln (strDisplayName & " Istallato con successo")
        Else
            Writeln strDisplayName & " Installazione fallita"
        End If
        End
    ElseIf Command = "-remove" Then
        If NTService.Uninstall Then
            Writeln strDisplayName & " Disistallato con successo"
        Else
            Writeln strDisplayName & " Disistallazione fallita"
        End If
        End
    ElseIf Command = "-debug" Then
        NTService.Debug = True
    ElseIf Command <> "" Then
        Writeln "Comando non valido"
        End
    End If


    NTService.ControlsAccepted = svcCtrlStartStop
    NTService.StartService
    
Exit Sub
Errore:
    If NTService.Interactive Then
        Writeln "[" & Err.Number & "] " & Err.Description
        End
    Else
        Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
        End
    End If
End Sub

Private Sub NTService_Control(ByVal e As Long)
On Error GoTo Errore
    
    Writeln CStr(e)
    
Exit Sub
Errore:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub

Private Sub NTService_Start(Success As Boolean)
On Error GoTo Errore
    
    Success = True

Exit Sub
Errore:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


Private Sub NTService_Stop()
On Error GoTo Errore
    
    Unload Me
    
Exit Sub
Errore:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub
'''''''''''''''


