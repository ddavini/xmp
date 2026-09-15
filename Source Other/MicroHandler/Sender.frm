VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FrmSender 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Sender.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrSender 
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TimeOut 
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock WsSender 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private IncomingData As String
Private Packet As String
Private Closing As Boolean
Private CloseOK As Boolean
Private FileData As String

Public strLog As String
Public FrmSenderCreate As Boolean

Private Sub Form_Load()
On Error Resume Next
   ' Me.Visible = True
    Closing = False
    If Not FrmSenderCreate Then
        Me.Print "Pacco"
        Call ScriviLOG("Form Pacco")
        Unload Me
        Exit Sub
    Else
        Me.Print strLog
    End If
    While FrmMicroweb.PilaLock
        Pause
    Wend
    With FrmMicroweb
        .PilaLock = True
        .Pila = .Pila + 1
'        Debug.Print "Pila:" & .Pila
        .PilaLock = False
    End With
    TimeOut.Interval = 5000
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim MaxReq As Integer
    
    If Closing Then
        MaxReq = CLng(GetIni(ConfigFileName, "MAIN", "MAXQREQ", "25"))
        While FrmMicroweb.PilaLock
            Pause
        Wend
        With FrmMicroweb
            .PilaLock = True
            .Pila = .Pila - 1
            If .Pila < MaxReq Then
                FrmMicroweb.Principale.Listen
            End If
'            Debug.Print "Pila:" & .Pila
            .PilaLock = False
        End With
    End If
  '  Set FrmSender = Nothing
End Sub

Private Sub TimeOut_Timer()
On Error Resume Next
    Me.Print "Timeout"
    If WsSender.State <> sckConnected And WsSender.State <> sckClosing Then
        If CBool(GetIni(ConfigFileName, "MAIN", "EVLOG", "False")) Then
            Call App.LogEvent("Request Timed Out - Sender", vbLogEventTypeError)
        End If
        Writeln "Request Timed Out - Sender.TimeOut_Timer"
        Call CloseSND
    End If
End Sub

Private Sub tmrSender_Timer()
Dim BandT As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Nota: IE (6 SP1) gestisce 2 download per sessione NON E' UN LIMITE di MW'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
BandT = CLng(GetIni(ConfigFileName, "MAIN", "BWT", "640"))
BandT = BandT * 100

On Error GoTo Errore
    
    'Come mai due IF identici?... in ottica di ottimizzazione per evitare di aspettare
    'inutilmente 1s solo per chiudere la connessione... non bello ma efficace.
    'sicuramente ci sara' anche un modo piu' elegante ma al momento non mi vien altro.
    If Len(FileData) > 0 Then
        'Debug.Print WsSender.State
         tmrSender.Interval = 1000
         If Len(FileData) < BandT Then
             BandT = Len(FileData)
         End If
         CloseOK = False
         WsSender.SendData Mid$(FileData, 1, BandT)
         FileData = Mid$(FileData, BandT + 1)
    End If
    
    'Permette lo scatenarsi della _SendConplete
    DoEvents
    
    If Len(FileData) > 0 Then
        Exit Sub
    Else
        'tmrSender.Interval = 0
        If CloseOK Then
            Call CloseSND
        End If
    End If
    
    Exit Sub
Errore:
    Call ErrorGST(Err.Description + " - " + "Sender.tmrSender_Timer")
    Call CloseSND
End Sub

Private Sub WsSender_SendComplete()
   CloseOK = True
End Sub

'Private Sub WsSender_Close()
    'Call CloseSND(Me)
'End Sub

Private Sub WsSender_DataArrival(ByVal bytesTotal As Long)
Dim Posizione As Integer
Dim Lunghezza As Integer

On Error Resume Next
    WsSender.GetData Packet, vbString
    
    IncomingData = IncomingData & Packet
        
    Posizione = InStr(1, IncomingData, vbCrLf & vbCrLf, vbTextCompare)
    Lunghezza = 4
    If Posizione = 0 Then
        Posizione = InStr(1, IncomingData, vbCr & vbCr, vbTextCompare)
        Lunghezza = 2
        If Posizione = 0 Then
            Posizione = InStr(1, IncomingData, vbLf & vbLf, vbTextCompare)
        End If
    End If

    If Posizione <> 0 Then
        Call Elabora(IncomingData)
    End If
    
End Sub

Private Sub WsSender_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFilez As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
'    If Number = sckConnectionReset Then
'     '   Call CloseSND
'    End If
    If CBool(GetIni(ConfigFileName, "MAIN", "EVLOG", "False")) Then
        Call App.LogEvent(Description & " - " & "Sender.WsSender_Error", vbLogEventTypeError)
    End If
    Writeln Description + " - " + "Sender.WsSender_Error"
End Sub
Private Sub Elabora(ByVal StrCommand As String)
    Dim Corpo As String
    Dim Filez As String, Footer As String, FileType As String, Header As String
    Dim HTTPerr As String, er As String
    Dim Esiste As String
    Dim Dimensione As Long
    Dim Host As String
    Dim Pos As Integer
    Dim UserAgent As String
    Dim Referrer As String
    Dim StrQuery As String
    Dim StrPost As String
    Dim QryQuestion As String
    Dim RequestMethod As String
    Dim vHeader() As String
    Dim vTMP() As String
    Dim I As Byte
    Dim GZIPEnabled As Boolean
    Dim CloseHearder As Boolean
    
On Error GoTo Errore
       
    vHeader = Split(StrCommand, vbCrLf)
        
    For I = 0 To UBound(vHeader)
        If vHeader(I) <> "" And vHeader(I) <> vbCrLf Then
            ReDim Preserve vTMP(I)
            vTMP(I) = vHeader(I)
        End If
    Next I
    Erase vHeader
    vHeader = vTMP
    Erase vTMP
    
    Host = KeyValue(vHeader, "host")
    UserAgent = KeyValue(vHeader, "user-agent")
    Referrer = KeyValue(vHeader, "referer")
    If Referrer = "" Then
        Referrer = "Direct"
    End If
    
    If InStr(1, KeyValue(vHeader, "accept-encoding"), "gzip", vbTextCompare) <> 0 Then
        GZIPEnabled = True
    End If
    'Set Server Variable'
    
    SetEnvironmentVariable "HTTP_ACCEPT_LANGUAGE", KeyValue(vHeader, "accept-language")
    SetEnvironmentVariable "HTTP_ACCEPT", KeyValue(vHeader, "accept")
    SetEnvironmentVariable "HTTP_COOKIE", KeyValue(vHeader, "cookie")
    SetEnvironmentVariable "REMOTE_ADDR", FrmMicroweb.Principale.RemoteHostIP
    SetEnvironmentVariable "HTTP_USER_AGENT", UserAgent
    SetEnvironmentVariable "HTTP_REFERER", Referrer
                            
    SetEnvironmentVariable "CONTENT_LENGTH", KeyValue(vHeader, "content-length")
    SetEnvironmentVariable "CONTENT-TYPE", KeyValue(vHeader, "content-type")
    
    vTMP = Split(vHeader(0), " ")
    RequestMethod = vTMP(0)
    SetEnvironmentVariable "REQUEST_METHOD", RequestMethod
    SetEnvironmentVariable "SERVER_PROTOCOL", vTMP(UBound(vTMP))
    
    SetEnvironmentVariable "GATEWAY_INTERFACE", "CGI/1.1"

    SetEnvironmentVariable "SERVER_SOFTWARE", "x.m MicroWeb version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " (" + StableStr + ") by www.zolnetwork.com"
    
    
    'PHP_DOCUMENT_ROOT"
    'DOCUMENT_ROOT
    'SERVER_PROTOCOL = "HTTP/1.0"
    'SERVER_ADDR
    'SCRIPT_FILENAME
    'SCRIPT_NAME
    'SERVER_NAME
    'SERVER_PORT
    
    If InStr(1, StrCommand, "http", vbTextCompare) Then
        StrCommand = Left(StrCommand, InStr(1, StrCommand, "http", vbTextCompare) - 1)
    End If
    If LCase$(vTMP(0)) = "get" Or LCase$(vTMP(0)) = "post" Then
        StrCommand = Mid(StrCommand, InStr(1, StrCommand, vTMP(0) & " ", vbTextCompare) + 4)
    Else
        StrCommand = ""
    End If
    
    'Repulisti, mai fatto un programma con tanti martelli
    'se' se', fosse vero ;)
    StrCommand = Trim(StrCommand)
    StrCommand = Replace(StrCommand, vbCr, "")
    StrCommand = Replace(StrCommand, vbLf, "")

    If StrCommand <> "" Then
        If StrCommand = "/" Then
            StrCommand = "/default.htm"
        ElseIf InStr(1, StrCommand, ".", vbTextCompare) = 0 Then
            Select Case Right$(StrCommand, 1)
                Case "/", "\"
                    StrCommand = StrCommand + "default.htm"
                Case Else
                    HTTPerr = "302"
                    StrCommand = StrCommand + "/"
            End Select
        End If
        
        If LCase$(RequestMethod) = "post" Then
            StrPost = vHeader(UBound(vHeader))
        Else
            Pos = InStr(1, StrCommand, "?", vbTextCompare)
            If Pos <> 0 Then
                StrQuery = ExtractQueryString(StrCommand, Pos)
            End If
        End If
        'Elaborazione Rischiesta (un tantinello maccheronico)
        If HTTPerr = "" Then
            Filez = StrCommand

            Filez = UnicodeTrans(Filez)
            Esiste = Trim$(Dir(RootDir & Filez))
            If Esiste <> "" And Not DenyUrlStr(Filez, strLog) Then
                Dimensione = FileLen(RootDir & Filez)
                If Dimensione <= CLng(GetIni(ConfigFileName, "MAIN", "MAXFILEDIM", "2097152")) Then
                        
                        'Lettura file da disco
                        'Il tutto avviene in una sola passata ecco il perche del MAXFILEDIM
                        'Questa parte di Microweb andrebbe ingegnerizzata a dovere per
                        'permettere una pacchettizzazione del trasferimento...
                        'Probabilmente prima o poi lo faro', se pero' qualcuno lo fa prima di
                        'me e' pregato di recapitarmi il codice.
                        Corpo = ExtCtrl(StrCommand, StrQuery, StrPost, _
                                        Filez, WsSender.SocketHandle, GZIPEnabled)
                        
                        HTTPerr = "200"
                        FileType = ""
                        If InStr(1, Filez, ".", vbTextCompare) > 0 Then
                            FileType = RetFileType(Filez)
                        End If
                Else
                     HTTPerr = "404"
                     FileType = "text/html"
                End If
            Else
                HTTPerr = "404"
                FileType = "text/html"
            End If
        End If
    Else
         HTTPerr = "400"
         FileType = "text/html"
    End If
    

    'Parsing Speciali o eventuale Retrive della pagina di errore'
    If HTTPerr = "200" Then
        Call Parser(Corpo, Filez)
    Else
        Corpo = RetHtmlHttpErr(HTTPerr, Host + StrCommand)
    End If
    
    If InStr(1, Left$(Corpo, 1024), "HTTP/", vbTextCompare) = 0 Then
        'Costruzione intestazione ed eventuale Foot della pagina
        Header = "HTTP/1.1 " & Format(HTTPerr) & " Document followSender" & vbCrLf
               
        If Host <> "" Then
            If HTTPerr = "302" Then
                Header = Header + "Location: http://" + Host + StrCommand + vbCrLf
            Else
                Header = Header + "Content-Location: http://" + Host + StrCommand + vbCrLf
            End If
        End If
        Header = Header & "Server: Microweb" + vbCrLf
        If FileType <> "" Then
            Header = Header & "Content-type: " & FileType & vbCrLf
        End If
    
        
        If (InStr(1, Left$(Corpo, 1024), "text/html", vbTextCompare) <> 0) Or (FileType = "text/html") Then
            Footer = "<p>&nbsp;</p><p align=left><HR><b>This Site is based on a x.m microweb server version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " (" + StableStr + ") by <a href=http://www.zolnetwork.com>ZOLNetwork</a></b></p>"
            Select Case GetIni(ConfigFileName, "MAIN", "FOOTER", "1")
                Case Is = "1"
                    Footer = vbCrLf & "<font color=black size=0>" + Footer + "</font>"
                Case Is = "2"
                    Footer = vbCrLf & "<font color=white size=0>" + Footer + "</font>"
                Case Else
                    Footer = ""
            End Select
        End If
        
        If (InStr(1, Left$(Corpo, 1024), "Content-type: ", vbTextCompare) <> 0) Then
            GZIPEnabled = False
            CloseHearder = False
        Else
            CloseHearder = True
        End If
        
        Corpo = Corpo & Footer
        
        If GZIPEnabled And FileType = "text/html" Then
            Header = Header & "Content-Encoding: gzip" & vbCrLf
            Corpo = GZIPCompress(Corpo, WsSender.SocketHandle)
            Header = Header & "Accep-Ranges: bytes" & vbCrLf
            Header = Header & "Vary: Accept-Encoding" & vbCrLf
        Else
            GZIPEnabled = False
        End If
        
        Header = Header & "Content-length: " & Format(Len(Corpo)) & vbCrLf
        Header = Header & "Connection: close"
         
        If CloseHearder Then
            Header = Header & vbCrLf & vbCrLf
        End If

    End If
    
    'Scrittura File di Log
    If StrQuery <> "" Then
        QryQuestion = "?"
    Else
        QryQuestion = ""
    End If
    
    strLog = strLog + SepTag & "GZIP_" & GZIPEnabled & SepTag & StrCommand + QryQuestion + StrQuery + SepTag + HTTPerr + SepTag + Host + _
                   SepTag + UserAgent + SepTag + Referrer
    Writeln "Connection Request: " & "GZIP_" & GZIPEnabled & SepTag & StrCommand & QryQuestion & StrQuery & SepTag & HTTPerr
    Call ScriviLOG(strLog)
    strLog = ""
    'Fine Scrittura Log File
        
    'Spedizione Dati in una connessione TCP singola'
    FileData = Header & Corpo
    'TimeOut.Enabled = False
    
    'Non so perche' ma senza questa rifa' due volte la elabora!!
    'Probabilmente la doevents permette lo scatenarsi di qualcosa
    'che altrimenti richiamerebbe l'elaborazione
    'Che codice marcio!
    DoEvents

    Call tmrSender_Timer
    'tmrSender.Interval = 1
    'WsSender.SendData Header & Corpo & Footer
    
    Exit Sub
Errore:
    Call ErrorGST(Err.Description + " - " + "Sender.Elabora")
    Call CloseSND
End Sub

Public Sub CloseSND()
On Error Resume Next
    If Not Closing Then
        FileData = ""
        Closing = True
        tmrSender.Enabled = False
        TimeOut.Enabled = False
        WsSender.Close
        Call Pause(True, 0)
        Unload Me
    End If
End Sub

Private Function KeyValue(vHeader() As String, KeyName As String) As String
    Dim I As Integer
    Dim Pos As Integer
    
    For I = 0 To UBound(vHeader)
        Pos = InStr(1, vHeader(I), KeyName & ": ", vbTextCompare)
        If Pos <> 0 Then
            Pos = Pos + Len(KeyName) + 2
            KeyValue = Mid$(vHeader(I), Pos, Len(vHeader(I)) - Pos + 1)
            Exit For
        End If
    Next I
    

End Function
