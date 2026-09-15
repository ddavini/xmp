Attribute VB_Name = "modMain"
Option Explicit

Private Sub Main()

    Dim Minimizzato As Boolean
    Dim Comando As String
    Dim ListFile As String
    
On Error GoTo ErrH
        
    ChDrive App.Path
    ChDir App.Path
        
    If App.PrevInstance Then
        Call PassaStrToPrevIstance(Command$)
        End
    End If
    
    'Setta file CFG
    cfgFile = App.Path + "\xmp.ini"
    Call LeggiCFG
       
    ''' Gestione Font '''
    Call AddFont
    'FONT1'
    frmListone.ListaMp3.Font = GetINI(cfgFile, "VISUALIZATION", "FONTNAME", "TAHOMA")
    
    Call xmpPensa(True)
    
    WriteINFO "Starting..."
    Call GestisciPosFrm(True, Minimizzato)
           
    If Command$ <> "/debug" Then
        WriteINFO "Setup UnhandledException..."
        SetUnhandledExceptionFilter AddressOf MyExceptionFilter
    End If

    Call InitFFT
                   
    Modo = -1
    Call InitStreamEngine
       
    WriteINFO "XmP Start Up..."
    
    If CBool(GetINI(cfgFile, "PREFERENCE", "MH", "FALSE")) Then
            WriteINFO "LCD LogIN..."
            Call modMicroHandler.SendCommand(LOGIN)
    End If
    
    
    xmp.UnSec.Interval = 1
    frmListone.frmListoneTimer.Interval = 1
    
    ListFile = GetINI(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
    Call LeggiListaMp3(ListFile)
    
    If Minimizzato Then
        MinimizzaXMP
    End If
    
    If Trim$(Command$) <> "" And _
       InStr(1, Command$, "/") = 0 Then
        Select Case LCase(Right$(Command$, 3))
            Case Is = ".ls"
                PulisciTutto
                ListFile = Right$(Command$, _
                Len(Command$) - InStrRev(Command$, " "))
                Call LeggiListaMp3(ListFile)
            Case Is = "m3u"
                PulisciTutto
                ListFile = Right$(Command$, _
                Len(Command$) - InStrRev(Command$, " "))
                Call LeggiListaM3U(ListFile)
            Case Else
                Call CommandMP3(Command$)
        End Select
    End If
    
    Call ICONtray(xmp, xmp.Name)
    
    Call frmSetFocus(xmp)
    
    Call xmpPensa(False)
    
    Call InfoNmp3
        
    WriteINFO "All System Active"
    Load frmVisualizzazioni
    frmVisualizzazioniAttivo = True
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Main" & vbCrLf & Err.Description)
        Call DeLoad(Brutal:=True)
    End If
End Sub

