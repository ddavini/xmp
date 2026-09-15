Attribute VB_Name = "modFunzioniGlobali"
Option Explicit

Public Function ExistTmpFile() As Boolean
On Error GoTo ErrH

    ExistTmpFile = FileExists(App.Path & "\" & Mp3TmpFiles)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ExistTmpFile" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Function

Public Sub DelFile(Index As Integer)
Dim Number As Integer
On Error GoTo ErrH
    
    Number = Index \ 100
    
    Call DelINIKey(App.Path & "\" & Mp3TmpFiles, "FILES" & Number, "F" & Index)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.DelFile" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Sub


Public Function GetFile(Index As Integer, _
                    Optional APPLICATION As String = "FILES", _
                    Optional INIFileName As String) As String
                    
    Dim Number As Integer
On Error GoTo ErrH
    
    If INIFileName = "" Then
        INIFileName = App.Path & "\" & Mp3TmpFiles
    End If
    Number = Index \ 100
    
    GetFile = GetINI(INIFileName, APPLICATION & Number, Left$(APPLICATION, 1) & Index, 0)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.GetFile" & _
                vbCrLf & "Index: " & CStr(IndiceGlobalissimo) & _
                vbCrLf & "FileName: " & INIFileName & _
                vbCrLf & "APP: " & APPLICATION & _
                vbCrLf & Err.Description)
    End If
End Function

Public Sub SetFile(Index As Integer, Mp3File As String, _
                    Optional APPLICATION As String = "FILES", _
                    Optional INIFileName As String)
Dim Number As Integer
On Error GoTo ErrH
       
    Dim fServer As Boolean
    
    If INIFileName = "" Then
        INIFileName = App.Path & "\" & Mp3TmpFiles
    End If
    
    Number = Index \ 100

    Call SetINI(INIFileName, APPLICATION & Number, Left$(APPLICATION, 1) & Index, Mp3File)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SetFile" & _
                vbCrLf & "Index: " & CStr(IndiceGlobalissimo) & _
                vbCrLf & "FileName: " & INIFileName & _
                vbCrLf & "APP: " & APPLICATION & _
                vbCrLf & Err.Description)
    End If
End Sub

Public Function NFile(Optional INIFileName As String, _
                      Optional APPLICATION As String = "FILES") As Integer
Dim Apps() As String
Dim Store() As String
Dim I As Integer
Dim NumeroFile As Integer
On Error GoTo ErrH
    
    If INIFileName = "" Then
        INIFileName = App.Path & "\" & Mp3TmpFiles
    End If

    Apps = GetINIAllApps(INIFileName)
    
    If isVector(Apps) Then
        For I = 0 To UBound(Apps)
            If InStr(1, Apps(I), APPLICATION, vbTextCompare) > 0 Then
                Store = GetINIAppAllKeys(INIFileName, Apps(I))
                NumeroFile = NumeroFile + UBound(Store) + 1
            End If
        Next I
        NFile = NumeroFile - 1
    Else
        NFile = -1
    End If
    
    Erase Apps
    Erase Store
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Indice: " & CStr(IndiceGlobalissimo) & _
                vbCrLf & "FunzioniGlobali.NFile" & vbCrLf & Err.Description)
    End If
End Function

Function GetNormFromDisk(Index As Integer) As Double
Dim Number As Integer
On Error GoTo ErrH
    
    Number = Index \ 100
    
    GetNormFromDisk = CDbl(GetINI(App.Path & "\" & Mp3TmpFiles, "XNORM" & Number, "N" & Index, "1"))
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.GetNormFromDisk" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Function

Public Sub SetNormToDisk(Index As Integer, Optional Normalize As Double = 1)
Dim Number As Integer
On Error GoTo ErrH
    
    Number = Index \ 100
    
    Call SetINI(App.Path & "\" & Mp3TmpFiles, "XNORM" & Number, "N" & Index, Normalize)
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SetNormToDisk" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Sub


Public Function LoadDataIntoFile(DataName As String) As IPictureDisp
On Error GoTo ErrH
    Dim myArray() As Byte
    Dim myFile As Long
    Dim FileName As String

    Call Randomize
    FileName = App.Path & "\" & App.hInstance & Rnd * 50 & Rnd * 50 & ".$$$"
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
            Put #myFile, , myArray
        Close #myFile
    End If
    Set LoadDataIntoFile = LoadPicture(FileName)
    Call Kill(FileName)
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.LoadDataIntoFile" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Function

Public Sub MoveInMp3(ByVal Direzione As Integer)
On Error GoTo ErrH
    If ExistTmpFile Then
            IndiceGlobalissimo = IndiceGlobalissimo + Direzione
            Select Case IndiceGlobalissimo
                Case Is < 0
                    IndiceGlobalissimo = NFile
                Case Is > NFile
                    IndiceGlobalissimo = 0
            End Select
            frmListone.ListaMp3.ListIndex = IndiceGlobalissimo
    
            If mStreamIsActive Then
                PlayStream
            End If
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.MoveInMp3" & vbCrLf & Err.Description _
                & vbCrLf & "Indice: " & CStr(IndiceGlobalissimo))
    End If
End Sub

Public Sub PlayStream(Optional Mp3 As Integer = -1)
On Error GoTo ErrH
    Dim STR As String
    Dim ListFile As String
    Dim fServer As Boolean
       
    If Mp3 >= 0 Then
        IndiceGlobalissimo = Mp3
        frmListone.ListaMp3.ListIndex = Mp3
    End If
    
    With xmp
        Call xmpPensa(True)
        If ExistTmpFile Then
            If IndiceGlobalissimo <= 0 Or _
               IndiceGlobalissimo >= NFile + 1 Then
                frmListone.ListaMp3.ListIndex = 0
                IndiceGlobalissimo = 0
            End If
            
            Call StopAll
            
            frmListone.ListaMp3.ListIndex = IndiceGlobalissimo
            WriteINFO "Loading Stream..."
            
            Call ICONtray(xmp, "Loading Stream...", NIM_MODIFY)
                                  
            If mOpenStream(GetFile(IndiceGlobalissimo)) Then
                        
                If frmMenu.mnuXNorm.Checked Then
                    Call mXNormLevel(GetNormFromDisk(IndiceGlobalissimo))
                End If
                
                Call SettaNomeMp3
                
                Call SettaIndicatoreModo
                
                .Durata.xCaption = DurataStream(mStreamLenInSeconds)
                If mGetMP3Info.VBR Then
                    .BitRate.xCaption = "VBR"
                Else
                    .BitRate.xCaption = CStr(mGetMP3Info.BitRate) + "k"
                End If
                
                .lblFreq.xCaption = CStr(mGetMP3Info.Frequency \ 1000) + "khz"
                frmListone.ListaMp3.ModifyItem IndiceGlobalissimo, _
                DurataStream(mStreamLenInSeconds) & " - " & .xmDisplay(0).xCaption
                            
                
                ListFile = GetINI(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
                Call ICONtray(xmp, .xmDisplay(0).xCaption + " - " + .Durata.xCaption, NIM_MODIFY)
                
                Call InitFFT
                
                Call modMicroHandler.SendCommand(SNDDATA, "ROWS=*** " & .xmDisplay(0).xCaption + " ***|" + .Durata.xCaption)
                
                If CBool(GetINI(cfgFile, "PREFERENCE", "SAVEALS", "True")) Then
                    Call ScriviListaMp3(ListFile)
                Else
                    ScriviPosizioneListaMp3
                End If
                
                mPlayStream
              
                If GetFile(IndiceGlobalissimo) = gPosForSave.Mp3File Then
                    Call mStreamPos(gPosForSave.SamplePos)
                    gPosForSave.Mp3File = ""
                End If
                
                If .Mute Then
                    Call mStreamVol(0)
                    .xmsVol.xValue = .xmsVol.xMax
                Else
                    Call mStreamVol(.xmsVol.xMax - .xmsVol.xValue)
                End If
                
                If Trim(mGetMP3Info.Artist) <> "" Then
                    WriteINFO mGetMP3Info.Artist
                Else
                    WriteINFO "Unknown"
                End If
                
                Call InfoNmp3
            Else
                fServer = CBool(GetINI(cfgFile, "PREFERENCE", "SERVER", "False"))
                gPosForSave.Mp3File = ""
                gPosForSave.SamplePos = 0
                xmp.g_PlayDone = False
                Call StopAll
                If fServer Then
                    Call Pause(10)
                    Call MoveInMp3(1)
                    Call PlayDone
                    Call PlayStream(frmListone.ListaMp3.ListIndex)
                End If
            End If
        End If
    End With
    
    'Scarico frmInfo che salcazzo perche' si carica / Unload frmInfo - who the hell knows why it gets loaded
    Unload frmInfo
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.PLAYStream" & vbCrLf & Err.Description)
    End If
    Call xmpPensa(False)
End Sub



Public Function DurataStream(StreamLenInSeconds As Integer) As String
On Error GoTo ErrH
    Dim Minuti, Secondi As Byte
            If StreamLenInSeconds < 0 Then
                StreamLenInSeconds = 0
            End If
            Minuti = (StreamLenInSeconds \ 60)
            Secondi = StreamLenInSeconds Mod 60
            DurataStream = Format(CStr(Minuti), "00") + ":" + Format(CStr(Secondi), "00")
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.DurataStream" & vbCrLf & Err.Description)
    End If
End Function

Public Sub AgganciaSub(Optional Tieni As Boolean)
On Error GoTo ErrH
    Const Sensibilta = 500
    Dim X, Y As Long
      
    X = frmListone.Left
    Y = frmListone.Top
    AgganciatoFlag = False
    If Tieni = False Then
        If X > (xmp.Left - Sensibilta) And _
           X < (xmp.Left + xmp.Width + Sensibilta) And _
           Y > (xmp.Top + xmp.Height - Sensibilta) And _
           Y < (xmp.Top + xmp.Height + Sensibilta) Then
           X = frmListone.Left + frmListone.Width
           Y = frmListone.Top
           If X > (xmp.Left - Sensibilta) And _
              X < (xmp.Left + xmp.Width + Sensibilta) And _
              Y > (xmp.Top + xmp.Height - Sensibilta) And _
              Y < (xmp.Top + xmp.Height + Sensibilta) Then
                frmListone.Top = xmp.Top + xmp.Height
                frmListone.Left = xmp.Left
                AgganciatoFlag = True
            End If
        End If
    Else
        frmListone.Top = xmp.Top + xmp.Height
        frmListone.Left = xmp.Left
        AgganciatoFlag = True
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.AgganciaSub" & vbCrLf & Err.Description)
        End
    End If
End Sub



Public Sub ScriviPosizioneListaMp3(Optional ByVal Filenamez As String = ".\mp3.ls")
    
    Dim Filez As String
    Dim FileNumber As Integer
    Dim sStore As String
    Dim Pos As Integer
    
    Dim xByte() As Byte
    
On Error GoTo ErrH
        Filez = Filenamez & ".$$$"
        FileNumber = FreeFile
        ChDrive App.Path
        ChDir App.Path
        
        WriteINFO "Save Pos List..."
        
        
        ReDim xByte(FileLen(Filenamez))
        
        Open Filenamez For Binary Access Read As #FileNumber
            Get #FileNumber, , xByte
        Close #FileNumber
        
        sStore = Space$(FileLen(Filenamez))
        
        Call CopyMemory(ByVal sStore, xByte(0), FileLen(Filenamez))
        
        Pos = InStr(1, sStore, vbCrLf)
        
        sStore = Right$(sStore, Len(sStore) - Pos - 1)
        
        sStore = Left$(sStore, Len(sStore) - 2)
        
        Open Filez For Output As #FileNumber
            Print #FileNumber, IndiceGlobalissimo
            Print #FileNumber, sStore
        Close #FileNumber
        Call SetINI(cfgFile, "POS.INFO", "LISTFILE", Filenamez)

        Call FileCopy(Filez, Filenamez)
                
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Funzionilobali.ScriviPosizionListaMp3" & vbCrLf & Err.Description)
        Close #FileNumber
        Call SetINI(cfgFile, "POS.INFO", "LISTFILE", App.Path + "\mp3.ls")
        Err.Clear
    End If
    
    If FileExists(Filez) Then
        Call Kill(Filez)
    End If
    
End Sub


Public Sub ScriviListaMp3(Optional ByVal Filenamez As String = ".\mp3.ls")
    Dim I As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim sStore As String
    Dim MaxFileNumber As Integer
    Dim sFile As String
    Dim sItem As String
    
On Error GoTo ErrH
        Filez = Filenamez & ".$$$"
        FileNumber = FreeFile
        ChDrive App.Path
        ChDir App.Path
        
        WriteINFO "Save List..."
        
        Open Filez For Output As #FileNumber
            Print #FileNumber, IndiceGlobalissimo
            MaxFileNumber = NFile
            If ExistTmpFile Then
                For I = 0 To MaxFileNumber

                    Call Pause

                    sFile = Trim$(GetFile(I))
                    sItem = Trim$(frmListone.ListaMp3.GetItem(I))

                    If sFile <> "" And sItem <> "" Then
                        sStore = sStore & sFile & ";" & GetNormFromDisk(I) _
                            & ";" & GetINI(cfgFile, "XMMP3", "XNORMLEVEL", 3, 0, 100) & vbCrLf & sItem

                        If I <> MaxFileNumber Then
                            sStore = sStore & vbCrLf
                        End If

                    End If
                Next
                Print #FileNumber, sStore
            End If
        Close #FileNumber
        Call SetINI(cfgFile, "POS.INFO", "LISTFILE", Filenamez)
        Call SetINI(cfgFile, "POS.INFO", "LSDATE", FileDateTime(Filenamez))

        Call FileCopy(Filez, Filenamez)
        
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Funzionilobali.ScriviListaMp3" & vbCrLf & Err.Description)
        Close #FileNumber
        Call SetINI(cfgFile, "POS.INFO", "LISTFILE", App.Path + "\mp3.ls")
        Err.Clear
    End If
    
    If FileExists(Filez) Then
        Call Kill(Filez)
    End If
End Sub

Private Sub ControllaIndex(ByRef Indice As Integer)
On Error Resume Next
    If GetFile(Indice) = "" Then
        Indice = -1
        Err.Clear
    End If
End Sub

Public Sub LeggiListaMp3(Optional ByVal Filenamez As String = ".\mp3.ls")
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim LstIndex As Integer
    Dim strListaMp3 As String
    Dim Mp3FileTmp As String
    Dim tmp() As String
    Dim NormLevel As String
On Error GoTo ErrH

        WriteINFO "Read List..."
               
        Filez = Filenamez
        FileNumber = FreeFile
        NormLevel = GetINI(cfgFile, "XMMP3", "XNORMLEVEL", 3, 0, 100)
        If Dir(Filez) <> "" Then
            Open Filez For Input As #FileNumber
                If Not EOF(FileNumber) Then
                    Input #FileNumber, LstIndex
                    While Not EOF(FileNumber)
                                                
                        Call Pause
                                            
                        Line Input #FileNumber, Mp3FileTmp
                        Line Input #FileNumber, strListaMp3
                        
                        'Per mantenere la comp' / To maintain compatibility
                        Mp3FileTmp = Replace$(Mp3FileTmp, Chr(34), "")
                        strListaMp3 = Replace$(strListaMp3, Chr(34), "")
                        '''''''''''''''''''''''
                        
                        tmp = Split(Mp3FileTmp, ";")
                        If isVector(tmp) Then
                            Mp3FileTmp = tmp(0)
                            If UBound(tmp) > 1 Then
                                If tmp(2) = NormLevel Then
                                    Call SetNormToDisk(I, CDbl(tmp(1)))
                                Else
                                    Call SetNormToDisk(I)
                                End If
                            End If
                        End If
        
                        Call SetFile(I, Mp3FileTmp)
                        If Trim$(Mp3FileTmp) <> "" And _
                           Trim$(strListaMp3) <> "" Then
                            frmListone.ListaMp3.AddItem (strListaMp3)
                            I = I + 1
                        End If
                    Wend
                    If LstIndex <> -1 Then
                        Call ControllaIndex(LstIndex)
                    End If
                    If LstIndex = -1 Then
                        frmListone.ListaMp3.ListIndex = LstIndex
                        IndiceGlobalissimo = LstIndex
                    Else
                        If LstIndex <= NFile And _
                           LstIndex <> 0 And _
                           LstIndex >= -1 Then
                            frmListone.ListaMp3.ListIndex = LstIndex
                            IndiceGlobalissimo = LstIndex
                        End If
                    End If
                End If
            Close #FileNumber
        Else
            IndiceGlobalissimo = -1
        End If
        Call SetINI(cfgFile, "POS.INFO", "LISTFILE", Filenamez)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.LeggiListaMp3" & vbCrLf & Err.Description)
        Close #FileNumber
    End If

End Sub

Public Sub StopAll()
    Call mStopStream
    Call AzzeraPosizione
    WriteINFO "Stop All"
End Sub

Public Sub LeggiCFG()
    Dim Vec() As String
    Dim I As Byte
    Dim Esiste As String
    Dim Filez As String
On Error GoTo ErrH
    FFTSAMPLE = GetINI(cfgFile, "VISUALIZATION", "FFTSAMPLE", "256")
    If FFTSAMPLE > 1024 Or FFTSAMPLE < 256 Then
        FFTSAMPLE = 256
    End If
    
    Gap = GetINI(cfgFile, "VISUALIZATION", "GAP", "2", 1, 50)
    Barwidth = GetINI(cfgFile, "VISUALIZATION", "BARWIDTH", "1", 1, xmp.gph.Width / Screen.TwipsPerPixelX)
    If Barwidth > Gap Then
        Barwidth = Gap
    End If
    PointFalls = GetINI(cfgFile, "VISUALIZATION", "POINTFALLS", "5")
    Tolleranza = GetINI(cfgFile, "VISUALIZATION", "TOLLERANCE", "1")
    FallsVel = GetINI(cfgFile, "VISUALIZATION", "FALLSVEL", "3")
    FallsVel = FallsVel / 10
    noFallsVel = GetINI(cfgFile, "VISUALIZATION", "NOFALLSVEL", "5")
    noFallsVel = noFallsVel / 10
    
Exit Sub
ErrH:
    Call ScriviLOG(Err.Source & " - FunzioniGlobali.LeggiCFG", Err.Number, Err.Description)
End Sub

Public Sub Pause(Optional Durata As Double = 0.03, Optional DoEv As Boolean = False)
On Error GoTo ErrH
    Dim Tick As Single
    Tick = Timer
    While Timer < Tick + Durata
        Call SleepEx(1, 1)
        If DoEv Then
            DoEvents
        End If
    Wend
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.Pause" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub InitStreamEngine()
On Error GoTo ErrH
    Dim Flag As Boolean
    Dim SampleRate As Long
    Dim x8bit As Boolean
    Dim Mono As Boolean
    Dim Device As Integer
    Dim Buffer As Long
    Dim HINFOS As Long
    Dim HMAXSIZE As Long
    Dim PriorityProcess As Long
    Dim PriorityThread As Long
    Dim hProcess As Long
    Dim hThread As Long
        
        WriteINFO "Init Stream Engine..."
        
        hProcess = GetCurrentProcess()
        hThread = GetCurrentThread()

        SampleRate = GetINI(cfgFile, "xmMP3", "SAMPLERATE", "44100")
        x8bit = GetINI(cfgFile, "xmMP3", "8BIT", "FALSE")
        Mono = GetINI(cfgFile, "xmMP3", "MONO", "FALSE")
        Device = GetINI(cfgFile, "xmMP3", "DEVICE", "-1")
        Buffer = GetINI(cfgFile, "xmMP3", "BUFFER", "1")
        PriorityProcess = Format(GetINI(cfgFile, "xmMP3", "PROCESS", "&H8000"))
        PriorityThread = GetINI(cfgFile, "xmMP3", "THREAD", "0")
        Call mInitxmMP3(xmp.hWnd, SampleRate, x8bit, Mono, Device)
        Call mBufferLengthInSecond(Buffer)
        HMAXSIZE = GetINI(cfgFile, "PREFERENCE", "HMAXSIZE", "102400")
        HINFOS = GetINI(cfgFile, "PREFERENCE", "HINFOS", "128")
        Call mSetReadHaedA(HINFOS, HMAXSIZE)
        
        Flag = GetINI(cfgFile, "XMMP3", "XSOUND", "False")
        Call SetXSound(Flag)
        
        Flag = GetINI(cfgFile, "XMMP3", "XNORM", "False")
        Call SetXNorm(Flag)
        
        Select Case x8bit
                Case True
                    mAmpBit 8
                Case Else
                    mAmpBit 16
        End Select
        Select Case Mono
                Case True
                    mMode 1
                Case Else
                    mMode 2
        End Select
        mFreq SampleRate
        SetPriorityClass hProcess, PriorityProcess
        SetThreadPriority hThread, PriorityThread

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.InitStreamEngine" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub GestisciPosFrm(Carica As Boolean, Optional ByRef Minimizzato As Boolean)
    Dim I As Byte
    Dim sCFG As String
    Dim Esiste As String
    Dim Filez As String
    Dim FileNumber As Integer
    
On Error GoTo ErrH
    If Carica Then
        
        WriteINFO "Read Preference..."
    
        xmp.Left = -1000 - xmp.Width
        frmListone.Left = -1000 - xmp.Width
                
        'xmp
        sCFG = GetINI(cfgFile, "POS.INFO", "TOP", CStr((Screen.Height / 2) - (xmp.Height / 2)))
        If (CInt(sCFG) + 2010) > Screen.Height Then
            xmp.Top = Screen.Height - 2010
        ElseIf CInt(sCFG) < 0 Then
            xmp.Top = 0
        Else
            xmp.Top = sCFG
        End If
        sCFG = GetINI(cfgFile, "POS.INFO", "LEFT", CStr((Screen.Width / 2) - (xmp.Width / 2)))
        If (CInt(sCFG) + xmp.Width) > Screen.Width Then
            xmp.Left = Screen.Width - xmp.Width
        ElseIf CInt(sCFG) < 0 Then
            xmp.Left = 0
        Else
            xmp.Left = sCFG
        End If
        'frmListone
        sCFG = GetINI(cfgFile, "POS.INFO", "LSTTOP", CStr(xmp.Top + xmp.Height))
        If (CInt(sCFG) + 2010) > Screen.Height Then
            frmListone.Top = Screen.Height - 2010
        ElseIf CInt(sCFG) < 0 Then
            frmListone.Top = 0
        Else
            frmListone.Top = sCFG
        End If
        sCFG = GetINI(cfgFile, "POS.INFO", "LSTLEFT", CStr(xmp.Left))
        If (CInt(sCFG) + frmListone.Width) > Screen.Width Then
            frmListone.Left = Screen.Width - frmListone.Width
        ElseIf CInt(sCFG) < 0 Then
            frmListone.Left = 0
        Else
            frmListone.Left = sCFG
        End If
        
        'Loop
        sCFG = GetINI(cfgFile, "POS.INFO", "LOOP", "False")
        Call SetLoop(CBool(sCFG))
        
        'Rnd
        sCFG = GetINI(cfgFile, "POS.INFO", "RND", "False")
        Call SetRnd(CBool(sCFG))
        
        xmp.Show
        frmListone.Show
        
        sCFG = GetINI(cfgFile, "POS.INFO", "SHOWLIST", "True")
        Call SetHideListone(CBool(sCFG))
        
        sCFG = GetINI(cfgFile, "PREFERENCE", "SERVER", "False")
        If CBool(sCFG) Then
            Minimizzato = True
        Else
            sCFG = GetINI(cfgFile, "POS.INFO", "NOMINIMIZED", "True")
            Minimizzato = Not CBool(sCFG)
        End If
        
        sCFG = GetINI(cfgFile, "POS.INFO", "ONTOP", "False")
        frmMenu.InPrimoPianoFlag = CBool(sCFG)
        If frmMenu.InPrimoPianoFlag Then
            frmMenu.InPrimoPiano_Click
        End If
        'Forse da mettere insieme / Maybe these should be grouped together
        sCFG = GetINI(cfgFile, "POS.INFO", "COSC", "True")
        SetCardioOSCVisibile CBool(sCFG)
        sCFG = GetINI(cfgFile, "POS.INFO", "ANALYZER", "False")
        xmp.analyzer.Visible = CBool(sCFG)
        sCFG = GetINI(cfgFile, "POS.INFO", "SPECTRUMMODE", "1")
        GlobalSpectrumeMode = CInt(sCFG)
        sCFG = GetINI(cfgFile, "POS.INFO", "CPULESS", "False")
        xmp.ImgLogo.Visible = CBool(sCFG)
        '''
        'Provvisorio Qui / Temporary here
        sCFG = GetINI(cfgFile, "POS.INFO", "VOLUME", "25")
        If CInt(sCFG) <> xmp.xmsVol.xMin Then
            xmp.xmsVol.xValue = xmp.xmsVol.xMax - CInt(sCFG)
        Else
            xmp.xmsVol.xValue = xmp.xmsVol.xMax - (xmp.xmsVol.xMax \ 4)
        End If
        'Mp3 di cui e' stata salvata la posizione / Mp3 whose position was saved
        sCFG = GetINI(cfgFile, "POS.INFO", "LASTMP3", "")
        gPosForSave.Mp3File = sCFG
        'Posizione Mp3 / Mp3 position
        sCFG = GetINI(cfgFile, "POS.INFO", "POS", "0")
        gPosForSave.SamplePos = sCFG
                
        If xmp.ImgLogo.Visible = False And _
           xmp.analyzer.Visible = False And _
           GetCardioOSCVisibile = False Then
           xmp.ImgLogo.Visible = True
        End If
        
        frmEQ.LoadEQ

        'Show Task
        sCFG = GetINI(cfgFile, "VISUALIZATION", "SHOWTASK", "False")
        Call ShowTaskBarIcon(CBool(sCFG))
    Else
    
        WriteINFO "Write Preference..."
    
        Call SetINI(cfgFile, "POS.INFO", "NOMINIMIZED", CStr(xmp.Visible))
        If xmp.WindowState = vbMinimized Then
           xmp.WindowState = vbNormal
           xmp.Show
           frmListone.Show
        End If
        Call SetINI(cfgFile, "POS.INFO", "TOP", CStr(xmp.Top))
        Call SetINI(cfgFile, "POS.INFO", "LEFT", CStr(xmp.Left))
        Call SetINI(cfgFile, "POS.INFO", "LSTTOP", CStr(frmListone.Top))
        Call SetINI(cfgFile, "POS.INFO", "LSTLEFT", CStr(frmListone.Left))
        Call SetINI(cfgFile, "POS.INFO", "ONTOP", CStr(frmMenu.InPrimoPianoFlag))
        Call SetINI(cfgFile, "POS.INFO", "SHOWLIST", CStr(Not frmListone.ScancellatoDaXmp))
        Call SetINI(cfgFile, "POS.INFO", "COSC", CStr(GetCardioOSCVisibile))
        Call SetINI(cfgFile, "POS.INFO", "ANALYZER", CStr(xmp.analyzer.Visible))
        Call SetINI(cfgFile, "POS.INFO", "SPECTRUMMODE", CStr(GlobalSpectrumeMode))
        Call SetINI(cfgFile, "POS.INFO", "CPULESS", CStr(xmp.ImgLogo.Visible))
        '''
        'Provvisorio Qui / Temporary here
        Call SetINI(cfgFile, "POS.INFO", "VOLUME", CStr(xmp.xmsVol.xMax - xmp.xmsVol.xValue))
        If GetFile(IndiceGlobalissimo) <> "" And GetINI(cfgFile, "PREFERENCE", "SSTREAMPOS", 0) _
            And mStreamPos <> mStreamLen Then
                gPosForSave.Mp3File = GetFile(IndiceGlobalissimo)
        Else
                gPosForSave.Mp3File = ""
                gPosForSave.SamplePos = 0
        End If
        Call SetINI(cfgFile, "POS.INFO", "LASTMP3", gPosForSave.Mp3File)
        Call SetINI(cfgFile, "POS.INFO", "POS", CStr(gPosForSave.SamplePos))
        'Loop
        Call SetINI(cfgFile, "POS.INFO", "LOOP", xmp.g_Ripeti)
        'Rnd
        Call SetINI(cfgFile, "POS.INFO", "RND", xmp.Acaso)
        'Mute
'        Call SetIni(cfgFile, "POS.INFO", "MUTE", xmp.Mute)
        'Show Task
        Call SetINI(cfgFile, "VISUALIZATION", "SHOWTASK", CStr(frmShowTask.ShowMe))
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
               ".FunzioniGlobali.GestisciPosFrm" & _
               vbCrLf & Err.Description)
        Err.Clear
        Call DeLoad(False, False)
    End If
End Sub

Public Sub DeLoad(Optional ScrivereListaMp3 As Boolean, _
                  Optional SaveData As Boolean = True, _
                  Optional Brutal As Boolean = False)
On Error GoTo ErrH
   Dim I As Integer
   Dim J As Integer
   Dim ListFile As String
    
    If Not gb_InShutdown Then
        gb_InShutdown = True
        If Not Brutal Then
            Call xmpPensa(True)
            '''''''''''''''''''''''''''
            WriteINFO "Stop Timer..."
            xmp.UnSec.Enabled = False
            frmListone.frmListoneTimer.Enabled = False
            frmVisualizzazioni.Timer.Enabled = False
            frmVisualizzazioniAttivo = False
            '''''''''''''''''''''''''''
            
            DeICONtray
                            
            mTerminatexmMP3
        
            If SaveData Then
                If ScrivereListaMp3 Then
                    ListFile = GetINI(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
                    Call ScriviListaMp3(ListFile)
                End If
                Call GestisciPosFrm(False)
            End If
            
            '''''''''''''''''''''''''''
            WriteINFO "Remove Exception Filter..."
            SetUnhandledExceptionFilter 0
            '''''''''''''''''''''''''''
                   
            '''''''''''''''''''''''''''
            WriteINFO "UnLoad Forms..."
            For I = Forms.count - 1 To 0 Step -1
                Unload Forms(I)
            Next I
            '''''''''''''''''''''''''''
        
        End If
       
        '''
        WriteINFO "LogOUT from MH..."
        Call modMicroHandler.SendCommand(LOGOUT)
        '''
        
        Call RemoveFont
        
        Call KillTmpDat(False)
        
        WriteINFO "All System Nominal..."
        Call Pause
        WriteINFO "Shutting Down"
        Call Pause
        
        If Command$ <> "/debug" Then
            'Fully terminate the current process
            WriteINFO "Process Killing"
            Call TerminateProcess(GetCurrentProcess, ByVal 0&)
        Else
            End
        End If
    End If
       
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.CloseXmP" & vbCrLf & Err.Description)
        End
    End If
End Sub

Public Function isVector(Vettore As Variant, Optional Indice As Integer) As Boolean
On Error GoTo ErrH
    If UBound(Vettore) <> -1 Then
        Vettore(Indice) = Vettore(Indice)
    End If
    isVector = True
Exit Function
ErrH:
     Err.Clear
     isVector = False
End Function

Public Sub MinimizzaXMP()
On Error GoTo ErrH
    Minimizzato = True
    xmp.WindowState = vbMinimized
    frmShowTask.Hide
    xmp.Hide
    frmListone.Hide
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.MinimizzaXMP" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub AzzeraPosizione()
On Error GoTo ErrH
    mStreamPos 0, True
    xmp.Durata.xCaption = DurataStream(mStreamLenInSeconds)
    xmp.PicPosizione.Visible = False
    xmp.PicPosizione.Left = xmp.lnPosizione.X1
    xmp.PicPosizione.Visible = True
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.AzzeraPosizione" & vbCrLf & Err.Description)
    End If
End Sub

Public Function PosizioneSlide(ByVal X As Integer, _
                          ByVal lWidth As Long, _
                          ByVal ValMax As Long) As Long

On Error GoTo ErrH

        lWidth = ValMax \ lWidth
        PosizioneSlide = (lWidth * X) - Gap

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.fncPosizione" & vbCrLf & Err.Description)
    End If
End Function


Public Sub SettaIndicatoreModo(Optional Azzera As Boolean)
On Error GoTo ErrH
    
    '0 mono 1 stereo -1 Azzerato / 0 mono, 1 stereo, -1 reset
    If Azzera = True Then
        xmp.picModo.Picture = LoadResPicture("MODO0", vbResBitmap)
        Modo = TipoModo.Azzerato
    Else
        xmp.picModo.Picture = LoadResPicture("MODO0", vbResBitmap)
        Modo = TipoModo.Azzerato
        If frmMenu.mnuXSound.Checked Then
            Modo = TipoModo.XSound
            xmp.picModo.Picture = LoadResPicture("XSOUNDON", vbResBitmap)
            xmp.picModo.ToolTipText = "XSound"
        Else
            If InStr(1, mGetMP3Info.Mode, "STEREO", vbTextCompare) <> 0 Then
                Modo = TipoModo.Stereo
                xmp.picModo.Picture = LoadResPicture("STEREOON", vbResBitmap)
                xmp.picModo.ToolTipText = mGetMP3Info.Mode
            ElseIf mGetMP3Info.Mode <> "" Then
                Modo = TipoModo.Mono
                xmp.picModo.Picture = LoadResPicture("MONOON", vbResBitmap)
                xmp.picModo.ToolTipText = mGetMP3Info.Mode
            End If
        End If
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SettaIndicatoreModo" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub PulisciTutto(Optional Domanda As Boolean = False)
Dim Res As Integer
On Error GoTo ErrH
            
            If Domanda Then
                Res = MessageBox("Do you want to clear?", MsgAsk + MsgYesNo)
            Else
                Res = vbYes
            End If
            If Res = vbYes Then
                Call StopAll
                Call KillTmpDat(True)
                IndiceGlobalissimo = -1
                xmp.xmDisplay(0).xCaption = "Nope"
                xmp.Durata.xCaption = "00:00"
                xmp.BitRate.xCaption = ""
                xmp.lblFreq.xCaption = ""
                frmListone.ListaMp3.Clear
                WriteINFO "Nope"
                Call InfoNmp3
                SettaIndicatoreModo True
                Pause
                Call SetINI(cfgFile, "POS.INFO", "LISTFILE", "mp3.ls")
            End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.PulisciTutto" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub ControllaStrToPrecIstance()
    Dim Filez As String
    Dim FileNumber As Integer
    Dim StrCommand As String
    Dim Store As String
    Dim ListFile As String
    Dim Ext As String
    
On Error GoTo ErrH
If Command$ <> "/debug" Then
        Filez = App.Path + "\commandtrs.tm$"
        FileNumber = FreeFile
        If Dir(Filez) <> "" Then
            Open Filez For Input Shared As #FileNumber
                While Not EOF(FileNumber)
                    Input #FileNumber, Store
                    StrCommand = StrCommand + Store + " "
                Wend
                If Err.Number = 0 Then
                    If Trim$(StrCommand) <> "" And _
                       InStr(1, StrCommand, "/") = 0 Then
                            Ext = LCase(Right$(Trim$(StrCommand), 3))
                            Select Case Ext
                                Case Is = ".ls"
                                    PulisciTutto
                                    StrCommand = Trim$(StrCommand) + " "
                                    ListFile = Left$(StrCommand, _
                                    InStr(StrCommand, " "))
                                    Call LeggiListaMp3(ListFile)
                                Case Is = "m3u"
                                    PulisciTutto
                                    StrCommand = Trim$(StrCommand) + " "
                                    ListFile = Left$(StrCommand, _
                                    InStr(StrCommand, " "))
                                    Call LeggiListaM3U(ListFile)
                                Case Else
                                    Call CommandMP3(StrCommand)
                            End Select
                    End If
                End If
            Close #FileNumber
            Kill Filez
        End If
ErrH:
    Err.Clear
End If
End Sub

Public Function ElaboraNomeFile(NomeFile As String) As String
Dim Posizione As Integer
Dim I, J As Integer
Dim CaratteriChiave() As Variant
On Error GoTo ErrH
    NomeFile = Trim$(NomeFile)
    CaratteriChiave = Array(" ", "-", ".", "(", ")")
    If NomeFile <> "" Then
        NomeFile = Replace$(NomeFile, "_", " ")
        NomeFile = UCase$(Left$(NomeFile, 1)) & LCase$(Right$(NomeFile, Len(NomeFile) - 1))
        For I = 0 To UBound(CaratteriChiave)
            Posizione = 1
            While Posizione <> 0 And Posizione <> Len(NomeFile)
                Posizione = InStr(Posizione + 1, NomeFile, CaratteriChiave(I), vbTextCompare)
                If Posizione <> Len(NomeFile) Then
                    NomeFile = Left$(NomeFile, Posizione) & _
                               UCase$(Mid$(NomeFile, Posizione + 1, 1)) & _
                               Right$(NomeFile, Len(NomeFile) - Posizione - 1)
                End If
            Wend
        Next
    End If
    'Gestione numeri romani maggiori di una lettera fino a X (provvisoria) / Handling of roman numerals longer than one letter, up to X (temporary)
    'Oltre ai numeri si puo estenderlo anche alle sigle credo ??? / Beyond the numbers this could probably be extended to abbreviations too, I think???
    Dim NumeroDeRoma() As Variant
    NumeroDeRoma = Array("Iii", "Ii", "Iv", "Vi", "Viiii", "Viii", "Vii")
    For J = 0 To UBound(NumeroDeRoma)
        For I = 0 To UBound(CaratteriChiave)
            NomeFile = Replace$(NomeFile, NumeroDeRoma(J) & CaratteriChiave(I), UCase(NumeroDeRoma(J)) & CaratteriChiave(I))
            NomeFile = Replace$(NomeFile, CaratteriChiave(I) & NumeroDeRoma(J), UCase(CaratteriChiave(I) & NumeroDeRoma(J)))
        Next
    Next
    '''
    ElaboraNomeFile = NomeFile
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ElaboraNomeFile" & vbCrLf & Err.Description)
    End If
End Function

Public Sub IlluminaPulsante(CommandImg As PictureBox, _
                            Optional LockButton As Boolean = True, _
                            Optional Colore As Long = vbGreen)
On Error GoTo ErrH
    If LockButton = False Then
         CommandImg.AutoRedraw = True
         CommandImg.Cls
         CommandImg.AutoRedraw = False
     Else
         Dim sX As Integer
         Dim sY As Integer
         Dim sXTwips As Integer
         Dim sYTwips As Integer
         sXTwips = Screen.TwipsPerPixelX
         sYTwips = Screen.TwipsPerPixelY
         sX = CommandImg.Width - (sXTwips * 3)
         sY = CommandImg.Height - (sYTwips * 3)

         CommandImg.AutoRedraw = True
         CommandImg.Line (0, 0)-(0, sY), Colore
         CommandImg.Line -(sX, sY), Colore
         CommandImg.Line -(sX, 0), Colore
         CommandImg.Line -(0, 0), Colore
         CommandImg.PSet (sXTwips, sYTwips), Colore
         CommandImg.PSet (sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sY - sYTwips), Colore
         CommandImg.PSet (sX - sXTwips, sYTwips), Colore
         CommandImg.AutoRedraw = False
     End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.IlluminaPulsante" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub SettaNomeMp3()
On Error GoTo ErrH
    
    If mGetMP3Info.SongName = "" Then
        xmp.xmDisplay(0).xCaption = ElaboraNomeFile(RetFileName(Dir(GetFile(IndiceGlobalissimo))))
    Else
        xmp.xmDisplay(0).xCaption = ElaboraNomeFile(mGetMP3Info.SongName)
    End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SettaNomeMp3" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub PlayDone()

Dim NuFile As Integer

On Error GoTo ErrH
        
    NuFile = NFile
    
    If frmMenu.mnuXNorm.Checked Then
        Call ElabNormalize
    End If

    With xmp
        gPosForSave.SamplePos = 0
        xmp.Durata.xCaption = DurataStream(0)
        
        Call Pause
    
        If mStreamLen <> -1 Then
            Call StopAll
            If ExistTmpFile Then
                If IndiceGlobalissimo <= NuFile Then
                    
                    If .Acaso Then
                        If NuFile >= 1 Then
                            Call PlayRnd
                        Else
                            Call PlayStream(0)
                        End If
                    ElseIf .g_Ripeti And (IndiceGlobalissimo = NuFile) Then
                        Call PlayStream(0)
                    ElseIf IndiceGlobalissimo < NuFile Then
                        Call PlayStream(IndiceGlobalissimo + 1)
                    Else
                        .g_PlayDone = False
                        WriteINFO "Play Done."
                    End If
                    
                End If
            End If
        End If
        
    End With
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.PlayDone" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub xmpPensa(Optional Flag As Boolean = False)
    If Flag Then
        xmp.MousePointer = vbHourglass
        DoEvents
        frmListone.MousePointer = vbHourglass
        DoEvents
    Else
        xmp.MousePointer = vbDefault
        frmListone.MousePointer = vbDefault
    End If
    'Call DisableFRM(Flag)
End Sub

Public Sub DisableFRM(Optional Flag As Boolean = False)
    Dim I As Integer
    
    For I = Forms.count - 1 To 0 Step -1
        Forms(I).Enabled = Not Flag
    Next I

End Sub


Public Sub frmSetFocus(frm As Form)
    If frm.Visible And frm.Enabled Then
        frm.SetFocus
    End If
End Sub

Public Function AddFont(Optional FontTag As String = "FONTFILE") As Boolean
    Dim StrPath As String
    
On Error GoTo ErrH:

    StrPath = GetINI(cfgFile, "VISUALIZATION", FontTag, "nope")
    If StrPath <> "nope" Then
        StrPath = App.Path + "\" + StrPath
        Call AddFontResource(StrPath)
        AddFont = True
    Else
        AddFont = False
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.AddFont" & vbCrLf & Err.Description)
    End If
End Function

Public Function RemoveFont(Optional FontTag As String = "FONTFILE") As Boolean
On Error GoTo ErrH:
    Dim StrPath As String
    
    WriteINFO "Remove Font Cache..."
    
    StrPath = GetINI(cfgFile, "VISUALIZATION", FontTag, "nope")
    If StrPath <> "nope" Then
        StrPath = App.Path + "\" + StrPath
        RemoveFontResource StrPath
        RemoveFont = True
    Else
        RemoveFont = False
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.RemoveFont" & vbCrLf & Err.Description)
    End If
End Function
Public Sub RoundForm(frm As Form)
    Dim Raggio As Integer
On Error GoTo ErrH:
    Raggio = GetINI(cfgFile, "VISUALIZATION", "ROUND", 15)
    Call SetRoundForm(frm.hWnd, 0, 0, (frm.Width / Screen.TwipsPerPixelY) + 1, (frm.Height / Screen.TwipsPerPixelX) + 1, Raggio, Raggio)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.RoundForm" & vbCrLf & Err.Description)
    End If
End Sub


Public Sub GestioneExtDrag(ByVal ObjDrag As DragEngine)
    Dim ListFile As String
On Error GoTo ErrH:
    ListFile = Replace(ObjDrag.FileName(0), Chr(0), "")
        Select Case LCase(Right$(ListFile, 3))
            Case Is = ".ls"
                PulisciTutto
                Call LeggiListaMp3(ListFile)
            Case Is = "m3u"
                PulisciTutto
                Call LeggiListaM3U(ListFile)
            Case Else
                Call ElaboraFileDrop(ObjDrag)
        End Select
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.GestioneExtDrag" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub ShowLedVol()
    Dim lVol As Long
On Error GoTo ErrH:
    xmp.xmDVol.Visible = True
    If IsObject(frmVolume) Then
        Load frmVolume
    End If
    lVol = mStreamVol
    DrawLED frmVolume.PercValLed, AddZeri(CStr(lVol), 3), frmVolume.LED, 1
    xmp.xmDVol.xCaption = AddZeri(CStr(lVol), 3)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ShowLedVol" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub WriteINFO(sInfo As String)
    
    If gb_InfoLastLine <= 1 Then
        frmListone.xmDInfo(gb_InfoLastLine).xCaption = sInfo
        gb_InfoLastLine = gb_InfoLastLine + 1
    Else
        frmListone.xmDInfo(0).xCaption = frmListone.xmDInfo(1).xCaption
        frmListone.xmDInfo(1).xCaption = sInfo
    End If
    DoEvents
End Sub


