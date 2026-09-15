Attribute VB_Name = "modFunzioniAccessorie"
Option Explicit

Public Sub ElabNormalize()

Dim XNTune As Double
Dim XNorm As Double
Dim CurXNorm As Double

On Error GoTo errH
    
    XNTune = CDbl(GetINI(cfgFile, "XMMP3", "XNORMTUNE", "1", "0", "50")) / 100
    
    CurXNorm = CDbl(GetNormFromDisk(IndiceGlobalissimo))
    XNorm = mXNormLevel()
    Select Case XNorm
        Case Is < 1
            If CurXNorm <> 1 Then
                XNorm = CurXNorm - XNTune
            End If
            Call SetNormToDisk(IndiceGlobalissimo, XNorm)
        Case Is = 1
            If CurXNorm >= 1 Then
                Call SetNormToDisk(IndiceGlobalissimo, _
                CurXNorm + XNTune)
            End If
    End Select
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniAccessorie.ElabNormalize" & vbCrLf & Err.Description)
    End If
End Sub



Public Sub PlayRnd()
Dim I As Integer, J As Integer
Dim rndOK As Boolean
Dim rndFile As String
Dim NuFile As Integer
Dim NuRndFile As Integer

On Error GoTo errH
    rndFile = App.Path & "\rnd.___"
    
    If Not FileExists(rndFile) Then
        Call SetFile(0, _
                     CStr(IndiceGlobalissimo), "RND", _
                     rndFile)
    End If
    
    NuRndFile = NFile(rndFile, "RND")
    NuFile = NFile
    
    Do
        Randomize
        IndiceGlobalissimo = Rnd * NuFile
        
        rndOK = True
        For I = 0 To NuRndFile
            If CInt(GetFile(I, "RND", rndFile)) = IndiceGlobalissimo Then
                WriteINFO "Randomizing."
                rndOK = False
                Exit For
            End If
        Next
    Loop Until (rndOK = True) Or (NuRndFile = NuFile)
    If NuRndFile = NuFile Then
        If FileExists(rndFile) Then
            Call Kill(rndFile)
        End If
        
        If xmp.g_Ripeti Then
            IndiceGlobalissimo = 0
            Call PlayStream(IndiceGlobalissimo)
        Else
            xmp.g_PlayDone = False
            WriteINFO "Play Done."
        End If
    Else
        Call SetFile(NuRndFile + 1, _
                     CStr(IndiceGlobalissimo), "RND", _
                     rndFile)
        Call PlayStream(IndiceGlobalissimo)
    End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniAccessorie.PlayRnd" & vbCrLf & Err.Description)
    End If
End Sub


Public Sub DrawOCenterLine(PicBox As PictureBox, lColor As Long)

On Error GoTo errH
    PicBox.Line (0, (PicBox.Height \ 2)) _
    -(PicBox.Width, (PicBox.Height \ 2)), lColor
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modFunzioniAccessorie.DrawOCenterLine" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub OpenMp3dlg()

On Error GoTo errH

Dim sOpen As SelectedFile
Dim LocalStr As String
Dim I As Integer
       
        FileDialog.sFilter = ""
        
        FileDialog.sDefFileExt = "All Supported file"
        FileDialog.sFilter = FileDialog.sFilter & "All Supported file ("
        For I = 0 To UBound(Estensioni)
            LocalStr = LocalStr & "*" & Estensioni(I) & ";"
        Next
        FileDialog.sFilter = FileDialog.sFilter & LocalStr & ")" + Chr$(0) & LocalStr
        FileDialog.sFilter = FileDialog.sFilter + Chr$(0)
        

        For I = 0 To UBound(Estensioni)
            FileDialog.sFilter = FileDialog.sFilter & _
                                Right$(Estensioni(I), 3) & " file (*" & Estensioni(I) & _
                                ")" & Chr$(0) & "*" & Estensioni(I) & Chr$(0)
        Next
                                
        FileDialog.sFilter = FileDialog.sFilter & "All Files (*.*)" & Chr$(0) & "*.*"
        FileDialog.sInitDir = GetINI(cfgFile, "POS.INFO", "LASTDIR", App.Path)
        FileDialog.sDlgTitle = "XmP"
        
        'FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
        FileDialog.Flags = OFS_FILE_OPEN_FLAGS
        
        sOpen = ShowOpen(xmp.hWnd, False)
        
        If isVector(sOpen.sFiles) Then
            Select Case LCase(Right$(sOpen.sFiles(0), 3))
                Case Is = ".ls"
                    Call PulisciTutto
                    Call LeggiListaMp3(sOpen.sLastDirectory + sOpen.sFiles(0))
                Case Is = "m3u"
                    Call PulisciTutto
                    Call LeggiListaM3U(sOpen.sLastDirectory + sOpen.sFiles(0))
                Case Else
                    Call ImpostaListaMp3(sOpen)
                    Call SetINI(cfgFile, "POS.INFO", "LASTDIR", sOpen.sLastDirectory)
            End Select
        End If
        
        EraseFileDialog
        
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modFunzioniAccessorie.OpenMp3dlg" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub SetXNorm(Flag As Boolean)
    frmMenu.mnuXNorm.Checked = Flag
    Call mXNorm(Flag)
    If Flag Then
        Call SetINI(cfgFile, "XMMP3", "XNORM", "True")
    Else
        Call SetINI(cfgFile, "XMMP3", "XNORM", "False")
        If ExistTmpFile Then
            Call SetNormToDisk(IndiceGlobalissimo)
        End If
    End If
End Sub

Public Sub SetXSound(Flag As Boolean)
    frmMenu.mnuXSound.Checked = Flag
    Call mXSrnd(Flag)
    If Flag Then
        Call SetINI(cfgFile, "XMMP3", "XSOUND", "True")
    Else
        Call SetINI(cfgFile, "XMMP3", "XSOUND", "False")
    End If
    Call SettaIndicatoreModo
End Sub


Public Sub SetLoop(Flag As Boolean)
    xmp.g_Ripeti = Flag
    frmMenu.mnuLoop.Checked = Flag
    frmListone.Ripeti.Visible = Flag
End Sub

Public Sub SetRnd(Flag As Boolean)
    xmp.Acaso = Flag
    frmMenu.mnuRnd.Checked = Flag
    If (Not Flag) And FileExists(App.Path & "\rnd.___") Then
        Call Kill(App.Path & "\rnd.___")
    End If
End Sub

Public Sub SetHideListone(Flag As Boolean)
    frmListone.Visible = Flag
    frmListone.ScancellatoDaXmp = Not Flag
    frmMenu.mnuHideList.Checked = Not Flag
End Sub
Public Sub HandCursorOver(Obj As Object)
    If Obj.MousePointer <> vbCustom Then
        Obj.MousePointer = vbCustom
        Obj.MouseIcon = LoadResPicture("INETHAND", vbResCursor)
    End If
End Sub

Public Sub ShowTaskBarIcon(ByVal Flag As Boolean)
    If Flag Then
        frmShowTask.Top = -1000 * frmShowTask.Height
        frmShowTask.Left = -1000 * frmShowTask.Width
        frmShowTask.Show
        frmShowTask.ShowMe = True
    Else
        frmShowTask.Hide
        frmShowTask.ShowMe = False
    End If
End Sub

Public Sub About()
    Call SempreInPrimoPiano(frmAbout.hWnd, frmAbout.Width, frmAbout.Height, _
                            frmAbout.Left, frmAbout.Top, frmMenu.InPrimoPiano)
    Call frmAbout.Show(vbModal, xmp)
End Sub

Public Sub PassaStrToPrevIstance(StrCommand As String)
    Dim Filez As String
    Dim FileNumber As Integer
On Error GoTo errH
    If Trim$(StrCommand) <> "" And _
        InStr(1, StrCommand, "/") = 0 Then
        
        Filez = App.Path + "\commandtrs.tm$"
        StrCommand = Replace$(StrCommand, Chr$(34), "")
        FileNumber = FreeFile
        Open Filez For Append Shared As #FileNumber
            Write #FileNumber, StrCommand
        Close #FileNumber
    End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "Funzionilobali.PassaStrToPrevIstance" & vbCrLf & Err.Description)
        Close #FileNumber
    End If
End Sub

Public Sub ScorriLabelNomeMp3(MaxChar As Integer, _
                       Optional Modo As Byte = 0, _
                       Optional LungPausa As Double = 0.08)

    Call frmVisualizzazioni.ScorriLabelNomeMp3(MaxChar, Modo, LungPausa)

End Sub

Public Sub ElaboraFileDrop(DragMP3 As DragEngine)
On Error GoTo errH
    Dim I As Integer
    Dim Indice As Byte
    Dim iFile As Integer
    Dim xFileName As String
    Dim EstensioneAbilitata As Integer
    Dim Primo As Integer


    'xmp.Visible = False
    'Listone.Visible = False
    EstensioneAbilitata = 0
    
    If ExistTmpFile Then
        iFile = NFile
    Else
        iFile = -1
    End If
    With DragMP3
        For I = 0 To .FileCount - 1
            xFileName = Replace$(.FileName(I), vbNullChar, "")
            For Indice = 0 To UBound(Estensioni)
                EstensioneAbilitata = InStr(1, UCase$(xFileName), Estensioni(Indice))
                If EstensioneAbilitata > 0 Then
                    frmListone.ListaMp3.AddItem RetFileName(xFileName)
                    'ReDim Preserve Directory(iFile + 1)
                    Call SetFile((iFile + 1), xFileName)
                    iFile = iFile + 1
                End If
            Next
        Next
    End With
    Call InfoNmp3
    DragMP3.ClearFileNames
    frmListone.ListaMp3.ListIndex = frmListone.ListaMp3.ListCount - 1
    
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ElaboraFileDrop" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub CommandMP3(Comando As String)
On Error GoTo errH
    Dim Posizione As Integer
    Dim StorePosizione As Integer
    Dim I As Integer
    Dim Indice As Byte
    Dim iFile As Integer
    Dim xFileName As String
    Dim TmpStr As String
    Dim EstensioneAbilitata As Integer
    Dim Uno As Boolean
    Dim OLDIndex As Integer
    Dim ExtLen As Byte
     
    
    Uno = False
    EstensioneAbilitata = 0
    If ExistTmpFile Then
        iFile = NFile + 1
    Else
        iFile = 0
    End If
    OLDIndex = iFile
    Comando = Replace$(Trim$(Comando), """", "")
    StorePosizione = 0
    I = 0
    Do
        Do
            Do
                Posizione = StorePosizione
                StorePosizione = InStr(StorePosizione + 1, UCase$(Comando), Estensioni(I))
                If Posizione + 3 <> Len(Comando) Then
                    If Mid$(Comando, Posizione + 4, 1) = " " Then
                        Exit Do
                    End If
                End If
            Loop Until StorePosizione = 0
            I = I + 1
        Loop Until Posizione <> 0 Or I > UBound(Estensioni)
        StorePosizione = 0
        xFileName = Left$(Comando, Posizione + 3)
        Comando = Trim$(Right$(Comando, Len(Comando) - (Posizione + 3)))

        I = 0
        For Indice = 0 To UBound(Estensioni)
            EstensioneAbilitata = InStr(1, UCase$(xFileName), Estensioni(Indice))
            If EstensioneAbilitata > 0 Then
                TmpStr = Dir(xFileName)
                ExtLen = Abs(Len(TmpStr) - Len(Estensioni(Indice)))
                frmListone.ListaMp3.AddItem Left$(TmpStr, ExtLen)
                'RetFileName(xFileName)
                'ReDim Preserve Directory(iFile)
                Call SetFile(iFile, xFileName)
                iFile = iFile + 1
                If iFile - OLDIndex = 1 Then
                    Uno = True
                Else
                    Uno = False
                End If
            End If
        Next
    Loop Until Trim$(Comando) = ""
    Call InfoNmp3
    If Uno Then
        frmListone.ListaMp3.ListIndex = frmListone.ListaMp3.ListCount - 1
        Call PlayStream(frmListone.ListaMp3.ListCount - 1)
    End If
    
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "CommandMP3" & vbCrLf & Err.Description)
    End If
End Sub

Public Function AddZeri(Valore As String, NumeroZeri As Byte) As String
On Error GoTo errH
    Dim I As Integer
        For I = 0 To NumeroZeri - Len(Valore) - 1
            Valore = "0" + Valore
        Next
    AddZeri = Valore
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.AddZeri" & vbCrLf & Err.Description)
    End If
End Function

Public Sub SpostaItemList(IndiceCorrente As Integer, Vettore As xmListBox, SuGiu As SuGiu)
On Error GoTo errH
    Dim SalvaItem As String
        If SuGiu = Giu Then
            If IndiceCorrente > 0 Then
                SalvaItem = Vettore.GetItem(IndiceCorrente)
                Vettore.ModifyItem IndiceCorrente, Vettore.GetItem(IndiceCorrente + SuGiu)
                Vettore.ModifyItem IndiceCorrente + SuGiu, SalvaItem
                Vettore.ListIndex = Vettore.ListIndex + SuGiu
            End If
        ElseIf SuGiu = Su Then
            If IndiceCorrente < Vettore.ListCount - 1 Then
                SalvaItem = Vettore.GetItem(IndiceCorrente)
                Vettore.ModifyItem IndiceCorrente, Vettore.GetItem(IndiceCorrente + SuGiu)
                Vettore.ModifyItem IndiceCorrente + SuGiu, SalvaItem
                Vettore.ListIndex = Vettore.ListIndex + SuGiu
            End If
        End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SpostaItem" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub SpostaItem(IndiceCorrente As Integer, SuGiu As SuGiu)
On Error GoTo errH
    Dim SalvaItem As String
        If ExistTmpFile Then
            If SuGiu = Giu Then
                If IndiceCorrente > 0 Then
                    SalvaItem = GetFile(IndiceCorrente)
                    Call SetFile(IndiceCorrente, GetFile(IndiceCorrente + SuGiu))
                    Call SetFile(IndiceCorrente + SuGiu, SalvaItem)
                End If
            ElseIf SuGiu = Su Then
                If IndiceCorrente <= NFile - 1 Then
                    SalvaItem = GetFile(IndiceCorrente)
                    Call SetFile(IndiceCorrente, GetFile(IndiceCorrente + SuGiu))
                    Call SetFile(IndiceCorrente + SuGiu, SalvaItem)
                End If
            End If
            Select Case IndiceGlobalissimo
                Case Is = IndiceCorrente
                    IndiceGlobalissimo = IndiceGlobalissimo + SuGiu
                Case Is = IndiceCorrente + SuGiu
                    IndiceGlobalissimo = IndiceCorrente
            End Select
        End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.SpostaItem" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub InfoBoxShow(Optional NomeFile As String)
    Dim lGetMP3Info As modInfoMp3.Mp3Info
On Error GoTo errH
    Dim STR As String
        If Trim(NomeFile) = "" And IndiceGlobalissimo <> -1 Then
            NomeFile = GetFile(IndiceGlobalissimo)
        End If
        If frmListone.ListaMp3.ListIndex <> -1 Then
            lGetMP3Info = ReadMP3(NomeFile, True, True, gb_Data)
        End If
        Select Case Right$(UCase(NomeFile), 4)
            Case Is = ".WAV"
                frmInfo.Height = 11 * 200
                frmInfo.cmdChiudi.Top = frmInfo.Height - frmInfo.cmdChiudi.Height - 480

                
                frmInfo.lstInfo.Print "Wave File"
                frmInfo.lstInfo.Print "Mode      : " + lGetMP3Info.Mode
                frmInfo.lstInfo.Print "Frequency : " + CStr(lGetMP3Info.Frequency) + "Hz"
                frmInfo.lstInfo.Print "Bit Rate  : " + CStr(lGetMP3Info.BitRate) + "Kbit/s"
                frmInfo.lstInfo.Print "Duration  : " & DurataStream(CInt(lGetMP3Info.Duration))
                    

            Case Else
                frmInfo.Height = 28 * 200
                frmInfo.cmdChiudi.Top = frmInfo.Height - frmInfo.cmdChiudi.Height - 480
                
                frmInfo.lstInfo.FontName = "Courier"
                frmInfo.lstInfo.ForeColor = vbGreen
                
                frmInfo.lstInfo.Print ("Mpeg" & lGetMP3Info.MpegVersion & " Layer " & lGetMP3Info.MpegLayer)
                frmInfo.lstInfo.Print ("Mode      : " + lGetMP3Info.Mode)
                frmInfo.lstInfo.Print ("Frequency : " + CStr(lGetMP3Info.Frequency) + "Hz")
                frmInfo.lstInfo.Print ("Bit Rate  : " + CStr(lGetMP3Info.BitRate) + "Kbit/s")
                frmInfo.lstInfo.Print ("VBR       : " & lGetMP3Info.VBR)
                frmInfo.lstInfo.Print ("Emphasis  : " & lGetMP3Info.Emphasis)
                frmInfo.lstInfo.Print ("CRC       : " & lGetMP3Info.CRC)
                frmInfo.lstInfo.Print ("Duration  : " & DurataStream(CInt(lGetMP3Info.Duration)))
                frmInfo.lstInfo.Print ("Frames    : " & lGetMP3Info.Frames)
                frmInfo.lstInfo.Print ("HasTag    : " & lGetMP3Info.HasTag)
                frmInfo.lstInfo.Print ("PrivateBit: " & lGetMP3Info.PrivateBit)
                frmInfo.lstInfo.Print ("Track     : " & lGetMP3Info.track)
                frmInfo.lstInfo.Print ("Padding   : " & lGetMP3Info.Padding)
                
                frmInfo.lstInfo.Print ("")
                frmInfo.lstInfo.Print ("Comment   : " & lGetMP3Info.Comment)
                frmInfo.lstInfo.Print ("Artist    : " & lGetMP3Info.Artist)
                frmInfo.lstInfo.Print ("Name      : " & lGetMP3Info.SongName)
                frmInfo.lstInfo.Print ("Album     : " & lGetMP3Info.Album)
                frmInfo.lstInfo.Print ("Original  : " & lGetMP3Info.Original)
                frmInfo.lstInfo.Print ("Year      : " & lGetMP3Info.Year)
                frmInfo.lstInfo.Print ("Genre     : " & GenreText(lGetMP3Info.Genre))
                frmInfo.lstInfo.Print ("CopyRight : " & lGetMP3Info.CopyRight)
            End Select
            
            frmInfo.lstInfo.Print ("-------")
            frmInfo.lstInfo.Print ("")
            frmInfo.lstInfo.Print ("xmMP3 Engine = v" + mGetxmMP3Version)
            frmInfo.lstInfo.Print ("XmP GUI = v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " " + App.Comments)
            
            SempreInPrimoPiano frmInfo.hWnd, frmInfo.Width, frmInfo.Height, frmInfo.Left, frmInfo.Top, frmMenu.InPrimoPiano.Checked
            Call frmInfo.Show(vbModal, frmListone)

errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.InfoBoxShow" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub ImpostaListaMp3(sFile As SelectedFile)
On Error GoTo errH
    Dim lstCnt As Integer
    Dim FirstStream As Integer
    Dim I As Integer
    Dim Indice As Integer

        If sFile.bCanceled = False Then
            If ExistTmpFile Then
                lstCnt = NFile
            Else
                lstCnt = -1
            End If
            FirstStream = lstCnt
            sFile.sFiles = OrderDirVec(sFile.sFiles)
            'ReDim Preserve Directory(lstCnt + sFile.nFilesSelected)
            For I = lstCnt + 1 To lstCnt + sFile.nFilesSelected
                If Right$(sFile.sLastDirectory, 1) = "\" Then
                    Call SetFile(I, sFile.sLastDirectory + sFile.sFiles(Indice))
                Else
                    Call SetFile(I, sFile.sLastDirectory + "\" + sFile.sFiles(Indice))
                End If
                frmListone.ListaMp3.AddItem RetFileName(sFile.sFiles(Indice))
                Indice = Indice + 1
            Next
            frmListone.ListaMp3.Refresh
            frmListone.ListaMp3.ListIndex = FirstStream + 1
            If Not mStreamIsActive Then
                IndiceGlobalissimo = frmListone.ListaMp3.ListIndex
            End If
        End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ImpostaListaMp3" & vbCrLf & Err.Description)
    End If
End Sub

Private Function OrderDirVec(ByVal Vector As Variant) As String()
On Error GoTo errH:
    If UBound(Vector) > 2 Then
        Dim Index As Integer, Index2 As Integer
        Dim FirstItem As Integer, NumberOfItems As Integer
        Dim Value As String
        Dim Fine As Boolean
        
        FirstItem = LBound(Vector)
        NumberOfItems = UBound(Vector)
        
        Do
            
            Fine = True
            For Index = FirstItem + 1 To NumberOfItems
                If (UCase(Vector(Index)) < UCase(Vector(Index - 1))) Then
                    Value = Vector(Index - 1)
                    Vector(Index - 1) = Vector(Index)
                    Vector(Index) = Value
                    Fine = False
                End If
            Next
            
        Loop Until Fine
    End If
    OrderDirVec = Vector
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.OrderDirVec" & vbCrLf & Err.Description)
    End If
End Function

'Private Sub KillFile(FileNumber As Integer, Filez As String)
'On Error Resume Next
'        Close #FileNumber
'        Call Kill(Filez)
'End Sub

Public Sub KillTmpDat(Optional All As Boolean = False)
On Error GoTo errH

    WriteINFO "Kill Temp Files"
        
    If FileExists(App.Path & "\*.$$$") Then
        Call Kill(App.Path & "\*.$$$")
    End If
    If All And FileExists(App.Path & "\*.___") Then
        Call Kill(App.Path & "\*.___")
    End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.KillTmpDat" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub InfoNmp3()
On Error GoTo errH
    WriteINFO "(" + CStr(frmListone.ListaMp3.ListCount) + " Stream)"
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.infoNmp3" & vbCrLf & Err.Description)
    End If
End Sub


Public Sub ScriviLOG(ByVal where As String, ByVal Info As String, ErrorDes As String)

    Dim I As Integer
    Dim sLog As String
    Dim sEsito As String
    Dim Esiste As String
    Dim Filez As String
    Dim FileNumber As Integer
    
On Error Resume Next
        Filez = App.Path + "\Logs\"
        Esiste = Dir(Filez)
        If Esiste = "" Then
            MkDir (App.Path + "\Logs\")
        End If
On Error GoTo errH

        sLog = CStr(time()) & " - Dove: " & where & " " & Info & "  Dettagli: " & ErrorDes

        Filez = App.Path + "\Logs\" + CStr(Date) + ".log"
        Filez = Replace$(Filez, "/", ".")
        Esiste = Dir(Filez)
        If Esiste = "" Then
            FileNumber = FreeFile
            Open Filez For Output As #FileNumber
                Print #FileNumber, "Error Log File."
                Print #FileNumber, sLog
            Close #FileNumber
        Else
            FileNumber = FreeFile
            Open Filez For Append As #FileNumber
                Print #FileNumber, sLog
            Close #FileNumber
        End If

errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.ScriviLog" & vbCrLf & Err.Description)
    End If
End Sub

Public Function LeggiListaM3U(ByVal Filenamez As String)
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim Mp3FileTmp As String
    Dim aExt() As String
    
On Error GoTo errH
        ReDim aExt(0)
        Filez = Filenamez
        FileNumber = FreeFile
        If Dir(Filez) <> "" Then
            Open Filez For Input As #FileNumber
                If Not EOF(FileNumber) Then
                    While Not EOF(FileNumber)
                        Line Input #FileNumber, Mp3FileTmp
                        If Trim$(Mp3FileTmp) <> "" Then
                            If Trim$(Left$(Mp3FileTmp, 1)) <> "#" Then
                                Mp3FileTmp = Replace$(Mp3FileTmp, "/", "\")
                                Call SetFile(I, Mp3FileTmp)
                                If UBound(aExt) > 0 Then
                                    frmListone.ListaMp3.AddItem (DurataStream(CInt(aExt(0))) & " - " & aExt(1))
                                Else
                                    frmListone.ListaMp3.AddItem (RetFileName(Mp3FileTmp))
                                End If
                                I = I + 1
                            Else
                                If Trim$(Left$(Mp3FileTmp, 8)) = "#EXTINF:" Then
                                    aExt = Split(Right$(Mp3FileTmp, Len(Mp3FileTmp) - 8), ",")
                                End If
                            End If
                        End If
                    Wend
                End If
            Close #FileNumber
        End If

errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniGlobali.LeggiListaM3U" & vbCrLf & Err.Description)
        Close #FileNumber
    End If
End Function

Public Function RetFileName(ByVal FileN As String, Optional NoExt As Boolean = True) As String
Dim I As Byte
Dim Ext() As String
Dim Posizione As Integer
On Error GoTo errH
    
    RetFileName = FileN
    If FileN <> "" Then
        Ext = Estensioni
        Posizione = InStr(1, FileN, "\")
        While Posizione <> 0
            FileN = Right$(FileN, Len(FileN) - Posizione)
            Posizione = InStr(1, FileN, "\")
        Wend
        For I = 0 To UBound(Ext)
            If InStr(1, FileN, Ext(I), vbTextCompare) <> 0 Then
                If NoExt Then
                    Posizione = InStr(1, FileN, Ext(I), vbTextCompare)
                    RetFileName = Left$(FileN, Posizione - 1)
                Else
                    RetFileName = FileN
                End If
                Exit For
            End If
        Next
    End If
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniAccessorie.RetFileName" & vbCrLf & Err.Description)
    End If
End Function

Public Sub EnableXmP(Optional Abilita As Boolean = True)
    xmp.Enabled = Abilita
    frmListone.Enabled = Abilita
End Sub

Public Function MessageBox(Prompt As String, Optional Flags As MsgType = 0, Optional PromptCaption As String = "", Optional iForm As Form) As MsgResult
    MessageBox = frmMessage.dwMessageBox(Prompt, Flags, PromptCaption, iForm)
End Function

Public Function SearchList(ToSearch As Integer, lstList As xmListBox) As Integer
    Dim I As Integer
    Dim C As String, lstC As String
    
On Error GoTo errH:


For I = 0 To lstList.ListCount - 1
        C = Chr$(ToSearch)

        If Mid$(lstList.GetItem(I), 6, 3) = " - " Then
            If UCase(Mid$(lstList.GetItem(I), 9, 1)) = C Then
                If I > gb_LstShearchIndex Or C <> gb_LstShearchChr Then
                    SearchList = I
                    lstList.ListIndex = SearchList
                    gb_LstShearchIndex = SearchList
                    gb_LstShearchChr = C
                    Exit Function
                End If
            End If
        Else
            If UCase(Left$(lstList.GetItem(I), 1)) = C Then
                If I > gb_LstShearchIndex Or C <> gb_LstShearchChr Then
                    SearchList = I
                    lstList.ListIndex = SearchList
                    gb_LstShearchIndex = SearchList
                    gb_LstShearchChr = C
                    Exit Function
                End If
            End If
        End If
    Next
    
    SearchList = -1
    gb_LstShearchIndex = SearchList
    gb_LstShearchChr = ""
    
errH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "FunzioniAccessorie.SearchList" & vbCrLf & Err.Description)
    End If
End Function















