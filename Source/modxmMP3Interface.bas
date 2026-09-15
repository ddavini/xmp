Attribute VB_Name = "modxmMP3Interface"
Option Explicit

Private AmpBit As Byte
Private Mode As Byte
Private Freq As Long
Public mGetMP3Info As modInfoMp3.Mp3Info

Public Function mGetCurrentBitRate() As String
    
    mGetCurrentBitRate = xmMP3_getPlayBitRate

End Function

Public Sub mFFT(ByRef SpecTblL() As Long, ByRef SpecTblR() As Long, Mode As Byte)
    Call xmMP3_setFftWindow(Mode)
    Call xmMP3_getSpectrum(SpecTblL(0), SpecTblR(0))
End Sub


Public Function mXNormLevel(Optional ValueSet As Double = 1) As Double
    
    If ValueSet <> 1 Then
        Call xmMP3_setXNormLevel(ValueSet)
    End If
    mXNormLevel = xmMP3_getXNormLevel

End Function

Public Sub mXSrnd(Enabled As Boolean)

    If Enabled Then
        Call xmMP3_setXSoundMode(XSOUND_SURROUND Or xmMP3_getXSoundMode)
    Else
        Call xmMP3_setXSoundMode(xmMP3_getXSoundMode And XSOUND_NORMALIZE)
    End If
    
End Sub

Public Sub mXNorm(Enabled As Boolean)
    Dim Level As Byte
    
    Level = GetINI(cfgFile, "XMMP3", "XNORMLEVEL", 3, 0, 100)
    If Enabled Then
        Call xmMP3_setXSoundMode(xmMP3_getXSoundMode Or XSOUND_NORMALIZE, Level)
    Else
        Call xmMP3_setXSoundMode(xmMP3_getXSoundMode And XSOUND_SURROUND)
    End If

End Sub


Public Function mPlayState() As Long
    Dim Sec As Long
    mPlayState = xmMP3_getState(Sec)
End Function

Public Sub mBufferLengthInSecond(BufferLength As Long)
On Error GoTo ErrH
    If xmMP3_SetBufferSize(BufferLength) Then
    
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mBufferLengthInSecond" & vbCrLf & Err.Description)
    End If
End Sub



Public Sub mSetEQ(pTable As Long)
On Error GoTo ErrH
    
    Call xmMP3_setEqualizer(pTable)
    Call Pause(1, True)
    Call xmMP3_reload

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mSetEQ" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mSetReadHaedA(NEqualFrame As Long, MaxHeadSize As Long)
On Error GoTo ErrH
    If xmMP3_SetReadHaedA(NEqualFrame, MaxHeadSize) Then
    
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mSetReadHaedA" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mAmpBit(lAmpBit As Byte)
On Error GoTo ErrH
    AmpBit = lAmpBit
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mAmpBit" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mMode(lMode As Byte)
On Error GoTo ErrH
    Mode = lMode
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mMode" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mFreq(lFreq As Long)
On Error GoTo ErrH
    Freq = lFreq
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mFreq" & vbCrLf & Err.Description)
    End If
End Sub

Public Function mOpenStream(ByVal StreamName As String) As Boolean
On Error GoTo ErrH
    Dim Data As InputInfo

    gb_Data = Data
    mOpenStream = False
    If xmMP3_open(StreamName, gb_Data) = False Then
        Call DisplayError("StreamName: " & StreamName & vbCrLf & "Can't create stream", _
                vbCritical, "Open Stream", "mOpenStream", True)
    Else
        mGetMP3Info = ReadMP3(StreamName, True, True, gb_Data)
        If mGetMP3Info.BitRate <> 0 Then
            mOpenStream = True
        Else
            mOpenStream = False
        End If
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mOpenStream" & _
                vbCrLf & Err.Description)
    End If
End Function

Public Sub mInitxmMP3(hWnd As Long, SampleRate As Long, Ottobit As Boolean, _
                    Mono As Boolean, DeviceToUse As Integer)
On Error GoTo ErrH
    Dim xmMP3Result As Integer
    Dim Flags As Long
    Dim Opt As DEC_OPTION
    Dim xmMP3Ver As VersionInfo
                     
    ChDrive App.Path
    ChDir App.Path
    
    If Command$ <> "/debug" Then
        Call xmMP3_debug
    End If
    
    If CBool(GetINI(cfgFile, "XMMP3", "FADEIN", "TRUE")) Then
        Call xmMP3_setFadeIn(1)
    End If
    
    If CBool(GetINI(cfgFile, "XMMP3", "FADEOUT", "TRUE")) Then
        Call xmMP3_setFadeOut(1)
    End If

    If xmMP3_getWaveOutSupport = False Then
        Call Err.Raise(vbError, "modxmMP3Interface", _
        "No Wave Out Support")
    End If
    
    Call xmMP3_setWaveOutDeviceId(DeviceToUse)
    
    If xmMP3_init() = False Then
        Call Err.Raise(vbError, "modxmMP3Interface", _
        "Can't initialize digital sound system")
    End If
    
    xmMP3Ver = DisplayVerInfo(App.Path & "\xmMP3.DLL")
    If xmMP3Ver.Major + xmMP3Ver.Minor + xmMP3Ver.Revision < 27 Then
        Call Err.Raise(vbError, "modxmMP3Interface", _
        "xmMP3.DLL version too old")
    End If

    If Not CBool(GetINI(cfgFile, "XMMP3", "NOCALLBACK", "FALSE")) Then
        If xmMP3_callback(AddressOf xmMP3_Proc) = False Then
            Call Err.Raise(vbError, "modxmMP3Interface", _
            "CallBack link error, Please contact the autor")
        End If
    End If
               
    If Mono Then
        Opt.convert = 1
    End If
    If Ottobit Then
        Opt.reduction = 1
    End If
    Opt.freqLimit = SampleRate
    
    Call xmMP3_setEqualizer(ByVal 0&)
       
    If xmMP3_setDecodeOption(Opt) = False Then
        Call Err.Raise(vbError, "modxmMP3Interface", _
        "Can't start digital output")
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
        vbCrLf & "modxmMP3Interface.mInitxmMP3" & vbCrLf & Err.Description)
        Call DeLoad(Brutal:=True)
    End If
    
End Sub

Public Sub mxmMP3BOOST(Value As Integer)
On Error GoTo ErrH

     
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mxmMP3BOOST" & vbCrLf & Err.Description)
    End If
End Sub


Public Sub mTerminatexmMP3()
On Error GoTo ErrH

    WriteINFO "Stop Stream Engine..."
    Call xmMP3_stop
    
    WriteINFO "Stop Callback.."
    Call xmMP3_stopCallback
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mTerminatexmMP3" & vbCrLf & Err.Description)
    End If
End Sub

Public Function mGetxmMP3Version() As String
    Dim xmMP3Ver As VersionInfo
    xmMP3Ver = DisplayVerInfo(App.Path & "\xmMP3.DLL")
    
    With xmMP3Ver
        mGetxmMP3Version = .Major & "." & .Minor & "." & .Revision
    End With
End Function

Public Sub mPlayStream()
    Dim WavePath As String
On Error GoTo ErrH
    
    WavePath = GetINI(cfgFile, "XMMP3", "OTODDIR", App.Path) & "\"
    WavePath = WavePath & ElaboraNomeFile(RetFileName(Dir(GetFile(IndiceGlobalissimo)))) & ".wav"
    
    If CBool(GetINI(cfgFile, "XMMP3", "OUTONDISK", "FALSE")) Then
        If xmMP3_decodeWave(WavePath) = False Then
            Call DisplayError("Can't play stream", vbCritical, "Play Stream", "mPlayStream")
        End If
    Else
        If xmMP3_play = False Then
            Call DisplayError("Can't play stream", vbCritical, "Play Stream", "mPlayStream")
        End If
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mPlayStream" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mStopStream()
On Error GoTo ErrH
    
    If xmMP3_stop = False Then
        Call DisplayError("Can't Stop stream", vbCritical, "Stop Stream", "mStopStream")
    End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStopStream" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mCloseStream()
On Error GoTo ErrH
    If xmMP3_close = False Then
        Call DisplayError("Can't Close stream", vbCritical, "Close Stream", "mCloseStream")
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mCloseStream" & vbCrLf & Err.Description)
    End If
End Sub

                
Public Function mStreamVol(Optional Valore As Long = 0, Optional Mute As Boolean) As Long
    Dim SoftVolume As Boolean
On Error GoTo ErrH
    SoftVolume = CBool(GetINI(cfgFile, "XMMP3", "SOFTVOLUME", "FALSE"))
    
    If SoftVolume Then
        If Valore <> 0 Or Mute Then
            Call xmMP3_setSoftVolume(Valore, Valore)
        End If
        Call xmMP3_getSoftVolume(mStreamVol, Valore)
    Else
        If Valore <> 0 Or Mute Then
            Call xmMP3_setVolume(Valore, Valore)
        End If
        Call xmMP3_getVolume(mStreamVol, Valore)
    End If
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamVol" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamLen() As Long
On Error GoTo ErrH
        
        mStreamLen = xmMP3_getTotalSamples
        If mStreamLen <= 0 Then
            mStreamLen = -1
        End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamLen" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamLenInSeconds() As Long
On Error GoTo ErrH
Dim MpegData As MPEG_INFO
        
        If mStreamLen <> -1 And mGetMP3Info.BitRate <> 0 Then
            'mStreamLenInSeconds = mGetMP3Info.Duration
            If xmMP3_getMpegInfo(MpegData) = False Then
                    
            End If
            mStreamLenInSeconds = MpegData.TotalSec
        Else
            mStreamLenInSeconds = 0
        End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamLenInSecond" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamPos(Optional Valore As Long = 0, Optional Reset As Boolean) As Long
On Error GoTo ErrH

    If Valore <> 0 Or Reset Then
        Call xmMP3_setPlaySamples(Valore)
    End If
    
    mStreamPos = xmMP3_getPlaySamples
    
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamPos" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamPosInSeconds() As Long
On Error GoTo ErrH
    If mGetMP3Info.BitRate <> 0 And mStreamLen <> -1 And _
       mStreamPos <> -1 Then
        Call xmMP3_getState(mStreamPosInSeconds)
    Else
        mStreamPosInSeconds = 0
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamPosInSeconds" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamIsActive() As Boolean
On Error GoTo ErrH
    Dim Sec As Long
    Dim State As Long
    State = xmMP3_getState(Sec)

    Select Case State
        Case Is = 1, 3
            mStreamIsActive = True
        Case Else
            mStreamIsActive = False
    End Select
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamIsActive" & vbCrLf & Err.Description)
    End If
End Function

Public Function mStreamIsPause() As Boolean
On Error GoTo ErrH
    Dim Sec As Long
    Dim State As Long
    State = xmMP3_getState(Sec)

    If State = 2 Then
        mStreamIsPause = True
    Else
        mStreamIsPause = False
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mStreamIsPause" & vbCrLf & Err.Description)
    End If
End Function

Public Function mSinLevel() As Long
On Error GoTo ErrH
    Dim DL(255) As Long
    Dim SL(255) As Long
    Call xmMP3_getWave(SL(0), DL(0))
    mSinLevel = Abs(SL(0))
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mSinLevel" & vbCrLf & Err.Description)
    End If
End Function

Public Function mDesLevel() As Long
On Error GoTo ErrH
    Dim DL(255) As Long
    Dim SL(255) As Long
    Call xmMP3_getWave(SL(0), DL(0))
    mDesLevel = Abs(DL(0))
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mDesLevel" & vbCrLf & Err.Description)
    End If
End Function

Public Sub mPauseStream()
On Error GoTo ErrH
    Call xmMP3_pause
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mPauseStream" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub mResumeStream()
On Error GoTo ErrH
    Call xmMP3_restart
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mResumeStream" & vbCrLf & Err.Description)
    End If
    
End Sub

'Private Function FileExists(FileName) As Boolean
'  On Local Error Resume Next
'  FileExists = (Dir$(FileName) <> "")
'End Function

Public Function mxmMP3INFO() As Long
On Error GoTo ErrH

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modxmMP3Interface.mxmMP3INFO" & vbCrLf & Err.Description)
    End If
End Function
  


