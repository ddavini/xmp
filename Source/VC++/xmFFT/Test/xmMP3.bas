Attribute VB_Name = "modxmMP3"
'xmMP3.BAS
Declare Function xmMP3_init Lib "xmMP3.dll" (Optional ByVal Flag As Long) As Boolean
Declare Function xmMP3_open Lib "xmMP3.dll" (ByVal pszName As String, pInfo As InputInfo) As Boolean
Declare Function xmMP3_close Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_SetBufferSize Lib "xmMP3.dll" (ByVal BufferRedimSize As Long) As Boolean
Declare Function xmMP3_SetReadHaedA Lib "xmMP3.dll" (ByVal NEqualFram As Long) As Boolean
Declare Function xmMP3_play Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_stop Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_pause Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_restart Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_seek Lib "xmMP3.dll" (ByVal Sec As Long) As Boolean
Declare Function xmMP3_setPlayFrames Lib "xmMP3.dll" (ByVal Frames As Long) As Boolean
Declare Function xmMP3_setPlaySamples Lib "xmMP3.dll" (ByVal samples As Long) As Boolean
Declare Function xmMP3_reload Lib "xmMP3.dll" () As Boolean

'Obsolete
'Declare Function xmMP3_getVersion Lib "xmMP3.dll" () As Long
Declare Function xmMP3_setxmMP3Option Lib "xmMP3.dll" (pxmMP3Option As xmMP3_OPTION) As Long
Declare Sub xmMP3_getxmMP3Option Lib "xmMP3.dll" (pxmMP3Option As xmMP3_OPTION)
Declare Function xmMP3_setDecodeOption Lib "xmMP3.dll" (pDecOption As DEC_OPTION) As Long
Declare Sub xmMP3_getDecodeOption Lib "xmMP3.dll" (pDecOption As DEC_OPTION)
Declare Sub xmMP3_setWaveOutDeviceId Lib "xmMP3.dll" (ByVal ID As Long)

Declare Function xmMP3_getState Lib "xmMP3.dll" (Sec As Long) As Long
Declare Function xmMP3_getPlayFrames Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getPlaySamples Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getTotalSamples Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getWinampPlayMs Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getWinampTotalSec Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getPlayBitRate Lib "xmMP3.dll" () As Long

Declare Function xmMP3_getFileType Lib "xmMP3.dll" (ByVal pszName As String) As Long
Declare Function xmMP3_getWaveOutSupport Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getSilentFrames Lib "xmMP3.dll" (ByVal pszName As String) As Long
Declare Function xmMP3_getMpegInfo Lib "xmMP3.dll" (pMpegTagInfo As MPEG_INFO) As Boolean
Declare Function xmMP3_getFileInfo Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO, pMpegTagInfo As MPEG_INFO) As Boolean
Declare Function xmMP3_getFileInfo2 Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO, pMpegTagInfo As MPEG_INFO, pListInfo As LIST_INFO) As Boolean

'** ID3 Tag **
Declare Function xmMP3_setTagInfo Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO, Optional ByVal tagSet As Long = 0, Optional ByVal tagAdd As Long = 0) As Boolean
Declare Function xmMP3_getTagInfo Lib "xmMP3.dll" (pTagInfo As TAG_INFO) As Boolean
Declare Function xmMP3_setTagInfoEX Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO_11, Optional ByVal tagSet As Long = 0, Optional ByVal tagAdd As Long = 0) As Boolean
Declare Function xmMP3_getFileTagInfo Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO) As Boolean
Declare Function xmMP3_getFileTagInfoEX Lib "xmMP3.dll" (ByVal pszName As String, pTagInfo As TAG_INFO_11) As Boolean
Declare Function xmMP3_getGenre Lib "xmMP3.dll" (pTagInfo As TAG_INFO) As Boolean
Declare Function xmMP3_cutID3v2 Lib "xmMP3.dll" (ByVal pszName As String) As Boolean

Declare Function xmMP3_setListInfo Lib "xmMP3.dll" (ByVal pszName As String, pListInfo As LIST_INFO) As Boolean
Declare Function xmMP3_getListInfo Lib "xmMP3.dll" (pListInfo As LIST_INFO) As Boolean
Declare Function xmMP3_setListInfoExVB Lib "xmMP3.dll" (ByVal pszName As String, pListInfo As LIST_INFO_EX_VB) As Boolean
Declare Function xmMP3_getListInfoExVB Lib "xmMP3.dll" (ByVal pszName As String, pListInfo As LIST_INFO_EX_VB) As Boolean
Declare Function xmMP3_setListInfoEX2 Lib "xmMP3.dll" (ByVal pszName As String, pListInfo As LIST_INFO_EX2) As Boolean
Declare Sub xmMP3_setNotDataSiWrite Lib "xmMP3.dll" (writeFlag As Long)

'** (Lyrics3) **
Declare Function xmMP3_delFileLyrics3Info Lib "xmMP3.dll" (ByVal pszName As String) As Boolean
Declare Function xmMP3_setFileLyrics3InfoVB Lib "xmMP3.dll" (ByVal pszName As String, pLyrics3Info As LYRICS3_INFO_VB, _
                                                             ByVal lyrData As String, ByVal InfData As String, ByVal ImgData As String) As Boolean
Declare Function xmMP3_getFileLyrics3InfoVB Lib "xmMP3.dll" (ByVal pszName As String, pLyrics3Info As LYRICS3_INFO_VB, _
                                                             ByVal lyrData As String, ByVal InfData As String, ByVal ImgData As String) As Boolean
Declare Function xmMP3_readLyrics3Data Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_setReadLyrics3Info Lib "xmMP3.dll" (ByVal pszName As String) As Boolean
Declare Sub xmMP3_setLyrics3InsField Lib "xmMP3.dll" (ByVal field As Long)
Declare Sub xmMP3_setLyrics3Use Lib "xmMP3.dll" (ByVal useLyrics3 As Long)

Declare Function xmMP3_setLyricsFile Lib "xmMP3.dll" (ByVal pszLyricsName As String) As Boolean
Declare Function xmMP3_getLyrics Lib "xmMP3.dll" (pLyricsInfo As LYRICS_INFO) As Boolean
Declare Function xmMP3_getLyrics2 Lib "xmMP3.dll" (pLyricsInfo As LYRICS_INFO2) As Boolean
Declare Function xmMP3_clearLyrics Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_checkKaraokeTag Lib "xmMP3.dll" () As Boolean
Declare Sub xmMP3_setKaraokeUse Lib "xmMP3.dll" (ByVal useKaraoke As Long)
Declare Sub xmMP3_setLyricsTime Lib "xmMP3.dll" (ByVal Flag As Long)
Declare Sub xmMP3_setNotTagLyricsLine Lib "xmMP3.dll" (ByVal readFlag As Long)
Declare Sub xmMP3_setLyricsAdjustTime Lib "xmMP3.dll" (ByVal MS As Long)
Declare Sub xmMP3_setLyricsNextAdjustTime Lib "xmMP3.dll" (ByVal MS As Long)
Declare Function xmMP3_convTimeN2W Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_convTimeW2N Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_outLyricsFile Lib "xmMP3.dll" (ByVal pszLyricsName As String) As Boolean
Declare Sub xmMP3_useAtMarkTag Lib "xmMP3.dll" (ByVal useFlag As Long)
Declare Sub xmMP3_setTimeRatio Lib "xmMP3.dll" (ByVal TimeRatio As Double)
Declare Function xmMP3_checkTimeTag Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getCallbackLyricsData Lib "xmMP3.dll" (ByVal pProc As Long) As Boolean
Declare Function xmMP3_getAtTagData Lib "xmMP3.dll" (pAtTagData As AT_TAG_DATA) As Boolean
Declare Function xmMP3_getLyricsPoint Lib "xmMP3.dll" () As Long
'Declare Function xmMP3_getLyricsData Lib "xmMP3.dll" (pLyricsData As LYRICS_DATA, ByVal elmNo As Long) As Long
'Declare Function xmMP3_getLyricsCount Lib "xmMP3.dll" () As Long

Declare Function xmMP3_setStepPitch Lib "xmMP3.dll" (ByVal pitch As Long, Optional ByVal flames As Long = 5) As Boolean
Declare Function xmMP3_getStepPitch Lib "xmMP3.dll" () As Long
Declare Function xmMP3_setPitch Lib "xmMP3.dll" (ByVal pitch As Long) As Boolean
Declare Function xmMP3_getPitch Lib "xmMP3.dll" () As Long
Declare Sub xmMP3_setFadeIn Lib "xmMP3.dll" (ByVal fin As Long)
Declare Sub xmMP3_setFadeOut Lib "xmMP3.dll" (ByVal fout As Long)
Declare Sub xmMP3_fadeOut Lib "xmMP3.dll" ()
Declare Sub xmMP3_setOverTime Lib "xmMP3.dll" (ByVal tmOver As Long)
'Declare Sub xmMP3_setSeekPlay Lib "xmMP3.dll" (ByVal seekPlay As Long)
Declare Sub xmMP3_setFrameReadFlag Lib "xmMP3.dll" (ByVal readFlag As Long)
'Declare Sub xmMP3_quickSeek Lib "xmMP3.dll" (ByVal quickSeekFlag As Long)

Declare Function xmMP3_setVolume Lib "xmMP3.dll" (ByVal lVol As Long, ByVal rVol As Long) As Boolean
Declare Function xmMP3_getVolume Lib "xmMP3.dll" (lVol As Long, rVol As Long) As Boolean
Declare Sub xmMP3_setSoftVolume Lib "xmMP3.dll" (ByVal lVol As Long, ByVal rVol As Long)
Declare Sub xmMP3_getSoftVolume Lib "xmMP3.dll" (lVol As Long, rVol As Long)

Declare Function xmMP3_playDecodeWave Lib "xmMP3.dll" (ByVal pszWaveName As String) As Boolean
Declare Function xmMP3_decodeWave Lib "xmMP3.dll" (ByVal pszWaveName As String) As Boolean

'Declare Function xmMP3_encodeOpen Lib "xmMP3.dll" (ByVal pszWaveName As String, pWaveForm As WAVE_FORM) As Boolean
'Declare Function xmMP3_encodeStart Lib "xmMP3.dll" (ByVal pszMp3Name As String) As Boolean
'Declare Function xmMP3_encodeStop Lib "xmMP3.dll" () As Boolean
'Declare Function xmMP3_getEncodeState Lib "xmMP3.dll" (readSize As Long, encodeSize As Long) As Long

Declare Sub xmMP3_getWave Lib "xmMP3.dll" (pWaveL As Long, pWaveR As Long)
Declare Sub xmMP3_getSpectrum Lib "xmMP3.dll" (pSpecL As Long, pSpecR As Long)
Declare Sub xmMP3_setFftWindow Lib "xmMP3.dll" (ByVal window As Long)
Declare Sub xmMP3_setEqualizer Lib "xmMP3.dll" (pTable As Long)

Declare Function xmMP3_changeWav Lib "xmMP3.dll" (ByVal pszName As String) As Boolean
Declare Function xmMP3_changeRmp Lib "xmMP3.dll" (ByVal pszName As String) As Boolean
Declare Function xmMP3_changeMp3 Lib "xmMP3.dll" (ByVal pszName As String) As Boolean
Declare Function xmMP3_cutMacBinary Lib "xmMP3.dll" (ByVal pszName As String) As Boolean

Declare Function xmMP3_callback Lib "xmMP3.dll" (ByVal pProc As Long) As Boolean
Declare Function xmMP3_startCallback Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_stopCallback Lib "xmMP3.dll" () As Boolean
'Declare Function xmMP3_startCallBackTimer Lib "xmMP3.dll" (ByVal pProc As Long, ByVal wDelay As Long, ByVal wResolution As Long, ByVal userData As Long) As Boolean
'Declare Function xmMP3_stopCallBackTimer Lib "xmMP3.dll" (ByVal pProc As Long) As Boolean

Declare Function xmMP3_getLastErrorNo Lib "xmMP3.dll" () As Long
Declare Function xmMP3_debug Lib "xmMP3.dll" () As Long

Declare Function xmMP3_setPlayFlames Lib "xmMP3.dll" (ByVal flames As Long) As Boolean
Declare Function xmMP3_getPlayFlames Lib "xmMP3.dll" () As Long

Declare Sub xmMP3_startAnalyzeThread Lib "xmMP3.dll" ()
Declare Sub xmMP3_stopAnalyzeThread Lib "xmMP3.dll" ()
Declare Sub xmMP3_startAnalyze Lib "xmMP3.dll" ()
Declare Sub xmMP3_stopAnalyze Lib "xmMP3.dll" ()
Declare Function xmMP3_getWaveData Lib "xmMP3.dll" (pWaveData As WAVE_DATA) As Boolean

Public Type InputInfo
    szTrackName As String * 128
    szArtistName As String * 128
    channels As Long
    BitRate As Long                 'kbit/s
    samplingRate As Long            'Hz
    TotalSec As Long                's
End Type

Public Type TAG_INFO
    szTrackName As String * 128
    szArtistName As String * 128
    szAlbumName As String * 128
    szYear As String * 5
    szComment As String * 128
    Genre As Long
    szGenreName As String * 128
End Type

Public Type TAG_INFO_11
    szTrackName As String * 128
    szArtistName As String * 128
    szAlbumName As String * 128
    szYear As String * 5
    szComment As String * 128
    Genre As Long
    szGenreName As String * 128
    trackNo As Long                '-1 = v1.0
End Type

Public Type MPEG_INFO
    Version As Long                 '1:MPEG-1, 2:MPEG-2, 3:MPEG-2.5
    Layer As Long                   '1:Layer1. 2:Layer2, 3:Layer3
    crcDisable As Long
    extension As Long
    Mode As Long                    '0:Stereo, 1:Joint stereo, 3:Dual channel, 4:Mono
    CopyRight As Long
    Original As Long
    Emphasis As Long                '0:None, 1:50/15ms, 2:Reserved, 3:CCITT j.17
    
    channels As Long
    BitRate As Long                 'kbit/s (0 VBR)
    samplingRate As Long            'Hz
    fileSize As Long                'Byte
    flames As Long
    TotalSec As Long                's
End Type

Public Type DEC_OPTION
    reduction As Long               '0:1/1 1:1/2 2:1/4 [Default = 0]
    convert As Long
    freqLimit As Long               'Default = 24000
End Type

Public Type xmMP3_OPTION
    inputBlock As Long              'Default = 40
    outputBlock As Long             'Default = 30
    inputSleep As Long              'Default = 5
    outputSleep As Long             'Default = 0
End Type

Public Type LIST_INFO
    INAM As String * 128
    IART As String * 128
    IPRD As String * 128
    ICMT As String * 128
    ICRD As String * 128
    IGNR As String * 128
    ICOP As String * 128
    IENG As String * 128
    ISRC As String * 128
    ISFT As String * 128
    IKEY As String * 128
    ITCH As String * 128
    ICMS As String * 128
    ILYC As String * 128
End Type

Public Type LIST_INFO_EX_VB
    INAM As String * 512
    IART As String * 512
    IPRD As String * 512
    ICMT As String * 512
    ICRD As String * 512
    IGNR As String * 512
    ICOP As String * 512
    IENG As String * 512
    ISRC As String * 512
    ISFT As String * 512
    IKEY As String * 512
    ITCH As String * 512
    ICMS As String * 512
    IMED As String * 512
    ISBJ As String * 512
    IMP3 As String * 512
    ILYC As String * 512
End Type

Public Type LIST_INFO_EX2
    INAM As String
    IART As String
    IPRD As String
    ICMT As String
    ICRD As String
    IGNR As String
    ICOP As String
    IENG As String
    ISRC As String
    ISFT As String
    IKEY As String
    ITCH As String
    ICMS As String
    IMED As String
    ISBJ As String
    IMP3 As String
    ILYC As String
End Type

Public Type LYRICS_INFO
    Sec As Long
    LyricsNext2 As String * 128
    LyricsNext1 As String * 128
    LyricsCurrent As String * 128
    LyricsPrev1 As String * 128
    LyricsPrev2 As String * 128
End Type

Public Type LYRICS_INFO2
    Sec As Long
    lineno As Long
    point As Long
    length As Long
    LyricsNext2 As String * 128
    LyricsNext1 As String * 128
    LyricsCurrent As String * 128
    LyricsCurrentBegin As String * 128
    LyricsCurrentLyrics As String * 128
    LyricsCurrentAll As String * 128
    LyricsPrev1 As String * 128
    LyricsPrev2 As String * 128
End Type

Public Type WAVE_DATA
    channels As Long
    bitsPerSample As Long
    Left As Long
    Right As Long
End Type

Public Type WAVE_FORM
    channels As Long
    bitsPerSample As Long
    samplingRate As Long
    dataSize As Long
End Type

Public Type LYRICS3_INFO_VB
    IND_LYR As String * 1
    IND_TIMETAG As String * 1
    AUT As String * 250
    EAL As String * 250
    EAR As String * 250
    ETT As String * 250
End Type

'Public lyrData As String * 60000
'Public InfData As String * 60000
'Public ImgData As String * 60000

Public Type AT_TAG_DATA
    Artist As String * 1024
    Title As String * 1024
    Album As String * 1024
    Bgfile As String * 1024
    Bgfolder As String * 1024
    TimeRatio As Double
    Offset As Long                  'ms
    SilencemSec As Long             'ms
    TaggingBy As String * 1024
    EditedBy As String * 1024
    Silence As Long
    flames As Long
    TotalSec As Long                's
    TimeType As String * 1024       'WinAmp or Normal
End Type

Public Type LYRICS_DATA
    no As Long
    Sec As Long
    lineno As Long
    point As Long
    tagFlag As Boolean
    textSize As Long            'byte
    lyrics As String * 128
End Type

Public Const FT_NOMAL = 0
Public Const FT_WAVE = 1
Public Const FT_RMP = 2
Public Const FT_ID3V2 = 4
Public Const FT_MAC = 8
Public Const FT_ID3V1 = 16
Public Const FT_LYRICS3 = 32

Public Const MSG_ERROR = 0
Public Const MSG_STOPING = 1
Public Const MSG_PLAYING = 2
Public Const MSG_PAUSING = 3
Public Const MSG_PLAYDONE = 4

Enum xmMP3_errNo
    ERR_MP3_FILE_OPEN = 1
    ERR_MP3_FILE_NOT_OPEN = 2
    ERR_MP3_FILE_READ = 3
    ERR_MP3_FILE_WRITE = 4
    ERR_WAV_FILE_OPEN = 5
    ERR_WAV_FORMAT = 6
    ERR_ENCODE_FILE_OPEN = 7
    ERR_LYRICS_FILE_OPEN = 8
    ERR_LYRICS_NON_DATA = 9
    ERR_FRAME_HEADER_NOT_FOUND = 10
    ERR_FRAME_HEADER_READ = 11
    ERR_STATE_STOP = 12
    ERR_NOT_STATE_STOP = 13
    ERR_NOT_STATE_PLAY = 14
    ERR_STATE_NON_ENCODE = 15
    ERR_PLAY = 16
    ERR_STOP = 17
    ERR_INVALID_VALUE = 18
    ERR_MALLOC = 19
    ERR_NON_RIFF = 20
    ERR_RIFF = 21
    ERR_NOT_MP3 = 22
    ERR_MAC_BIN = 23
    ERR_UNKNOWN_FILE = 24
    ERR_OPEN_OUT_DEVICE = 25
    ERR_DECODE = 26
    ERR_DECODE_THREAD = 27
    ERR_ENCODE_THREAD = 28
    ERR_CREATE_EVENT = 29
    ERR_CODEC_NOT_FOUND = 30
    ERR_WAVE_TABLE_NOT_FOUND = 31
    ERR_ACM_OPEN = 32
End Enum

Enum fft_window
    rectangle = 0
    hanning = 1
    hamming = 2
    blackman = 3
End Enum


'---------------------------------------------------------
'Function NTrim()
'---------------------------------------------------------
Function NTrim(Word As String) As String
    If InStr(Word, Chr(0)) > 0 Then
        NTrim = Left(Word, InStr(Word, Chr(0)) - 1)
    Else
        NTrim = Word
    End If
End Function

'---------------------------------------------------------
'Function LNTrim()
'---------------------------------------------------------
Function LNTrim(Word As String) As String
    Dim textLength As Long
    Dim cnt As Long
    textLength = Len(Word)
    
    LNTrim = Word
    If InStr(LNTrim, Chr(0)) > 0 Then
        For cnt = 1 To textLength
            If Mid(LNTrim, 1, 1) = Chr(0) Or Mid(LNTrim, 1, 1) = vbCr Or Mid(LNTrim, 1, 1) = vbLf Then
                If cnt < textLength Then
                    LNTrim = Mid(LNTrim, 2)
                Else
                    LNTrim = ""
                End If
            Else
                Exit For
            End If
        Next cnt
    Else
        LNTrim = Word
    End If
End Function

'---------------------------------------------------------
'Function NTrim2()
'---------------------------------------------------------
Function NTrim2(Word As String) As String

    Word = LNTrim(Word)
    NTrim2 = NTrim(Word)

End Function

