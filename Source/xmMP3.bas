Attribute VB_Name = "modxmMP3"
'xmMP3.BAS
Declare Function xmMP3_init Lib "xmMP3.dll" (Optional ByVal Flag As Long) As Boolean
Declare Function xmMP3_open Lib "xmMP3.dll" (ByVal pszName As String, pInfo As InputInfo) As Boolean
Declare Function xmMP3_close Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_SetBufferSize Lib "xmMP3.dll" (ByVal BufferRedimSize As Long) As Boolean
Declare Function xmMP3_SetReadHaedA Lib "xmMP3.dll" (ByVal NEqualFram As Long, ByVal MaxHeadSize As Long) As Boolean
Declare Function xmMP3_play Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_stop Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_pause Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_restart Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_seek Lib "xmMP3.dll" (ByVal Sec As Long) As Boolean
Declare Function xmMP3_setPlayFrames Lib "xmMP3.dll" (ByVal Frames As Long) As Boolean
Declare Function xmMP3_setPlaySamples Lib "xmMP3.dll" (ByVal samples As Long) As Boolean
Declare Function xmMP3_reload Lib "xmMP3.dll" () As Boolean

'Obsolete'''''''''''''''''''''''''''''''''''''''''''''''''''
Declare Function xmMP3_getVersion Lib "xmMP3.dll" () As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Declare Function xmMP3_getWaveOutSupport Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getMpegInfo Lib "xmMP3.dll" (pMpegTagInfo As MPEG_INFO) As Boolean

Declare Sub xmMP3_setXSoundMode Lib "xmMP3.dll" (Optional ByVal Mode As Long, Optional ByVal Options As Long)
Declare Function xmMP3_getXSoundMode Lib "xmMP3.dll" () As Long
Declare Function xmMP3_getXNormLevel Lib "xmMP3.dll" () As Double
Declare Function xmMP3_setXNormLevel Lib "xmMP3.dll" (ByVal Level As Double) As Double

Declare Function xmMP3_setStepPitch Lib "xmMP3.dll" (ByVal pitch As Long, Optional ByVal flames As Long = 5) As Boolean
Declare Function xmMP3_getStepPitch Lib "xmMP3.dll" () As Long
Declare Function xmMP3_setPitch Lib "xmMP3.dll" (ByVal pitch As Long) As Boolean
Declare Function xmMP3_getPitch Lib "xmMP3.dll" () As Long
Declare Sub xmMP3_setFadeIn Lib "xmMP3.dll" (ByVal fin As Long)
Declare Sub xmMP3_setFadeOut Lib "xmMP3.dll" (ByVal fout As Long)
Declare Sub xmMP3_fadeOut Lib "xmMP3.dll" ()
Declare Sub xmMP3_setOverTime Lib "xmMP3.dll" (ByVal tmOver As Long)

Declare Function xmMP3_setVolume Lib "xmMP3.dll" (ByVal lVol As Long, ByVal rVol As Long) As Boolean
Declare Function xmMP3_getVolume Lib "xmMP3.dll" (lVol As Long, rVol As Long) As Boolean
Declare Sub xmMP3_setSoftVolume Lib "xmMP3.dll" (ByVal lVol As Long, ByVal rVol As Long)
Declare Sub xmMP3_getSoftVolume Lib "xmMP3.dll" (lVol As Long, rVol As Long)

Declare Function xmMP3_playDecodeWave Lib "xmMP3.dll" (ByVal pszWaveName As String) As Boolean
Declare Function xmMP3_decodeWave Lib "xmMP3.dll" (ByVal pszWaveName As String) As Boolean

Declare Sub xmMP3_getWave Lib "xmMP3.dll" (pWaveL As Long, pWaveR As Long)
Declare Sub xmMP3_getSpectrum Lib "xmMP3.dll" (pSpecL As Long, pSpecR As Long)
Declare Sub xmMP3_setFftWindow Lib "xmMP3.dll" (ByVal window As Long)
Declare Sub xmMP3_setEqualizer Lib "xmMP3.dll" (pTable As Long)

Declare Function xmMP3_callback Lib "xmMP3.dll" (ByVal pProc As Long) As Boolean
Declare Function xmMP3_startCallback Lib "xmMP3.dll" () As Boolean
Declare Function xmMP3_stopCallback Lib "xmMP3.dll" () As Boolean

Declare Function xmMP3_getLastErrorNo Lib "xmMP3.dll" () As Long
Declare Function xmMP3_debug Lib "xmMP3.dll" () As Long

Declare Function xmMP3_setPlayFlames Lib "xmMP3.dll" (ByVal flames As Long) As Boolean
Declare Function xmMP3_getPlayFlames Lib "xmMP3.dll" () As Long

Declare Function xmMP3_getWaveData Lib "xmMP3.dll" (pWaveData As WAVE_DATA) As Boolean

Public Const XSOUND_DISABLE As Byte = 0
Public Const XSOUND_SURROUND As Byte = 1
Public Const XSOUND_NORMALIZE As Byte = 2

Public Type InputInfo
    szTrackName As String * 128
    szArtistName As String * 128
    channels As Long
    BitRate As Long                 'kbit/s
    samplingRate As Long            'Hz
    TotalSec As Long                's
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
    FileSize As Long                'Byte
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

Public Type WAVE_DATA
    channels As Long
    bitsPerSample As Long
    Left As Long
    Right As Long
End Type

Enum Play_Status
    STATE_STOP = 0
    STATE_PLAY = 1
    STATE_PAUSE = 2
    STATE_SEEK = 3
End Enum

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
