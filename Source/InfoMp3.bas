Attribute VB_Name = "modInfoMp3"
Option Explicit
''''''''''''''''''''''''''''''''''''''''
'Based on Michael KarathaNasis Sample  '
'http://home12.iNet.tele.dk/mkaratha   '
''''''''''''''''''''''''''''''''''''''''

'Pesantemente riveduto

Private Type WAVEFORMATEX
    wFormatTag As Long
    nChannels As Long
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Long
    wBitsPerSample As Long
    cbSize As Long
End Type

Private Type FileFormat
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
End Type

Private Type FileHeader
    dwRiff As Long
    dwFileSize As Long
    dwWave As Long
    dwFormat As Long
    dwFormatLength As Long
End Type


Public Type Mp3Info
   BitRate              As Integer
   Frequency            As Long
   Mode                 As String
   Emphasis             As String
   MpegVersion          As Single
   MpegLayer            As Integer
   Padding              As String
   CRC                  As String
   Duration             As Long
   CopyRight            As String
   Original             As String
   PrivateBit           As String
   HasTag               As Boolean
   Tag                  As String
   SongName             As String
   Artist               As String
   Album                As String
   Year                 As String
   Comment              As String
   Genre                As Integer
   track                As String
   VBR                  As Boolean
   Frames               As Integer
End Type

Private GetMP3Info As Mp3Info
'Buffer di immagazzinamento dell'Header
Private X() As Byte
Private MaxCapa As Long

Private Sub SetMaxCapa()
    MaxCapa = GetINI(cfgFile, "PREFERENCE", "HMAXSIZE", "102400")
    If MaxCapa > 999999 Or MaxCapa < 0 Then
        MaxCapa = 999999
    End If
End Sub
'This Function Converts Binary String to Decimal Integer
Private Function BinToDec(BinValue As String) As Long
    Dim I As Integer
    Dim BinValueLen As Long
On Error GoTo ErrH
    BinValueLen = Len(BinValue)
   BinToDec = 0
   For I = 1 To Len(BinValue)
      If Mid$(BinValue, I, 1) = "1" Then
         BinToDec = BinToDec + 2 ^ (BinValueLen - I)
      End If
   Next I
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modInfoMp3.BinToDec" & vbCrLf & Err.Description)
    End If
End Function

Public Function ByteToBit(ByteArray) As String
Dim Z As Integer
Dim I As Integer
   'convert 4*1 byte array to 4*8 bits
   ByteToBit = ""
   For Z = 0 To 3
      For I = 7 To 0 Step -1
         If Int(ByteArray(Z) / (2 ^ I)) = 1 Then
            ByteToBit = ByteToBit & "1"
            ByteArray(Z) = ByteArray(Z) - (2 ^ I)
         Else
            If ByteToBit <> "" Then
               ByteToBit = ByteToBit & "0"
            End If
         End If
      Next
   Next Z
End Function

Public Function BinaryHeader(FileName As String, _
                             ReadHeader As Boolean, _
                             Optional ByRef StartByte As Long = 0) As String
    Dim ByteArray(3) As Byte
    Dim XingH As String
    Dim FIO As Long
    Dim N As Long
    Dim I As Long
    Dim Z As Integer
    Dim Headstart As Long
    Dim Frames As Integer
    Dim P As Integer

On Error Resume Next
   
    N = FileLen(FileName)
    If N < 256 Then
         Exit Function
    End If
    If ReadHeader Then

        For I = StartByte To MaxCapa
           If X(I) = 255 Then
                If X(I + 1) >= 224 And X(I + 1) <= 255 Then
                    If X(I + 2) < 252 Then
                       Headstart = I
                       StartByte = I
                       Exit For
                    End If
                End If
            Else
                StartByte = I
           End If
        Next I

     
        ''start check for XingHeader'''
        XingH = ""
        For I = 0 To 3
            XingH = XingH + Chr$(X(Headstart + 36 + I))
        Next
        If XingH = "Xing" Then
            GetMP3Info.VBR = True
            FIO = FreeFile
            Open FileName For Binary Access Read As FIO
                Get #FIO, Headstart + 45, ByteArray
            Close #FIO
            Frames = BinToDec(ByteToBit(ByteArray))
            GetMP3Info.Frames = Frames
        Else
            GetMP3Info.VBR = False
        End If
        '''end check for XingHeader'''
     
        Call CopyMemory(ByteArray(0), X(Headstart), 4)
    End If

    BinaryHeader = ByteToBit(ByteArray)
    If Err.Number <> 0 Then
        BinaryHeader = "Error: " & Err.Description
        Err.Clear
    End If
End Function

Private Sub GetTAG(FileName As String)
    Dim FIO As Long
    Dim N As Long
    Dim P As Integer
    Dim INbuf As String * 256

On Error Resume Next
   
    N = FileLen(FileName)
    ''''start id3 tag''''
    FIO = FreeFile
    Open FileName For Binary Access Read As FIO
       Get #FIO, (N - 255), INbuf
    Close #FIO
    P = InStr(1, INbuf, "tag", 1)  'Ny
    If P = 0 Then
       With GetMP3Info
          .HasTag = False
          .SongName = ""
          .Artist = ""
          .Album = ""
          .Year = ""
          .Comment = ""
          .track = ""
          .Genre = 255
       End With
    Else
       With GetMP3Info
          .HasTag = True
          .SongName = Replace$(RTrim$(Mid$(INbuf, P + 3, 30)), vbNullChar, "")
          .Artist = Replace$(RTrim$(Mid$(INbuf, P + 33, 30)), vbNullChar, "")
          .Album = Replace$(RTrim$(Mid$(INbuf, P + 63, 30)), vbNullChar, "")
          .Year = Replace$(RTrim$(Mid$(INbuf, P + 93, 4)), vbNullChar, "")
          .Comment = Replace$(RTrim$(Mid$(INbuf, P + 97, 29)), vbNullChar, "")
          .track = Replace$(RTrim$(Mid$(INbuf, P + 126, 1)), vbNullChar, "")
          .Genre = Replace$(Asc(RTrim$(Mid$(INbuf, P + 127, 1))), vbNullChar, "")
       End With
    End If
    If Err.Number <> 0 Then
        Err.Clear
    End If
    ''''stop id3 tag''''
End Sub

Public Function ReadMP3(FileName As String, ReadTAG As Boolean, ReadHeader As Boolean, _
                        Data As InputInfo) As Mp3Info
    
    Dim FileSize As Long
    Dim Volte As Integer
    Dim Ok As Boolean
    Dim OLDSamplerate As Long
    Dim OLDBitrate As Long
    Dim MaxVolte As Integer
    Dim CurrentMp3 As Boolean
    Dim LayerVersioN As String
    Dim BIN As String
    Dim I As Integer
    Dim Version(), BRate() As Variant
    Dim Layer(), Freq() As Variant
    Dim EsteNsioNeValida As Integer
    Dim MpegVersion As Single, MpegLayer As Integer
    Dim SMode(), Emph() As Variant
    Dim Mode As String, Emphasis As String
    Dim Frequency As Single
    Dim Temp(), NoYes() As Variant
    Dim BitRate As Long, MS As Long
    Dim Duration As Integer
    Dim Original As String
    Dim CopyRight As String
    Dim Padding As String
    Dim PrivateBit As String, CRC As String
    Dim StartByte As Long
    Dim Sinc As Integer
    Dim FIO As Integer
    Dim WaveFormat As WAVEFORMATEX
    
On Error GoTo ErrH
    
    If Not FileExists(FileName) Then
        Err.Raise vbError, Description:="File Not Found"
    End If
    
    If gb_InfoMp3Active Then
        Exit Function
    Else
        gb_InfoMp3Active = True
    End If
        
    StartByte = 0
    For I = 0 To UBound(Estensioni)
        EsteNsioNeValida = EsteNsioNeValida + _
                           InStr(1, UCase$(FileName), _
                           Estensioni(I))
    Next
    Call SetMaxCapa
    Call AzzeraStrutturaMp3Info(GetMP3Info)
    If EsteNsioNeValida > 0 Then
        Select Case Right$(UCase(FileName), 4)
            Case Is = ".WAV"
                If GetFile(IndiceGlobalissimo) = FileName And _
                    Data.samplingRate <> 0 And Data.BitRate <> 0 Then
                    GetMP3Info.BitRate = (Data.samplingRate * Data.BitRate _
                        * Data.channels / 8) / 1024
                    GetMP3Info.Duration = Data.TotalSec
                    GetMP3Info.Frequency = Data.samplingRate
                    If Data.channels = 2 Then
                        GetMP3Info.Mode = "Stereo"
                    Else
                        GetMP3Info.Mode = "Mono"
                    End If
                    GetMP3Info.PrivateBit = Data.BitRate
                Else
                    WaveFormat = FillFormat(FileName)
                    GetMP3Info.BitRate = WaveFormat.nAvgBytesPerSec / 1024
                    GetMP3Info.Duration = FileLen(FileName) / WaveFormat.nAvgBytesPerSec
                    GetMP3Info.Frequency = WaveFormat.nSamplesPerSec
                    GetMP3Info.PrivateBit = WaveFormat.wBitsPerSample
                    If WaveFormat.nChannels = 2 Then
                        GetMP3Info.Mode = "Stereo"
                    Else
                        GetMP3Info.Mode = "Mono"
                    End If
                End If
            Case Else
                WriteINFO "Header Reading... "
                
                FileSize = FileLen(FileName)

                If MaxCapa >= FileSize Then
                    MaxCapa = FileSize
                End If
                
                ReDim X(MaxCapa)
                
                FIO = FreeFile
                Open FileName For Binary Access Read As FIO
                    If Err.Number <> 0 Then
                        WriteINFO "File Open Filed"
                        Exit Function
                    End If
                    Get #FIO, , X
                Close #FIO
                
                MaxVolte = GetINI(cfgFile, "PREFERENCE", "HINFOS", "128")
                If MaxVolte > 1024 Then
                    MaxVolte = 1024
                ElseIf MaxVolte < 0 Then
                    MaxVolte = 0
                End If
                MaxVolte = MaxVolte / 256
                
                StartByte = -1
                CurrentMp3 = GetFile(IndiceGlobalissimo) = FileName
                
                Version = Array(25, 0, 2, 1) 'Mpegversion table
                Layer = Array(0, 3, 2, 1) 'layer table
                SMode = Array("Stereo", "Joint Stereo", "Dual Channel", "Single Channel") 'mode table
                Emph = Array("No", "50/15", "reserved", "CCITT J 17") 'empasis table
                NoYes = Array("No", "Yes")
                Temp = Array(, 12, 144, 144) 'defiNe to calculate correct bitrate
                Freq = Array(44100, 48000, 32000, -1)
                
                Ok = False
                
                Do
                    StartByte = StartByte + 1
                    DoEvents
                    'Sleep 1
                    WriteINFO "Scan at " & StartByte & "/" & Volte
                    
                    BIN = BinaryHeader(FileName, ReadHeader, StartByte)
                                    
                    
                    If Left$(BIN, 5) = "Error" Then
                        WriteINFO BIN
                        Call AzzeraStrutturaMp3Info(GetMP3Info)
                    Else
                        
                        If ReadHeader = False Then
                            Exit Function
                        End If
                    
                        MpegVersion = Version(BinToDec(Mid$(BIN, 12, 2)))
                        MpegLayer = Layer(BinToDec(Mid$(BIN, 14, 2)))
                        
                        Select Case MpegVersion
                            Case 1
                                Frequency = BinToDec(Mid$(BIN, 21, 2))
                                Frequency = Freq(Frequency)
                            Case 2
                                Frequency = BinToDec(Mid$(BIN, 21, 2))
                                Frequency = Freq(Frequency) \ 2
                            Case 25
                                Frequency = BinToDec(Mid$(BIN, 21, 2))
                                Frequency = Freq(Frequency) \ 4
                            Case Else
                                Frequency = 0
                        End Select
                        
                        If GetMP3Info.VBR = True Then
                           BitRate = (FileSize * Frequency) / (CLng(GetMP3Info.Frames)) / 1000 / Temp(MpegLayer)
                           Ok = True
                        Else
                           LayerVersioN = MpegVersion & MpegLayer
                           Select Case Val(LayerVersioN)
                              Case 11
                                 BRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448, -1)
                              Case 12
                                 BRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384, -1)
                              Case 13
                                 BRate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, -1)
                              Case 21, 251
                                 BRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256, -1)
                              Case 22, 252, 23, 253
                                 BRate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, -1)
                           End Select
                           If isVector(BRate) Then
                                BitRate = BRate(BinToDec(Mid$(BIN, 17, 4)))
                           Else
                                BitRate = 0
                           End If
                        End If
                        
                        If ((OLDSamplerate = Frequency) And (OLDBitrate = BitRate)) _
                            Or Volte = 0 Then
                            Volte = Volte + 1
                            OLDSamplerate = Frequency
                            OLDBitrate = BitRate
                        ElseIf Frequency > 0 And BitRate > 0 Then
                            Volte = 0
                            OLDSamplerate = Frequency
                            OLDBitrate = BitRate
                        End If
                        
                        If CurrentMp3 And _
                            Data.samplingRate <> 0 And Data.BitRate <> 0 Then
                            If Data.samplingRate = Frequency And _
                               Data.BitRate = BitRate Then
                                Ok = True
                            End If
                        ElseIf (Volte >= MaxVolte) Or (StartByte = 0) Then
                            Ok = True
                        End If
                    End If
                    
                Loop Until ((BitRate > 0) And Ok) Or (StartByte >= MaxCapa)
                
                If MpegVersion > 10 Then
                    MpegVersion = MpegVersion / 10
                End If
                Original = NoYes(Mid$(BIN, 30, 1))
                CopyRight = NoYes(Mid$(BIN, 29, 1))
                Padding = NoYes(Mid$(BIN, 23, 1))
                PrivateBit = NoYes(Mid$(BIN, 24, 1))
                NoYes = Array("Yes", "No")
                CRC = NoYes(Mid$(BIN, 16, 1))
                Mode = SMode(BinToDec(Mid$(BIN, 25, 2)))
                Emphasis = Emph(BinToDec(Mid$(BIN, 31, 2)))
                Sinc = BinToDec(Mid$(BIN, 1, 11))
                
                Erase X
                
                               
                If BitRate < 0 Or StartByte >= MaxCapa Then
                    WriteINFO "Information Not Found"
                    Call AzzeraStrutturaMp3Info(GetMP3Info)
                Else
                
                    WriteINFO "Header Found, Scan " & StartByte & " byte"
                    If ReadTAG Then
                        Call GetTAG(FileName)
                    End If
                    
                    MS = (FileLen(FileName) * 8) / (1024 * BitRate) 'calculate duration
                    Duration = Int(MS)
                    
                    With GetMP3Info 'set values
                       .BitRate = BitRate
                       .CRC = CRC
                       .Duration = Duration
                       .Emphasis = Emphasis
                       .Frequency = Frequency
                       .Mode = Mode
                       .MpegLayer = MpegLayer
                       .MpegVersion = MpegVersion
                       .Padding = Padding
                       .Original = Original
                       .CopyRight = CopyRight
                       .PrivateBit = PrivateBit
                    End With

                End If
        End Select
    End If
    ReadMP3 = GetMP3Info

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modInfoMp3.ReadMP3" & vbCrLf & Err.Description)
        Call AzzeraStrutturaMp3Info(GetMP3Info)
    End If
    gb_InfoMp3Active = False
End Function

Public Function WriteTag(FileName As String, SongName As String, _
   Artist As String, Album As String, Year As String, Comment As String, Genre As Integer) As Long
   Dim SN As String * 30
   Dim Com As String * 30
   Dim Art As String * 30
   Dim Alb As String * 30
   Dim Yr As String * 4
   Dim Gr As String * 1
   Const Tag = "TAG"
On Error GoTo ErrH
    If Right$(UCase(FileName), 4) <> ".WAV" Then
       SN = SongName
       Com = Comment
       Art = Artist
       Alb = Album
       Yr = Year
       If Genre > -1 Then
        Gr = Chr$(Genre)
       End If
       Open FileName For Binary Access Write As #1
            Seek #1, FileLen(FileName) - 127
            Put #1, , Tag
            Put #1, , SN
            Put #1, , Art
            Put #1, , Alb
            Put #1, , Yr
            Put #1, , Com
            Put #1, , Gr
       Close #1
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modInfoMp3.WriteTag" & vbCrLf & Err.Description)
    End If
End Function

Public Function GenreText(Index As Integer) As String
    Dim Matrix() As Variant
On Error GoTo ErrH
    If Index = 255 Then
        GenreText = "Nope"
    Else
        Matrix = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
           "Hip -Hop", "Jazz", "Metal", "new Age", "Oldies", "Other", "Pop", "R&b", "Rap", "Reggae", _
           "Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
           "Soundtrack", "Euro -Techno", "Ambient", "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", _
           "Trance", "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", _
           "noise", "AlternRock", "bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
           "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno -Industrial", "Electronic", _
           "Pop -Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
           "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "new Wave", _
           "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo -Fi", "Tribal", "Acid Punk", "Acid Jazz", _
           "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "national Folk", _
           "Swing", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
           "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
           "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty bass", _
           "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
           "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
           "Euro - House", "Dance Hall", "Goa", "Drum & bass", "Club - House", "Hardcore", "Terror", "Indie", "BritPop", _
           "negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
           "Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Trash Metal", "Anime", "JPop", "SynthPop")
        GenreText = Matrix(Index)
    End If
ErrH:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

Private Sub AzzeraStrutturaMp3Info(ByRef Info As Mp3Info)
    With Info
       .BitRate = 0
       .Frequency = 0
       .Mode = ""
       .Emphasis = ""
       .MpegVersion = 0
       .MpegLayer = 0
       .Padding = ""
       .CRC = ""
       .Duration = 0
       .CopyRight = ""
       .Original = ""
       .PrivateBit = ""
       .HasTag = False
       .Tag = ""
       .SongName = ""
       .Artist = ""
       .Album = ""
       .Year = ""
       .Comment = ""
       .Genre = 0
       .track = ""
       .VBR = False
       .Frames = 0
    End With
End Sub

Private Function FillFormat(Music_FileName As String) As WAVEFORMATEX

    Dim By As Byte
    Dim I As Long
    Dim Header As FileHeader
    Dim HdrFormat As FileFormat
    Dim FileFree As Integer
On Error GoTo ErrH


    FileFree = FreeFile
    Open Music_FileName For Binary Access Read As #FileFree
                        
        Get #FileFree, , Header
        If Header.dwRiff <> &H46464952 Then
            Exit Function
        End If
        If Header.dwWave <> &H45564157 Then
            Exit Function
        End If
        Dim lCount As Long
        
        If Header.dwFormatLength < 16 Then
            Exit Function
        End If
        
        Get #FileFree, , HdrFormat
                          
        For I = 1 To Header.dwFormatLength - 16
            Get #FileFree, , By
        Next
                  
        With FillFormat
            .nAvgBytesPerSec = HdrFormat.nAvgBytesPerSec
            .nSamplesPerSec = HdrFormat.nSamplesPerSec
            .wBitsPerSample = HdrFormat.wBitsPerSample
            .nBlockAlign = HdrFormat.nBlockAlign
            .nChannels = HdrFormat.nChannels
            .wFormatTag = HdrFormat.wFormatTag
        End With
    
    Close #FileFree
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modInfoMp3.FillFormat" & vbCrLf & Err.Description)
    End If
End Function
