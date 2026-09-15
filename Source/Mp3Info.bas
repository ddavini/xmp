Attribute VB_Name = "Mp3Info"
'Option Explicit

' #VBIDEUtils#************************************************************
' * Programmer Name  : Michael Karathanasis
' * Web Site         : http://home12.inet.tele.dk/mkaratha
' * E-Mail           :
' * Date             : 6/10/99
' * Time             : 16:54
' **********************************************************************
' * Comments         : Read MP3 header
' *
' *
' **********************************************************************
'''Created By Michael Karathanasis 990728'''''''''''''''''''''
'''Now support for variable bit rate'''
'''please report improvements to http://home12.inet.tele.dk/mkaratha
'''User may use or distrubute the code without restrictions
'Sample how to use:

'Dim filename as string
'filename="c:\mymp3.mp3"
'Call ReadMP3(filename, True, True)
'Text1.Text = GetMP3Info.Bitrate
Private Type Mp3Info
   Bitrate              As Integer
   Frequency            As Long
   Mode                 As String
   Emphasis             As String
   'ModeExtension       As String
   MpegVersion          As Integer
   MpegLayer            As Integer
   Padding              As String
   CRC                  As String
   Duration             As Long
   CopyRight            As String
   Original             As String
   PrivateBit           As String
   HasTag               As Boolean
   Tag                  As String
   Songname             As String
   Artist               As String
   Album                As String
   Year                 As String
   Comment              As String
   Genre                As Integer
   Track                As String
   VBR                  As Boolean
   Frames               As Integer
End Type
Public GetMP3Info       As Mp3Info
''this function converts Binary string to decimal integer
Public Function BinToDec(BinValue As String) As Long
   BinToDec = 0
   For I = 1 To Len(BinValue)
      If Mid(BinValue, I, 1) = 1 Then
         BinToDec = BinToDec + 2 ^ (Len(BinValue) - I)
      End If
   Next I
End Function

Public Function ByteToBit(ByteArray) As String
   'convert 4*1 byte array to 4*8 bits'''''
   ByteToBit = ""
   For Z = 1 To 4
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

Public Function BinaryHeader(FileName As String, ReadTag As Boolean, ReadHeader As Boolean) As String
   Dim ByteArray(4)     As Byte
   Dim XingH            As String * 4
   FIO% = FreeFile
   Open FileName For Binary Access Read As FIO%
   n& = LOF(FIO%): If n& < 256 Then Close FIO%: Return 'ny
   If ReadHeader = False Then GoTo 5:   'if we only want to read the IDtag goto 5
   Dim x                As Byte
   '''''start check startposition for header''''''''''''
   '''''if start position <>1 then id3v2 tag exists'''''
   For I = 1 To 5000            'check up to 5000 bytes for the header
      Get #FIO%, I, x
      If x = 255 Then             'header always start with 255 followed by 250 or 251
         Get #FIO%, I + 1, x
         If x > 249 And x < 252 Then
            Headstart = I       'set header start position
            Exit For
         End If
      End If
   Next I
   '''end check start position for header'''''''''''''

   ''start check for XingHeader'''
   Get #1, Headstart + 36, XingH
   If XingH = "Xing" Then
      GetMP3Info.VBR = True
      For Z = 1 To 4 '
         Get #1, Headstart + 43 + Z, ByteArray(Z)  'get framelength to array
      Next Z
      Frames = BinToDec(ByteToBit(ByteArray))   'calculate # of frames
      GetMP3Info.Frames = Frames                'set frames
   Else: GetMP3Info.VBR = False
   End If
   '''end check for XingHeader

   '''start extract the first 4 bytes (32 bits) to an array
   For Z = 1 To 4 '
      Get #1, Headstart + Z - 1, ByteArray(Z)
   Next Z
   '''stop extract the first 4 bytes (32 bits) to an array
5:
   If ReadTag = False Then GoTo 10     'if we dont want to read the tag goto 10
   ''''start id3 tag''''''''''''''''''''''''''''''''''''''''''''''''
   Dim Inbuf            As String * 256
   Get #FIO%, (n& - 255), Inbuf:  Close FIO% 'ny
   p = InStr(1, Inbuf, "tag", 1)  'ny
   If p = 0 Then
      With GetMP3Info
         .HasTag = False
         .Songname = ""
         .Artist = ""
         .Album = ""
         .Year = ""
         .Comment = ""
         .Track = ""
         .Genre = 255
      End With
   Else
      With GetMP3Info
         .HasTag = True
         .Songname = RTrim(Mid$(Inbuf, p + 3, 30))
         .Artist = RTrim(Mid$(Inbuf, p + 33, 30))
         .Album = RTrim(Mid$(Inbuf, p + 63, 30))
         .Year = RTrim(Mid$(Inbuf, p + 93, 4))
         .Comment = RTrim(Mid$(Inbuf, p + 97, 29))
         .Track = RTrim(Mid$(Inbuf, p + 126, 1))
         .Genre = Asc(RTrim(Mid$(Inbuf, p + 127, 1)))
      End With
   End If
   ''''stop id3 tag''''''''''''''''''''''''''''''
10:
   Close FIO%
   BinaryHeader = ByteToBit(ByteArray)
End Function

Public Function ReadMP3(FileName As String, ReadTag As Boolean, ReadHeader As Boolean) As Mp3Info
   bin = BinaryHeader(FileName, ReadTag, ReadHeader)                     'extract all 32 bits

   If ReadHeader = False Then Exit Function
   Version = Array(25, 0, 2, 1)                        'Mpegversion table
   MpegVersion = Version(BinToDec(Mid(bin, 12, 2)))    'get mpegversion from table
   layer = Array(0, 3, 2, 1)                           'layer table
   MpegLayer = layer(BinToDec(Mid(bin, 14, 2)))        'get layer from table
   SMode = Array("stereo", "joint stereo", "dual channel", "single channel") 'mode table
   Mode = SMode(BinToDec(Mid(bin, 25, 2)))              'get mode from table
   Emph = Array("no", "50/15", "reserved", "CCITT J 17") 'empasis table
   Emphasis = Emph(BinToDec(Mid(bin, 31, 2)))           'get empasis from table
   Select Case MpegVersion                                 'look for version to create right table
      Case 1                                                  'for version 1
         Freq = Array(44100, 48000, 32000)
      Case 2 Or 25                                            'for version 2 or 2.5
         Freq = Array(22050, 24000, 16000)
      Case Else
         Frequency = 0
         Exit Function
   End Select
   Frequency = Freq(BinToDec(Mid(bin, 21, 2)))             'look for frequency in table
   If GetMP3Info.VBR = True Then                           'check if variable bitrate
      Temp = Array(, 12, 144, 144)                        'define to calculate correct bitrate
      Bitrate = (FileLen(FileName) * Frequency) / (Int(GetMP3Info.Frames)) / 1000 / Temp(MpegLayer)
   Else                                                 'if not variable bitrate

      Dim LayerVersion     As String
      LayerVersion = MpegVersion & MpegLayer          'combine version and layer to string
      Select Case Val(LayerVersion)                        'look for the right bitrate table
         Case 11                                              'Version 1, Layer 1
            Brate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
         Case 12                                              'V1 L1
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
         Case 13                                               'V1 L3
            Brate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
         Case 21 Or 251                                         'V2 L1 and 'V2.5 L1
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
         Case 22 Or 252 Or 23 Or 253                            ''V2 L2 and 'V2.5 L2 etc...
            Brate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
         Case Else                                               'if variable bitrate
            Bitrate = 1                                             'e.g. for Variable bitrate
            Exit Function
      End Select
      Bitrate = Brate(BinToDec(Mid(bin, 17, 4)))
   End If

   NoYes = Array("no", "yes")
   Original = NoYes(Mid(bin, 30, 1))                       'Set original bit
   CopyRight = NoYes(Mid(bin, 29, 1))                      'Set copyright bit
   Padding = NoYes(Mid(bin, 23, 1))                        'get padding bit
   PrivateBit = NoYes(Mid(bin, 24, 1))
   YesNo = Array("yes", "no")                              'CRC table
   CRC = YesNo(Mid(bin, 16, 1))                            'Get CRC
   ms = (FileLen(FileName) * 8) / Bitrate                  'calculate duration
   Duration = Int(ms / 1000)
   With GetMP3Info                                          'set values
      .Bitrate = Bitrate                                  '
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
End Function

Public Function WriteTag(FileName As String, Songname As String, _
   Artist As String, Album As String, Year As String, Comment As String, Genre As Integer) As Long
   Tag = "TAG"
   Dim sn               As String * 30
   Dim com              As String * 30
   Dim art              As String * 30
   Dim alb              As String * 30
   Dim yr               As String * 4
   Dim gr               As String * 1
   sn = Songname
   com = Comment
   art = Artist
   alb = Album
   yr = Year
   gr = Chr(Genre)
   Open FileName For Binary Access Write As #1
   Seek #1, FileLen(FileName) - 127
   Put #1, , Tag
   Put #1, , sn
   Put #1, , art
   Put #1, , alb
   Put #1, , yr
   Put #1, , com
   Put #1, , gr
   Close #1

End Function

Public Function GenreText(Index As Integer) As String
   Matrix = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
      "Hip -Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&b", "Rap", "Reggae", _
      "Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
      "Soundtrack", "Euro -Techno", "Ambient", "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", _
      "Trance", "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", _
      "Noise", "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
      "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno -Industrial", "Electronic", _
      "Pop -Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
      "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", _
      "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo -Fi", "Tribal", "Acid Punk", "Acid Jazz", _
      "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "National Folk", _
      "Swing", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
      "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
      "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", _
      "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
      "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
      "Euro - House", "Dance Hall", "Goa", "Drum & Bass", "Club - House", "Hardcore", "Terror", "Indie", "BritPop", _
      "Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
      "Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop")
   GenreText = Matrix(Index)
End Function
