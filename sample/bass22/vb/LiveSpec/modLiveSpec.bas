Attribute VB_Name = "modLiveSpec"
'/////////////////////////////////////////////////////////////////////////////////
' modLiveSpec.bas - Copyright (c) 2002-2005 (: JOBnik! :) [Arthur Aminov, ISRAEL]
'                                                         [http://www.jobnik.org]
'                                                         [  jobnik@jobnik.org  ]
' Other source: frmLiveSpec.frm
'
' BASS "Live" spectrum analyser example
' Originally translated from - livespec.c - Example of Ian Luck
'/////////////////////////////////////////////////////////////////////////////////

Option Explicit

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&    'color table in RGBs

Public Type BITMAPINFOHEADER    '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Public SPECWIDTH As Long    'display width
Public SPECHEIGHT As Long   'height (changing requires palette adjustments too)
Public specmode As Boolean, specpos As Integer  ' spectrum mode (and marker pos for 2nd mode)
Public specbuf() As Byte    'a pointer

Public chan As Long         'recording channel

Public bh As BITMAPINFO     'bitmap header

Public Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function

'update the spectrum display - the interesting bit :)
Public Sub UpdateSpectrum()
    Static quietcount As Integer
    Dim X As Long, Y As Long, Y1 As Long
    Dim fft(2048) As Single     'get the FFT data
    Call BASS_ChannelGetData(chan, fft(0), BASS_DATA_FFT4096)

    If (Not specmode) Then    '"normal" FFT
        ReDim specbuf(SPECWIDTH * (SPECHEIGHT + 1)) As Byte  'clear display
        For X = 0 To (SPECWIDTH / 2) - 1
#If 1 Then
            Y = Sqrt(fft(X + 1)) * 3 * SPECHEIGHT - 4 'scale it (sqrt to make low values more visible)
#Else
            Y = fft(X + 1) * 10 * SPECHEIGHT 'scale it (linearly)
#End If
            If (Y > SPECHEIGHT) Then Y = SPECHEIGHT 'cap it
            If (X) Then  'interpolate from previous to make the display smoother
                Y1 = (Y + Y1) / 2
                Y1 = Y1 - 1
                While (Y1 >= 0)
                    specbuf(Y1 * SPECWIDTH + X * 2 - 1) = Y1 + 1
                    Y1 = Y1 - 1
                Wend
            End If
            Y1 = Y
            Y = Y - 1
            While (Y >= 0)
                specbuf(Y * SPECWIDTH + X * 2) = Y + 1 'draw level
                Y = Y - 1
            Wend
        Next X
    Else    '"3D"
        For X = 0 To SPECHEIGHT - 1
            Y = Sqrt(fft(X + 1)) * 3 * 127 'scale it (sqrt to make low values more visible)
            If (Y > 127) Then Y = 127 'cap it
            specbuf(X * SPECWIDTH + specpos) = 128 + Y 'plot level
        Next X
        'move marker onto next position
        specpos = (specpos + 1) Mod SPECWIDTH
        For X = 0 To SPECHEIGHT - 1
            specbuf(X * SPECWIDTH + specpos) = 255
        Next X
    End If

    'update the display
    'to display in a PictureBox, simply change the .hDC to Picture1.hDC :)
    Call SetDIBitsToDevice(frmLiveSpec.hdc, 0, 0, SPECWIDTH, SPECHEIGHT, 0, 0, 0, SPECHEIGHT, specbuf(0), bh, 0)
    If (LoWord(BASS_ChannelGetLevel(chan)) < 500) Then ' check if it's quiet
        quietcount = quietcount + 1
        If (quietcount > 40 And (quietcount And 16)) Then 'it's been quiet for over a second
            Dim sNoise As String
            sNoise = "make some noise!"
            With frmLiveSpec
                .ForeColor = &HFFFFFF
                .CurrentX = (SPECWIDTH - .TextWidth(sNoise)) / 2
                .CurrentY = (SPECHEIGHT - .TextHeight(sNoise)) / 2
                frmLiveSpec.Print sNoise
            End With
        End If
    Else
        quietcount = 0 'not quiet
    End If
End Sub

'Recording callback - not doing anything with the data
Public Function DuffRecording(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Integer
    DuffRecording = BASSTRUE 'continue recording
End Function
