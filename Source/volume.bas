Attribute VB_Name = "modVisualMusic"
Option Explicit
 
Private Const MERGEPAINT = &HBB0226
Private Const SRCINVERT = &H660046
Private Const SRCERASE = &H440328
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const WHITENESS = &HFF0062
Private Const BLACKNESS = &H42

Private IndiceOsc As Integer

Public Enum SpectrumeMode
    PeakFalls = 0
    noPeakFalls = 1
    PeaknoFalls = 2
    noPeaknoFalls = 3
    OSC = 4
    StereoOSC = 5
End Enum

Private SpectrumModeAttivo As SpectrumeMode

Public Const AmpiezzaSpectrumDisplay = 100
Public Const AltezzaSpectrumDisplay = 23

Private Const Altezza = AltezzaSpectrumDisplay

Private NBarre As Integer
Private W() As Single
Private P() As Single

Private TM As Long
Private Spec(255) As Long
Private Spec2(255) As Long

Public Function GetxmFFTVersion() As String
    Dim xmFFTVer As VersionInfo
    
    xmFFTVer = DisplayVerInfo(App.Path & "\xmFFT.dll")
    
    With xmFFTVer
        GetxmFFTVersion = .Major & "." & .Minor & "." & .Revision
    End With
    
    
End Function

Private Sub DisegaSpectrum(Mode As SpectrumeMode)
        Dim I As Long
                
        If SpectrumModeAttivo <> Mode Then
            xmp.analyzer.AutoRedraw = True
            xmp.analyzer.Cls
            xmp.analyzer.AutoRedraw = False
            xmp.abuff.Cls
        End If
        SpectrumModeAttivo = Mode
        
        Select Case Mode
            Case Is = SpectrumeMode.PeakFalls
                xmp.analyzer.ToolTipText = "xmMP3 Stone Peak Falls"
                Call DisgnaxmFFT
            Case Is = noPeakFalls
                xmp.analyzer.ToolTipText = "xmMP3 Stone Falls"
                Call DisgnaxmFFT(False, True)
            Case Is = PeaknoFalls
                xmp.analyzer.ToolTipText = "xmMP3 Stone Peak"
                Call DisgnaxmFFT(True, False)
            Case Is = noPeaknoFalls
                xmp.analyzer.ToolTipText = "xmMP3 Fade FFT"
                Call DisgnaxmMP3FFT
            Case Is = OSC
                xmp.analyzer.ToolTipText = "Osc"
                ShowMovingLine
            Case Is = StereoOSC
                xmp.analyzer.ToolTipText = "Stereo Osc"
                ShowMovingLineStereo
        End Select
End Sub

Public Sub InitFFT()
    
    WriteINFO "Init FFT..."
    
    NBarre = (AmpiezzaSpectrumDisplay / Gap) - 1
    ReDim W(NBarre)
    ReDim P(NBarre)
    StepRate = GetINI(cfgFile, "VISUALIZATION", "STONESTEP", 1, 1, 10)
    
    If NBarre * StepRate > 255 Then
        StepRate = 1
    End If
    
    'rectangle = 0
    'hanning = 1
    'hamming = 2
    'blackman = 3
    SpecMode = GetINI(cfgFile, "VISUALIZATION", "STONEMODE", 0, 0, 3)
End Sub

Private Sub BitBltSpec(DatiFFT() As Single, _
                       Optional NoOfLinesStart As Long = 0, _
                       Optional bPeak As Boolean = True, _
                       Optional bFall As Boolean = True)
On Error GoTo ErrH:
Dim I As Long
Dim pY As Single
Dim Y As Single
Dim D As Single
Dim XX As Byte

'Const Gap = 2
'Const Barwidth = 1
'Const PointFalls = 5
'Const Tolleranza = 1
'Const FallsVel = 0.3
'Const noFallsVel = 0.5

Const PointY = Altezza
  
  For I = 0 To NBarre
    'Point Generator
    D = DatiFFT(I + NoOfLinesStart)
    D = D * Altezza
    D = (Altezza - D) \ 1
    
    'Bar Falling Code
    If bFall Then
        If D > W(I) Then
            Y = W(I) + FallsVel
        Else
            Y = D
        End If
        If Y <= 0 Then
            Y = 1
        End If
    Else
        If D > W(I) Then
            Y = W(I) + noFallsVel
        Else
            Y = D
        End If
        If Y <= 0 Then
            Y = 1
        End If
    End If
    If Y >= PointY - Tolleranza Then
        Y = PointY - Tolleranza
    End If



    'Peaks Falling Code
    If bPeak Then
        If P(I) > Y Then
            pY = Y - PointFalls
        Else
            pY = P(I) + 0.05
        End If
        If pY <= 0 Then
            pY = 1
        End If
        If pY >= PointY - Tolleranza Then
            pY = PointY - Tolleranza
        End If
    End If
    
    XX = (I + (2 / Gap)) * Gap
    'clear the rest of area than x
    Call BitBlt(xmp.analyzer.hdc, XX, 0, Barwidth, Y, xmp.abuff.hdc, XX, 0, SRCCOPY)
    '''
    'update Bar
    Call BitBlt(xmp.analyzer.hdc, XX, Y - 3, Barwidth, PointY - Y, xmp.gph.hdc, 0, Y, SRCCOPY)
    '''
    'Plots The Points  if point falls <> 0
    If bPeak Then
        Call BitBlt(xmp.analyzer.hdc, XX, pY - 3, Barwidth, 1, xmp.gpeak.hdc, 0, 0, SRCCOPY)
    End If
    '''
    'Store the last Value
    W(I) = Y
    P(I) = pY
    '''
Next I

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modVisualMusic.BitBltSpec" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub DisegnaCardioOSC(ByVal LeftVol As Long, ByVal RightVol As Integer, Optional Azzera As Boolean)
On Error GoTo ErrH
    Dim Valore As Long
    Dim SinVH As Long
    Dim DesVH As Long
    Dim TMax As Byte
    
    TMax = 100
    
    If Azzera = False Then
        
        If Modo = TipoModo.Mono Then
            RightVol = LeftVol
        End If
        
        SinVH = xmp.picCardioSin.Height - (Screen.TwipsPerPixelY * 3)
        DesVH = xmp.picCardioDes.Height - (Screen.TwipsPerPixelY * 3)
        Valore = (SinVH / TMax) * LeftVol
        xmp.picCardioSin.PSet (IndiceOsc + (Screen.TwipsPerPixelX * 3), _
                                SinVH - Valore), vbGreen
        Valore = (DesVH / TMax) * RightVol
        xmp.picCardioDes.PSet (IndiceOsc + (Screen.TwipsPerPixelX * 3), _
                                DesVH - Valore), vbGreen
                                
        Call xmp.SpectrumDes.WriteSpec(CLng(RightVol) + 1, TMax)
        Call xmp.SpectrumSin.WriteSpec(CLng(LeftVol) + 1, TMax)
        
        IndiceOsc = IndiceOsc + 1
        If IndiceOsc > xmp.picCardioSin.Width - 1 Then
            xmp.picCardioSin.Cls
            xmp.picCardioDes.Cls
            IndiceOsc = 0
        End If
    Else
        IndiceOsc = 0
        Call xmp.SpectrumDes.WriteSpec(1, TMax)
        Call xmp.SpectrumSin.WriteSpec(1, TMax)
        Call xmp.picCardioSin.Cls
        Call xmp.picCardioDes.Cls
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modVisualMusic.DisegnaCardioOSC" & vbCrLf & Err.Description)
    End If
End Sub

Public Sub SetCardioOSCVisibile(Optional Visibile As Boolean)
On Error GoTo ErrH
    xmp.SpectrumDes.Visible = Visibile
    xmp.SpectrumSin.Visible = Visibile
    xmp.picCardioSin.Visible = Visibile
    xmp.picCardioDes.Visible = Visibile
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modVisualMusic.SetCardioOSCVisibile" & vbCrLf & Err.Description)
    End If
End Sub

Public Function GetCardioOSCVisibile() As Boolean
On Error GoTo ErrH
    GetCardioOSCVisibile = xmp.SpectrumDes.Visible And _
                        xmp.SpectrumSin.Visible And _
                        xmp.picCardioSin.Visible And _
                        xmp.picCardioDes.Visible
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modVisualMusic.GetCardioOSCVisibile" & vbCrLf & Err.Description)
    End If
End Function


Public Sub DisplayAll(Optional ModoSpec As SpectrumeMode = PeakFalls)
On Error GoTo ErrH
            If mStreamIsActive Then
              If xmp.Visible Then
                    If xmp.SpectrumDes.Visible Then
                        If frmVisualizzazioniAttivo Then
                            'Non faccio l'unload per via del mouse wheel
                            'nel mainloop delle visualizzazioni
                            'Unload frmVisualizzazioni
                        End If
                        Call DisegnaCardioOSC(mSinLevel, mDesLevel)
                    End If

                    If xmp.analyzer.Visible Then
                        Call DisegaSpectrum(ModoSpec)
                    End If

                End If
            Else
                If xmp.Visible Then
                    If xmp.SpectrumDes.Visible Then
                        Call DisegnaCardioOSC(0, 0, True)
                    End If
    
                    If xmp.analyzer.Visible Then
                        'xmp.analyzer.Cls
                        Call DisegaSpectrum(ModoSpec)
                    End If
                End If
            End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                 "modVisualMusic.DisplayAll" & vbCrLf & Err.Description)
    End If
End Sub

Private Function DisgnaxmMP3FFT()
On Error GoTo ErrH:
    Dim SpecTblL(255) As Long
    Dim SpecTblR(255) As Long
   
    Dim I As Long
    Dim J As Long
    Dim Y As Long
    Dim XX As Long
    Dim YY As Long
    Dim GG As Integer
    Dim HH As Long
    Dim Handler As Long

    
        
        If frmVisualizzazioniAttivo Then
            'Non faccio l'unload per via del mouse wheel
            'nel mainloop delle visualizzazioni
            'Unload frmVisualizzazioni
        End If
        xmp.analyzer.AutoRedraw = True
        xmp.analyzer.Cls
        
        If mStreamIsActive Then
            Call mFFT(SpecTblL, SpecTblR, fft_window.rectangle)
            
            HH = xmp.analyzer.ScaleHeight
            Handler = xmp.analyzer.hdc
            For I = 0 To xmp.analyzer.ScaleWidth
                If TM < SpecTblL(I) Then
                    TM = SpecTblL(I)
                End If
                If TM < SpecTblR(I) Then
                    TM = SpecTblR(I)
                End If
                
                If SpecTblL(I) > SpecTblR(I) Then
                    Y = (HH / 70) * (SpecTblL(I) - 40)
                Else
                    Y = (HH / 70) * (SpecTblR(I) - 40)
                End If
                
                If Spec(I) < Y Then
                    Spec(I) = Y
                Else
                    Spec(I) = Spec(I) - 100
                    If Spec(I) < 0 Then
                        Spec(I) = 0
                    End If
                    If Spec(I) < Y Then
                        Spec(I) = Y
                    End If
                    Y = Spec(I)
                End If
                If Spec2(I) < Y Then
                    Spec2(I) = Y
                    YY = Spec2(I)
                Else
                    Spec2(I) = Spec2(I) - 30
                    If Spec2(I) < 0 Then
                        Spec2(I) = 0
                    End If
                    If Spec2(I) < Y Then
                        Spec2(I) = Y
                    End If
                    YY = Spec2(I)
                End If


                Y = HH - Y - 3
                YY = HH - YY - 3


                For J = HH - 3 To Y Step -1
                    GG = J * 255 / HH
                    GG = GG * 2 - 255
                    If GG > 255 Then GG = 255
                    If GG < 0 Then GG = 0

                    Call SetPixel(Handler, I + 2, J, RGB(64, GG, 0))
                Next J
                Call SetPixel(Handler, I + 2, YY, vbGreen)
            Next I
        Else
            YY = xmp.analyzer.ScaleHeight - 3
            xmp.analyzer.Line (2, YY)-(xmp.analyzer.ScaleWidth, YY), vbGreen
        End If
        xmp.analyzer.AutoRedraw = False
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                 "modVisualMusic.DisegnaMP3FFT" & vbCrLf & Err.Description)
    End If
End Function

Private Function DisgnaxmFFT(Optional Peaks As Boolean = True, _
                               Optional Falls As Boolean = True)
On Error Resume Next
    Dim L(255) As Long
    Dim R(255) As Long
    Dim Buffer(255) As Single
   
    Dim I As Integer
    Dim J As Integer

    If frmVisualizzazioniAttivo = False Then
        Load frmVisualizzazioni
    End If
            
    If mStreamIsActive Then
        Call mFFT(L, R, SpecMode)

        For I = 0 To NBarre * StepRate Step StepRate
            Buffer(J) = (L(I) + R(I)) / 255
            J = J + 1
        Next
    End If
    
    Call BitBltSpec(Buffer, 0, Peaks, Falls)

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                 "modVisualMusic.DisegnaxmFFT" & vbCrLf & Err.Description)
    End If
    Err.Clear
End Function


Private Function ShowMovingLine()
On Error GoTo ErrH
    Dim waveTblL(255) As Long
    Dim waveTblR(255) As Long
    
    Dim I As Long

    Dim XX As Long
    Dim YY As Long
    
        If frmVisualizzazioniAttivo Then
            'Non faccio l'unload per via del mouse wheel
            'nel mainloop delle visualizzazioni
            'Unload frmVisualizzazioni
        End If
        
        xmp.analyzer.AutoRedraw = True
        Call xmMP3_getWave(waveTblL(0), waveTblR(0))
        xmp.analyzer.Cls
        
        If mStreamIsActive Then
            waveTblL(0) = (waveTblL(0) + waveTblR(0)) / 2 + 100
            XX = 0
            YY = (xmp.analyzer.ScaleHeight / 200) * waveTblL(0)
            xmp.analyzer.Line (XX, YY)-(XX, YY), vbGreen
            
            For I = 2 To 255 Step 2
                waveTblL(I) = (waveTblL(I) + waveTblR(I)) / 2 + 100
                XX = (xmp.analyzer.ScaleWidth / 256) * I
                YY = (xmp.analyzer.ScaleHeight / 200) * waveTblL(I)
                xmp.analyzer.Line -(XX, YY), vbGreen
            Next I
        Else
            YY = xmp.analyzer.ScaleHeight / 2
            xmp.analyzer.Line (0, YY)-(xmp.analyzer.ScaleWidth, YY), vbGreen
        End If
        xmp.analyzer.AutoRedraw = False
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                 "modVisualMusic.ShowMovingLine" & vbCrLf & Err.Description)
    End If
End Function

Private Function ShowMovingLineStereo()
On Error GoTo ErrH
    Dim m_waveTblL(255) As Long
    Dim m_waveTblR(255) As Long
    
    Dim I As Long
    Dim Y As Long
    Dim XX As Long
    Dim YY As Long
    Dim waveData As WAVE_DATA
    
        If frmVisualizzazioniAttivo Then
            'Non faccio l'unload per via del mouse wheel
            'nel mainloop delle visualizzazioni
            'Unload frmVisualizzazioni
        End If

        
        xmp.analyzer.AutoRedraw = True
        Call xmMP3_getWave(m_waveTblL(0), m_waveTblR(0))
        xmp.analyzer.Cls
        
        If mStreamIsActive Then
            'Left
            m_waveTblL(0) = m_waveTblL(0) + 100
            XX = 0
            YY = (xmp.analyzer.ScaleHeight / 200) * m_waveTblL(0)
            xmp.analyzer.Line (XX, YY)-(XX, YY), vbGreen

            For I = 2 To 127 Step 2
                m_waveTblL(I) = m_waveTblL(I) + 100
                XX = (xmp.analyzer.ScaleWidth / 256) * I
                YY = (xmp.analyzer.ScaleHeight / 200) * m_waveTblL(I)
                xmp.analyzer.Line -(XX, YY), vbGreen
            Next I
            
            'Right
            m_waveTblR(0) = m_waveTblR(0) + 100
            XX = (xmp.analyzer.ScaleWidth / 2) + 1
            YY = (xmp.analyzer.ScaleHeight / 200) * m_waveTblR(0)
            xmp.analyzer.Line (XX, YY)-(XX, YY), vbGreen

            For I = 4 To 127 Step 2
                m_waveTblR(I) = (m_waveTblR(I)) + 100
                XX = ((xmp.analyzer.ScaleWidth / 256) * I) + (xmp.analyzer.ScaleWidth / 2)
                YY = (xmp.analyzer.ScaleHeight / 200) * m_waveTblR(I)
                xmp.analyzer.Line -(XX, YY), vbGreen
            Next I
            
        Else
            XX = xmp.analyzer.ScaleWidth / 2
            YY = xmp.analyzer.ScaleHeight / 2
            xmp.analyzer.Line (0, YY)-(XX, YY), vbGreen
            xmp.analyzer.Line (XX + 1, YY)-(xmp.analyzer.ScaleWidth, YY), vbGreen
        End If
        xmp.analyzer.AutoRedraw = False
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & vbCrLf & _
                 "modVisualMusic.ShowMovingLine" & vbCrLf & Err.Description)
    End If
End Function

