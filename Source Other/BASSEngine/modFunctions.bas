Attribute VB_Name = "modFunctions"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef sorg As Any, ByVal noByte As Long)

Public LSIndex As Long
Public LSRestart As Boolean
Public NumeroMp3 As Long
Public vRND() As Boolean

Public Function ReadMp3Last(ByVal Filenamez As String) As Long
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim LstIndex As Integer
    Dim strListaMp3 As String
    Dim Mp3FileTmp As String

    Filez = Filenamez
    FileNumber = FreeFile
    ReadMp3Last = 0
    If Dir(Filez) <> "" Then
        Open Filez For Input Shared As #FileNumber
            Input #FileNumber, LstIndex
            ReadMp3Last = LstIndex
            If Not EOF(FileNumber) Then
                While Not EOF(FileNumber) And I <= LstIndex
                    Line Input #FileNumber, Mp3FileTmp
                    Line Input #FileNumber, strListaMp3
                    I = I + 1
                Wend
                LstIndex = I
                frmBassEngine.SONG.Caption = strListaMp3
            End If
        Close #FileNumber
    End If

End Function

Public Function ReadMp3Number(ByVal Filenamez As String) As Long
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim LstIndex As Integer
    Dim strListaMp3 As String
    Dim Mp3FileTmp As String

    Filez = Filenamez
    FileNumber = FreeFile
    ReadMp3Number = 0
    If Dir(Filez) <> "" Then
        Open Filez For Input Shared As #FileNumber
            Input #FileNumber, LstIndex
            If Not EOF(FileNumber) Then
                While Not EOF(FileNumber)
                    Line Input #FileNumber, Mp3FileTmp
                    Line Input #FileNumber, strListaMp3
                    I = I + 1
                Wend
                ReadMp3Number = I - 1
            End If
        Close #FileNumber
    End If

End Function

Public Function ReadMp3List(ByVal Filenamez As String, LstIndex As Long) As String
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    'Dim LstIndex As Integer
    Dim strListaMp3 As String
    Dim Mp3FileTmp As String

    Filez = Filenamez
    FileNumber = FreeFile
    ReadMp3List = "Nope"
    If Dir(Filez) <> "" Then
        Open Filez For Input Shared As #FileNumber
            Input #FileNumber, strListaMp3
            If Not EOF(FileNumber) Then
                While Not EOF(FileNumber) And I <= LstIndex
                    Line Input #FileNumber, Mp3FileTmp
                    Line Input #FileNumber, strListaMp3
                    I = I + 1
                Wend
                ReadMp3List = Mp3FileTmp
                frmBassEngine.SONG.Caption = strListaMp3
                
                If EOF(FileNumber) Then
                    LSRestart = True
                Else
                    LSRestart = False
                End If
            End If
        Close #FileNumber
    End If

End Function

Public Sub ScriviPosizioneListaMp3(Optional ByVal Filenamez As String = ".\mp3.ls")
    
    Dim Filez As String
    Dim FileNumber As Integer
    Dim sStore As String
    Dim pos As Integer
    
    Dim xByte() As Byte
    
On Error GoTo ErrH
        Filez = Filenamez & ".$$$"
        FileNumber = FreeFile
        ChDrive App.path
        ChDir App.path
        
        
        ReDim xByte(FileLen(Filenamez))
        
        Open Filenamez For Binary Access Read As #FileNumber
            Get #FileNumber, , xByte
        Close #FileNumber
        
        sStore = Space$(FileLen(Filenamez))
        
        Call CopyMemory(ByVal sStore, xByte(0), FileLen(Filenamez))
        
        pos = InStr(1, sStore, vbCrLf)
        
        sStore = Right$(sStore, Len(sStore) - pos - 1)
        
        sStore = Left$(sStore, Len(sStore) - 2)
        
        Open Filez For Output As #FileNumber
            Print #FileNumber, LSIndex
            Print #FileNumber, sStore
        Close #FileNumber

        Call FileCopy(Filez, Filenamez)
                
ErrH:
    If Err.Number <> 0 Then
        Close #FileNumber
        Err.Clear
    End If
    
    If FileExists(Filez) Then
        Call Kill(Filez)
    End If
    
End Sub

Public Function CollectMP3(ByVal path As String, ByVal LSFile As String) As String

    Dim sDir As String
    Dim I As Integer
    Dim sDirArray() As String
    Dim sDirUser() As String
 
    Dim sExt As String
    Dim vExt() As String
       
    sExt = GetIni("", "MAIN", "EXT", "*.mp3")
    
    vExt = Split(sExt, ";")
    
    NumeroMp3 = 0
    path = CapPath(path, "\")
    sDir = Dir(path, vbDirectory)
    
    ReDim sDirArray(0)
    ReDim sDirUser(0)
    sDirArray(0) = path
    sDirUser(0) = path
    I = 1
    Do
        
        If (Left(sDir, 1) <> ".") Then
             If ((GetAttr(path & sDir)) And vbDirectory) <> 0 Then
                ReDim Preserve sDirArray(I)
                ReDim Preserve sDirUser(I)
                sDirArray(I) = CapPath(path) & sDir
                sDirUser(I) = sDir
                I = I + 1
            End If
        End If
        
        sDir = Dir
    Loop Until sDir = ""
    
    I = 0
    
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open LSFile For Output As #FileNumber
        Print #FileNumber, 0
        Do
            Call WriteLS(sDirArray(I), LSFile, FileNumber, vExt)
            I = I + 1
        Loop Until I = UBound(sDirArray) + 1
    Close #FileNumber
    
    If CBool(GetIni("", "MAIN", "RANDOM", "False")) Then
        Call CreateRNDFile(NumeroMp3)
    End If
    
    LSIndex = 0
    
    CollectMP3 = "OK"
    
End Function

Public Sub WriteRNDFile(Indice As Long)
    
    Call InitRND
    
    vRND(Indice) = True
    
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open App.path & "\RND.dat" For Binary Access Write As #FileNumber
        Put #FileNumber, Indice * 2 + 1, vRND(Indice)
    Close #FileNumber

End Sub

Public Sub InitRND()
    Const EOFc As Byte = 2
    If NumeroMp3 = 0 Then
        NumeroMp3 = ReadMp3Number(App.path & "\mp3.ls")
        ReDim vRND(NumeroMp3 - 1)
    End If
    
    If Dir(App.path & "\RND.dat") <> "" Then
        'If (FileLen(App.path & "\RND.dat") - EOFc) <> ((NumeroMp3 - 1) * 2) Then
        '    Call DisplayError("RND File Wrong Size", "Warning", "InitRND")
        '    Call CreateRNDFile(NumeroMp3)
        'End If
    Else
        Call DisplayError("RND File Not Found", "Warning", "InitRND")
        Call CreateRNDFile(NumeroMp3)
    End If
End Sub

Public Sub KillRNDFile()
    If Dir(App.path & "\RND.dat") <> "" Then
        Call DisplayError("Kill RND File", "Information", "KillRNDFile")
        Call Kill(App.path & "\RND.dat")
    End If
End Sub


Public Sub CreateRNDFile(ByVal Dimension As Long)
    
    Dim FileNumber As Integer
    
    Call KillRNDFile
    
    Call DisplayError("Create RND File", "Information", "CreateRNDFile")
    
    ReDim vRND(Dimension)
    
    FileNumber = FreeFile
    Open App.path & "\RND.dat" For Binary Access Write As #FileNumber
        Put #FileNumber, , vRND
    Close #FileNumber

End Sub

Public Sub LoadRNDDat(ByVal Dimension As Long)
    
    Dim FileNumber As Integer
    
    ReDim vRND(Dimension)
    
    FileNumber = FreeFile
    Open App.path & "\RND.dat" For Binary Access Read As #FileNumber
        Get #FileNumber, , vRND
    Close #FileNumber
    
    LSIndex = 0

End Sub

Public Sub WriteLS(ByVal sDirPath As String, LSFile As String, FileNumber As Integer, vExt() As String)
On Error Resume Next
    Dim sFile As String, sDir As String, sDirs() As String
    Dim lDirCount As Integer
    'Static NumeroMp3 As Integer
    
    Dim I As Long
    
    sDirPath = CapPath(sDirPath, "\")
    ReDim sDirs(0)
    sDir = Dir(sDirPath, vbDirectory)


    While Len(sDir) > 0
        If (Left(sDir, 1) <> ".") Then
            If ((GetAttr(sDirPath & sDir)) And vbDirectory) <> 0 Then
                lDirCount = lDirCount + 1
                ReDim Preserve sDirs(lDirCount)
                sDirs(lDirCount) = CapPath(sDir, "\")
            End If
        End If
        sDir = Dir
    Wend
    
    For I = 0 To UBound(vExt)
        sFile = Dir(sDirPath & vExt(I))
        While Len(sFile) > 0
            If (Left(sDir, 1) <> ".") Then
                If (GetAttr(sDirPath & sFile) And vbDirectory) = 0 Then
                    Print #FileNumber, sDirPath & sFile & ";1:10"
                    NumeroMp3 = NumeroMp3 + 1
                    Print #FileNumber, Format(NumeroMp3, "0#") & " - " & sFile
                End If
            End If
            sFile = Dir
        Wend
    Next
    
    Debug.Print NumeroMp3
    For lDirCount = 1 To (UBound(sDirs))
        Call WriteLS(sDirPath & sDirs(lDirCount), LSFile, FileNumber, vExt)
    Next lDirCount

End Sub

Public Sub frmHotKey(KeyAscii As Integer)
  Select Case KeyAscii
    Case Asc("p") 'Paly
        frmBassEngine.cmdStreamPlay_Click
    Case Asc("s") 'Stop
        Call frmBassEngine.cmdStreamStop_Click
    Case Asc("b") 'Back
    Case Asc("n") 'Next
        Call frmBassEngine.cmdNext_Click
    Case Asc("k") 'Kill
        Unload frmBassEngine
    Case Asc("+") 'Volume +
    Case Asc("-") 'Volume -
    Case Asc("m") 'Mute
    Case Asc("f") 'Pause
        Call frmBassEngine.cmdPause_Click
    Case Asc("r") 'Reload LS
        If frmBassEngine.cmdWriteLs.Enabled Then
            frmBassEngine.cmdWriteLs_Click
        End If
  End Select
End Sub

Public Sub DisplayError(Message As String, _
                        Optional Tipo As String = "Error", _
                        Optional Dove As String = "N/A")
    Dim ErrorNum As Long
On Error GoTo ErrH
    
    ErrorNum = BASS_ErrorGetCode
       
    Call ScriviLOG(Dove, CStr(ErrorNum), Message, Tipo)
    
ErrH:
    If Err.Number <> 0 Then
        Call ScriviLOG("DisplayError", Err.Number, Err.Description, "Runtime Error")
        Unload frmBassEngine
    End If
End Sub

Public Sub ScriviLOG(ByVal where As String, ByVal info As String, ErrorDes As String, Optional ByVal Tipo As String)

    Dim I As Integer
    Dim sLog As String
    Dim sEsito As String
    Dim Esiste As String
    Dim Filez As String
    Dim FileNumber As Integer
    
On Error Resume Next
        Filez = App.path + "\Logs\"
        Esiste = Dir(Filez)
        If Esiste = "" Then
            MkDir (App.path + "\Logs\")
        End If
On Error GoTo ErrH

        sLog = CStr(time()) & vbTab
        sLog = sLog & "Type: " & Tipo & vbTab
        sLog = sLog & "Where: " & where & vbTab
        sLog = sLog & "BASSEngine_Err: " & info & vbTab
        sLog = sLog & "Details: " & ErrorDes

        Filez = App.path + "\Logs\" + CStr(Date) + ".log"
        Filez = Replace$(Filez, "/", ".")
        Esiste = Dir(Filez)
        If Esiste = "" Then
            FileNumber = FreeFile
            Open Filez For Output As #FileNumber
                Print #FileNumber, "Bass Engine Log File Created."
                Print #FileNumber, sLog
            Close #FileNumber
        Else
            FileNumber = FreeFile
            Open Filez For Append As #FileNumber
                Print #FileNumber, sLog
            Close #FileNumber
        End If

ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & "::" & Err.Description, "Runtime Error", "modFunctions.ScriviLog")
    End If
End Sub

'-------------------------
'some useful functions :)
'-------------------------

'check if any file exists
Public Function FileExists(ByVal fp As String) As Boolean
    FileExists = (Dir(fp) <> "")
End Function

' RPP = Return Proper Path
Public Function RPP(ByVal fp As String) As String
    RPP = IIf(Mid(fp, Len(fp), 1) = "\", fp, fp & "\")
End Function

'get file name from file path
Public Function GetFileName(ByVal fp As String) As String
    GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function

Public Function CapPath(ByVal sPath As String, Optional ByVal sEnd As String, _
    Optional ByVal sFile As String) As String
    If (Not IsMissing(sEnd)) And (Right(sPath, Len(sEnd)) <> sEnd) Then
        sPath = sPath & sEnd
    End If
    
    If Not IsMissing(sFile) Then
        sPath = sPath & sFile
    End If
    
    CapPath = sPath
End Function

