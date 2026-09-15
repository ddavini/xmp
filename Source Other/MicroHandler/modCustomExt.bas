Attribute VB_Name = "modCustomExt"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_CHAR = &H102
Private cPlugIN As Object
Private PlugINCreated As Boolean
Private StartData As String
Private NumeroMp3 As Integer

Private Type Commands
    Action As String
    AppName As String
    SessionID As Long
    Data As String
    DataType As String
    Done As Boolean
End Type

Private Const TimeOut As Integer = 120

Public Function CustomExt(Optional StrC As String _
                        , Optional StrQ As String _
                        , Optional GetFile As String _
                        , Optional SockHnd As Long) As String

    Dim tCommand As Commands
    Dim lSessionID As Long
    Dim sKey As String
    Dim cEntropy As New clsECG
    Dim cMD5 As New clsMD5
    Dim NumeroFile As Long
    Dim sTMP As String
On Error GoTo Error
    
    
    If StartData = "" Then
        StartData = GetIni(ConfigFileName, "MAIN", "XMPSTART", "#")
        If StartData = "#" Then
            StartData = GetIni(ConfigFileName, "MAIN", "SEPATH", "#") & "\" & GetIni(ConfigFileName, "MAIN", "ENGINEEXE", "#")
        End If
    End If
    
  '  If Not PlugINCreated Then
    If IsNothing(cPlugIN) Then
        Set cPlugIN = CreateObject(GetIni(ConfigFileName, "MAIN", "ENTRYPLUG", ""))
        PlugINCreated = True
        'hMMTimer = timeSetEvent(6000, 0, AddressOf TimerProc, 0, TIME_ONESHOT Or TIME_CALLBACK_FUNCTION)
    End If
    
    GetFile = RootDir & GetFile
    NumeroFile = FreeFile
    Open GetFile For Binary Access Read Shared As NumeroFile
        CustomExt = String(LOF(NumeroFile), 1)
        Get #NumeroFile, 1, CustomExt
    Close NumeroFile
    
    If StrQ <> "" Then
        tCommand = ParseQueryString(StrQ)
        
        With tCommand
        
            .AppName = Replace$(.AppName, "[enginename]", GetIni(ConfigFileName, "MAIN", "ENGINENAME", "#"), , , vbTextCompare)
            .AppName = Replace$(.AppName, "[engineexe]", GetIni(ConfigFileName, "MAIN", "ENGINEEXE", "#"), , , vbTextCompare)
            
            .Data = Replace$(.Data, "[mp3root]", GetIni(ConfigFileName, "MAIN", "MP3ROOT", "#"), , , vbTextCompare)
            .Data = Replace$(.Data, "[selogroot]", GetIni(ConfigFileName, "MAIN", "SELOGROOT", "#"), , , vbTextCompare)
            .Data = Replace$(.Data, "[sepath]", GetIni(ConfigFileName, "MAIN", "SEPATH", "#"), , , vbTextCompare)
            
            Select Case LCase$(.Action)
                Case Is = "login"
                    'RND \ socketID
                    lSessionID = cEntropy.RandomNumber(0, 2147483646) \ SockHnd
                    'MD5(RND * sessionid)
                    sKey = LCase$(cMD5.DigestStrToHexStr(CStr(cEntropy.RandomNumber(0, 2147483646) * lSessionID)))
                    Call cPlugIN.login(lSessionID, .AppName, TimeOut, sKey)
                    CustomExt = "sessionid=" & CStr(lSessionID) & vbCrLf & _
                                "sessiontimeout=" & CStr(TimeOut) & vbCrLf & _
                                "key=" & sKey & vbCrLf & _
                                "response=" & cPlugIN.response.sresponse & vbCrLf & _
                                "data=" & CStr(cPlugIN.response.lResponse) & vbCrLf
                Case Is = "senddata"
                    Call cPlugIN.SendData(.SessionID, .Data, .DataType)
                    CustomExt = "response=" & cPlugIN.response.sresponse & vbCrLf & _
                                "data=" & CStr(cPlugIN.response.lResponse) & vbCrLf
                Case Is = "sendcmd"
                    Call cPlugIN.sendcommand(.SessionID, .Data, .DataType)
                    CustomExt = "response=" & cPlugIN.response.sresponse & vbCrLf & _
                                "data=" & CStr(cPlugIN.response.lResponse) & vbCrLf
                Case Is = "logout"
                    Call cPlugIN.logout(.SessionID)
                    CustomExt = "response=" & cPlugIN.response.sresponse & vbCrLf & _
                                "data=" & CStr(cPlugIN.response.lResponse) & vbCrLf
                                
                '''XmP Server Command'''
                Case Is = "app_start"
                    StartData = UnicodeTrans(.Data) & "\" & .AppName
                    
                    CustomExt = Shell(StartData, vbNormalFocus)
                Case Is = "app_sendkey"
                    CustomExt = SendCmd(.AppName, .Data)
                Case Is = "app_readls"
                    CustomExt = _
                    Replace$(CustomExt, "[readls]", ReadMp3List(UnicodeTrans(.Data)), _
                    1, -1, vbTextCompare)
                Case Is = "app_writels"
                    CustomExt = SendCmd(.AppName, "k")
                    CustomExt = CustomExt & " - " & CollectMP3(UnicodeTrans(.Data), UnicodeTrans(.AppName))
                Case Is = "app_readtxt"
                    sTMP = Replace$(CStr(Date), "/", ".")
                    .Data = Replace$(.Data, "[date]", sTMP, 1, -1, vbTextCompare)
                    CustomExt = _
                    Replace$(CustomExt, "[txt]", ReadTXT(UnicodeTrans(.Data)), _
                    1, -1, vbTextCompare)
                Case Else
                    Err.Raise 666, Description:="Unknown Action"
            End Select
        End With
    End If
    Exit Function
Error:
    CustomExt = Err.Description
    Set cPlugIN = Nothing
    PlugINCreated = False
End Function

Private Function ParseQueryString(ByVal StrQ As String) As Commands
    Dim sArray() As String
    Dim sCommand As String
    Dim sCmdData As String
    Dim I As Byte
    
    sArray = Split(StrQ, "&")
    For I = 0 To UBound(sArray)
        sCommand = LCase$(Left$(sArray(I), InStr(1, sArray(I), "=", vbTextCompare) - 1))
        sCmdData = Right$(sArray(I), Len(sArray(I)) - Len(sCommand) - 1)
        With ParseQueryString
            .Done = True
            Select Case sCommand
                Case Is = "action"
                    .Action = sCmdData
                Case Is = "appname"
                    .AppName = sCmdData
                Case Is = "sessionid"
                    .SessionID = sCmdData
                Case Is = "data"
                    .Data = sCmdData
                Case Is = "datatype"
                    .DataType = sCmdData
                Case Else
                    .Done = False
            End Select
        End With
    Next I
End Function

Private Function IsNothing(Obj As Object) As Boolean
On Error GoTo Error

    Obj.Nothing
  
Error:
    If Err.Number = 91 Then
        IsNothing = True
    Else
        IsNothing = False
    End If
    Err.Clear
End Function

Private Function SendCmd(sWindow As String, sKey As String) As String
    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, sWindow)
    
    If hwnd <> 0 Then
      '  Call SetForegroundWindow(hwnd)
      '  Call SetFocus(hwnd)
      '  Call SetActiveWindow(hwnd)
        SendCmd = SendMessage(hwnd, WM_CHAR, Asc(sKey), 0)
    Else
        SendCmd = "XmP Engine not found!"
        If StartData <> "" Then
            SendCmd = "Auto Start Engine..." & Shell(StartData, vbNormalFocus)
        End If
    End If
    
End Function

Public Function ReadTXT(ByVal Filenamez As String) As String
    Dim Filez As String
    Dim FileNumber As Integer
    Dim Str As String

    Filez = Filenamez
    FileNumber = FreeFile
    ReadTXT = "."
    If Dir(Filez) <> "" Then
        Open Filez For Input Shared As #FileNumber
            While Not EOF(FileNumber)
                Line Input #FileNumber, Str
                ReadTXT = ReadTXT & Str & vbCrLf
            Wend
        Close #FileNumber
    End If
    Call EraseC34(ReadTXT)

End Function

Public Function ReadMp3List(ByVal Filenamez As String) As String
    Dim I  As Integer
    Dim Filez As String
    Dim FileNumber As Integer
    Dim LstIndex As Integer
    Dim strListaMp3 As String
    Dim Mp3FileTmp As String

    Filez = Filenamez
    FileNumber = FreeFile
    ReadMp3List = "Nope"
    If Dir(Filez) <> "" Then
        Open Filez For Input Shared As #FileNumber
            If Not EOF(FileNumber) Then
                Input #FileNumber, LstIndex
                While Not EOF(FileNumber) And I <= LstIndex
                    Line Input #FileNumber, Mp3FileTmp
                    Line Input #FileNumber, strListaMp3
                    I = I + 1
                Wend
                ReadMp3List = strListaMp3
            End If
        Close #FileNumber
    End If
    Call EraseC34(ReadMp3List)

End Function

Public Function CollectMP3(ByVal path As String, ByVal LSFile As String) As String

    Dim sDir As String
    Dim I As Integer
    Dim sDirArray() As String
    Dim sDirUser() As String
 
    Dim sExt As String
    Dim vExt() As String
       
    sExt = GetIni(App.path & "\config.ini", "MAIN", "EXT", "*.mp3")
    
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
    
    CollectMP3 = "OK"
    
End Function

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

