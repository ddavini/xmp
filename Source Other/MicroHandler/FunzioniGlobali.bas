Attribute VB_Name = "modFunzioniGlobali"
Option Explicit
Option Compare Text

Public Sub Pause(Optional APIMode As Boolean = True, Optional How As Double = 0.5)
Dim Tick As Date
    If APIMode Then
         DoEvents
         Call Sleep(CInt(How) + 1)
    Else
        Tick = Timer
        While Timer < Tick + How
            DoEvents
        Wend
    End If
End Sub

Public Function ExtractQueryString(ByRef StrC As String, Pos As Integer) As String

    ExtractQueryString = Right$(StrC, Len(StrC) - Pos)
    StrC = Left$(StrC, Pos - 1)

End Function

Public Function DivByChar(ByRef Str As String, DivChr As String) As String
Dim Pos As Integer
    Pos = InStr(1, LCase$(Str), LCase$(DivChr), vbTextCompare)
    DivByChar = Right$(Str, Len(Str) - Pos)
    Str = Left$(Str, Pos - 1)
End Function

Public Sub Shell32Bit(ByVal JobToDo As String)
        Dim hProcess As Long
        Dim RetVal As Long
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, vbHide))
        Do
            GetExitCodeProcess hProcess, RetVal
            Call Pause(How:=0)
        Loop While RetVal = STILL_ACTIVE
End Sub

Public Function ExtCtrl(StrC As String, StrQ As String, StrP As String, _
                        GetFile As String, SockHnd As Long, GZIPEnabled As Boolean) As String
Dim FilePHP As String
Dim NumeroFile As Integer
Dim StrStore As String
Dim PhpFolder As String
Dim TmpFileFolder As String
Dim STDIN As String
Dim CommandTodo As String

On Error GoTo Errore
    'Martello di Thor
    If StrQ = "" Then
        StrQ = StrP
    ElseIf StrP = "" Then
        StrP = StrQ
    End If
    
    SetEnvironmentVariable "QUERY_STRING", StrQ
    SetEnvironmentVariable "SCRIPT_FILENAME", RootDir + StrC
    SetEnvironmentVariable "PATH_TRANSLATED", RootDir + StrC
    SetEnvironmentVariable ("PATH_INFO"), StrC
    SetEnvironmentVariable ("SCRIPT_NAME"), StrC
    
    If StrP <> "" Then
        TmpFileFolder = (App.path + "\" + GetIni(ConfigFileName, "MAIN", "TMPFOLDER", "\")) & _
                        "\" & SockHnd & ".$$$"
                        
        NumeroFile = FreeFile
        Open TmpFileFolder For Output As #NumeroFile
            Print #NumeroFile, StrP
        Close #NumeroFile
        STDIN = " < " & """" & TmpFileFolder & """"
    End If

    Select Case True
        Case InStr(1, StrC, "/cgi-bin", vbTextCompare)
            
            StrStore = CurDir$
            ChDir App.path
        
            CommandTodo = Environ$("COMSPEC") & " /c " & RootDir & StrC & _
                          STDIN
            
            ExtCtrl = GetCommandOutput(CommandTodo, fStdErr:=True, fOEMConvert:=False)
            
            ChDir StrStore
        Case InStr(1, StrC, ".php", vbTextCompare)
            PhpFolder = (App.path + "\" + GetIni(ConfigFileName, "MAIN", "PHPFOLDER", "\"))
            
            StrStore = CurDir$
            ChDir PhpFolder
            ChDrive PhpFolder
            
            CommandTodo = Environ$("COMSPEC") & " /c php.exe " & """" & RootDir + StrC & """" & _
                          STDIN
            
            ExtCtrl = GetCommandOutput(CommandTodo, fStdErr:=True, fOEMConvert:=False) '& CurDir$
            
            ChDir StrStore
        Case InStr(1, StrC, "." & GetIni(ConfigFileName, "MAIN", "CUSTOMEXT", "CST"), vbTextCompare)
            ExtCtrl = CustomExt(StrC, StrQ, GetFile, SockHnd)
        Case Else
            GetFile = RootDir & GetFile
            NumeroFile = FreeFile
            Open GetFile For Binary Access Read Shared As NumeroFile
                ExtCtrl = String(LOF(NumeroFile), 1)
                Get #NumeroFile, 1, ExtCtrl
            Close NumeroFile
    End Select
    
    SetEnvironmentVariable "QUERY_STRING", "="
    SetEnvironmentVariable "CONTENT_LENGTH", "="
    SetEnvironmentVariable "CONTENT-TYPE", "="
    
    If TmpFileFolder <> "" Then
        Kill TmpFileFolder
    End If
    
Exit Function
Errore:
    Call ErrorGST(Err.Description + " - " + "ExtCtrl")
End Function

Public Sub ScriviLOG(ByVal strLog As String)
    Dim I As Integer
    Dim Esiste As String
    Dim Filez As String
    Dim FileNumber As Integer
    
On Error GoTo Errore
    Filez = LogsDir + "\" + CStr(Date) + ".log"
    Filez = Replace(Filez, "/", ".")
    Esiste = Dir(Filez)
    If Esiste = "" Then
        FileNumber = FreeFile
        Open Filez For Output As #FileNumber 'Len = 32767
            Print #FileNumber, strLog
        Close #FileNumber
    Else
        FileNumber = FreeFile
        Open Filez For Append As #FileNumber ' Len = 32767
            Print #FileNumber, strLog
        Close #FileNumber
    End If
Exit Sub
Errore:
    Call ErrorGST(Err.Description + " - " + "ScriviLog")
End Sub


Public Function RetHtmlHttpErr(ByVal HTTPerr As String, Optional URL As String) As String
Dim Corpo As String
    Select Case HTTPerr
          Case Is = "302"
                Corpo = "<head><title>Document Moved</title></head><body><h1>Object Moved</h1>This document may be found <a HREF='http://" + URL + "'>here</a></body>" & vbCrLf
          Case Is = "404"
                Corpo = "<html><head><title>Error 404</title><Corpo><h2>HTTP Error 404</h2><p><strong>404 Not Found</strong></p><p>Please check the URL to ensure that the path is correct.</p></Corpo></html>" & vbCrLf
          Case Is = "400"
                Corpo = "<html><head><title>Error 400</title><Corpo><h2>HTTP Error 400</h2><p><strong>400 Bad Request</strong></p><p>The parameter is correct.</p></Corpo></html>" & vbCrLf
          Case Else
                Corpo = ""
    End Select
    RetHtmlHttpErr = Corpo
End Function
     
Public Sub Parser(ByRef Corpo As String, ByVal StrCommand As String)
    Dim DirStr As String
    Dim Esce As String
    Dim sFSize As String
    Dim iFSize As Long
    
On Error Resume Next
    If InStr(1, StrCommand, ".log", vbTextCompare) Then
        Corpo = Replace(Corpo, vbCrLf, "<br>")
    End If
       
    If InStr(1, StrCommand, ".htm", vbTextCompare) Then
        If InStr(1, Corpo, "<!--loglist-->", vbTextCompare) Then
            DirStr = "<pre>" + vbCrLf
            Esce = Dir(LogsDir + "\*.log")
            While Esce <> ""
                DirStr = DirStr + "<a href=" + Esce + ">" + Esce + "</a>" + vbCrLf
                Esce = Dir
                DoEvents
            Wend
            If DirStr = "" Then DirStr = "Nessun Log"
            DirStr = DirStr + "</pre>" + vbCrLf
            Corpo = Replace$(Corpo, "<!--loglist-->", DirStr)
            Corpo = Replace$(Corpo, "<!--LOGLIST-->", DirStr)
         ElseIf InStr(1, Corpo, "<!--dirb-->", vbTextCompare) > 0 And _
                Right$(StrCommand, 11) _
                = "default.htm" Then
            Esce = Dir(RootDir + Left$(StrCommand, Len(StrCommand) - 11) + "*.*", vbDirectory)
            DirStr = "<pre>" + vbCrLf
            DirStr = DirStr + "<table border=0>"
            While Esce <> ""
                If LCase$(Esce) <> "default.htm" Then
                    If FileSystem.GetAttr(RootDir + _
                        Left$(StrCommand, Len(StrCommand) - 11) + "\" + Esce) And vbDirectory Then
                        sFSize = "<b>Dir</b>"
                        Esce = Esce + "/"
                    Else
                        iFSize = FileSystem.FileLen(RootDir + Left$(StrCommand, Len(StrCommand) - 11) + "\" + Esce)
                        sFSize = CStr(Round(iFSize / 1024, 2)) + " Kb"
                    End If
                    DirStr = DirStr + "<tr><td><a href='" + Esce + "'>" + Esce + "</a></td><td align=right>" + _
                    sFSize + "</td></tr>" + vbCrLf
                End If
                Esce = Dir
                DoEvents
            Wend
            DirStr = DirStr + "</table></pre>" + vbCrLf
            Corpo = Replace$(Corpo, "<!--dirb-->", vbCrLf + DirStr)
            Corpo = Replace$(Corpo, "<!--DIRB-->", vbCrLf + DirStr)
        End If
    End If
End Sub

Public Function RetFileType(ByVal Filez As String) As String
On Error Resume Next
    Select Case Mid$(Filez, InStrRev(Filez, ".", , vbTextCompare) + 1)
        Case "jpg", "jpeg"
            RetFileType = "image/jpeg"
        Case "zip"
            RetFileType = "application/x-zip-compressed"
        Case "tar"
            RetFileType = "application/x-tar"
        Case "shar"
            RetFileType = "application/x-shar"
        Case "sit"
            RetFileType = "application/x-stuffit"
        Case "pl", "perl"
            RetFileType = "application/x-perl"
        Case "gz", "gzip"
            RetFileType = "application/x-gzip"
        Case "z"
            RetFileType = "application/x-compress"
        Case "exe", "bin"
            RetFileType = "application/octet-stream"
        Case "eps", "ai", "ps"
            RetFileType = "application/postscript"
        Case "js", "mocha"
            RetFileType = "application/x-javascript"
        Case "txt", "text"
            RetFileType = "text/plain"
        Case "htm", "html", LCase$(GetIni(ConfigFileName, "MAIN", "CUSTOMEXT", "CST"))
            RetFileType = "text/html"
        Case "gif"
            RetFileType = "image/gif"
        Case "tif", "tiff"
            RetFileType = "image/tiff"
        Case "rtf"
            RetFileType = "application/rtf"
        Case "pdf"
            RetFileType = "application/pdf"
        Case "ra"
            RetFileType = "audio/x-realaudio"
        Case "mid"
            RetFileType = "audio/x-midi"
        Case "wav"
            RetFileType = "audio/x-wav"
        Case "dll"
            RetFileType = "dll/dll"
        Case "wrl", "wrz"
            RetFileType = "x-world/x-vrml"
        Case "qt", "mov", "moov"
            RetFileType = "video/quicktime"
        Case "mpg", "mpeg", "mpe"
            RetFileType = "video/mpeg"
        Case "avi"
            RetFileType = "video/x-msvideo"
        Case "php", "php3", "php4"
            RetFileType = ""
        Case Else
            RetFileType = "application/x-unknown"
    End Select
End Function

Public Function FineScrittura(Filez As String) As Boolean
    Dim NumeroFile As Integer
    NumeroFile = FreeFile
On Error GoTo Writing:
    Open Filez For Binary Access Write Lock Write As NumeroFile
    FineScrittura = True
    Close NumeroFile
Exit Function
Writing:
    FineScrittura = False
    Close NumeroFile
    Err.Clear
End Function

Public Function DenyUrlStr(StrC As String, ByRef sLog As String) As Boolean
Dim RetVal() As String
Dim I As Integer
    
On Error GoTo Error
    RetVal = GetIniAppAllKeys(ConfigFileName, "DENYURLSTRING")
    
    For I = 0 To UBound(RetVal)
        If InStr(1, StrC, GetIni(ConfigFileName, "DENYURLSTRING", RetVal(I), ""), vbTextCompare) <> 0 Then
            sLog = sLog & SepTag & "[Unallowed String in URL]"
            DenyUrlStr = True
            Exit For
        End If
    Next I
Error:

End Function

Public Function BannedIP(StrC As String) As Boolean
Dim RetVal() As String
Dim Temp() As String
Dim bIPClient(3) As Byte
Dim bIP(3) As Byte
Dim bMask(3) As Byte
Dim bTemp As Byte
Dim I As Integer
Dim J As Integer

On Error GoTo Error
    RetVal = GetIniAppAllKeys(ConfigFileName, "BANNEDIP")
    Call DeCompIP(StrC, bIPClient)

    
    For I = 0 To UBound(RetVal)
        BannedIP = True
        RetVal(I) = GetIni(ConfigFileName, "BANNEDIP", RetVal(I), "255.255.255.255 255.255.255.255")
        Temp = Split(RetVal(I), " ")
        Call DeCompIP(Temp(0), bIP)
        Call DeCompIP(Temp(UBound(Temp)), bMask)
        For J = 0 To 3
            bTemp = bIPClient(J) And bMask(J)
            If bTemp <> bIP(J) Then
     
                BannedIP = False
                Exit For
            
            End If
        Next J
    Next I
    
    If BannedIP Then
        RetVal = GetIniAppAllKeys(ConfigFileName, "TRUSTIP")
        
        For I = 0 To UBound(RetVal)
            BannedIP = False
            RetVal(I) = GetIni(ConfigFileName, "TRUSTIP", RetVal(I), "255.255.255.255 255.255.255.255")
            Temp = Split(RetVal(I), " ")
            Call DeCompIP(Temp(0), bIP)
            Call DeCompIP(Temp(UBound(Temp)), bMask)
            For J = 0 To 3
                bTemp = bIPClient(J) And bMask(J)
                If bTemp <> bIP(J) Then
         
                    BannedIP = True
                    Exit For
                ElseIf J = 3 Then
                    BannedIP = False
                    Exit Function
                End If
            Next J
        Next I
    End If
Error:

End Function

Private Sub DeCompIP(sIP As String, ByRef aDest() As Byte)
    Dim Temp() As String
    Dim aDeCompIP(3) As Byte
    Dim I As Byte
    
    Temp = Split(sIP, ".")
    For I = 0 To 3
        aDeCompIP(I) = Temp(I)
    Next

    Call CopyMemory(aDest(0), aDeCompIP(0), 4)
    
End Sub


'Public Function GZIPCompressStaticFiles(sFile As String) As String
'
'    Dim R As Long
'    Dim Raw_Len As Long
'    Dim Raw_Data As String
'    Dim FreeFileNumber As Integer
'    Dim Cmp_Len As Long
'    Dim Cmp_Data As String
'    Dim gFile As Long
'
'    Raw_Len = FileLen(sFile)
'    Raw_Data = Space(Raw_Len)
'
'    FreeFileNumber = FreeFile
'    Open sFile For Binary As #FreeFileNumber
'        Get #FreeFileNumber, , Raw_Data
'    Close #FreeFileNumber
'
'    'establish a buffer for the compressed data
'    Cmp_Len = Raw_Len + 12 + (Raw_Len / 1000)
'    Cmp_Data = Space(Cmp_Len)
'
'    gFile = gzopen(sFile & ".1", "wb")
'    If gFile <= 0 Then
'        GZIPCompressStaticFiles = "The gzopen() for write call failed."
'        Exit Function
'    End If
'    R = gzwrite(gFile, Raw_Data, Raw_Len)
'    If R <> Raw_Len Then
'        GZIPCompressStaticFiles = "The gzwrite() call returned an error"
'        Exit Function
'    End If
'    R = gzclose(gFile)
'    If R <> 0 Then
'        GZIPCompressStaticFiles = "The gzclose() called failed."
'    End If
'
'End Function

Public Function GZIPCompress(Raw_Data As String, SockHnd As Long) As String

    Dim R As Long
    Dim Raw_Len As Long
    Dim gFile As Long
    Dim sFile As String
    
    sFile = (App.path + "\" + GetIni(ConfigFileName, "MAIN", "TMPFOLDER", "\")) & _
             "\gzip_" & SockHnd & ".$$$"
          
    Raw_Len = Len(Raw_Data)
   
    gFile = gzopen(sFile, "wb")
    If gFile <= 0 Then
        GZIPCompress = "The gzopen() for write call failed."
        Exit Function
    End If
    R = gzwrite(gFile, Raw_Data, Raw_Len)
    If R <> Raw_Len Then
        GZIPCompress = "The gzwrite() call returned an error"
        Exit Function
    End If
    R = gzclose(gFile)
    If R <> 0 Then
        GZIPCompress = "The gzclose() called failed."
    End If
    
    Dim NumeroFile As Integer
    NumeroFile = FreeFile
    Open sFile For Binary Access Read Shared As NumeroFile
        Raw_Data = String(LOF(NumeroFile), 1)
        Get #NumeroFile, 1, Raw_Data
    Close NumeroFile
    
    Kill sFile
    
    GZIPCompress = Raw_Data

End Function

Public Function UnicodeTrans(sStr As String) As String
    Dim I As Byte
    
    For I = 15 To 254
            sStr = Replace$(sStr, LCase$("%" & Hex(I + 1)), Chr(I + 1))
    Next I
    UnicodeTrans = sStr
End Function

Public Sub ErrorGST(ErrorString As String)

    If CBool(GetIni(ConfigFileName, "MAIN", "EVLOG", "False")) Then
        Call App.LogEvent(ErrorString, vbLogEventTypeError)
    End If
    Writeln ErrorString

End Sub
