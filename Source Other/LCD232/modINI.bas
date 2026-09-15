Attribute VB_Name = "modGestioneIni"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Base On Daniel Cohen Gindi Sample                          '
'http://www.vbdiamond.com                                   '
'dcgsmix@surfree.net.il                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'v2 0#

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function GetIni(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal Default As String, _
                       Optional ByVal Min As Variant, Optional Max As Variant) As String
   
   
   Dim Prv As Long
   Dim RetVal As String
   
   If FileName = "" Then
        FileName = App.Path & "\" & App.EXEName & ".ini"
   End If
   
   RetVal = String$(255, 0)
   Prv = GetPrivateProfileString(AppName, KeyName, "", RetVal, Len(RetVal), FileName)
   If Prv = 0 Then
      GetIni = Default
   Else
      GetIni = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
   End If
   If IsNumeric(GetIni) Then
    If Not IsMissing(Min) Then
        If CLng(GetIni) < Min Then
            GetIni = CStr(Min)
        End If
    End If
    If Not IsMissing(Max) Then
        If CLng(GetIni) > Max Then
            GetIni = CStr(Max)
        End If
    End If
   End If
End Function

Public Function SetIni(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String) As Long
   SetIni = WritePrivateProfileString(AppName, KeyName, KeyValue, FileName)
End Function

Public Function DelIniApp(ByVal FileName As String, ByVal AppName As String) As Long
   DelIniApp = WritePrivateProfileString(AppName, vbNullString, "", FileName)
End Function

Public Function DelIniKey(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String) As Long
   DelIniKey = WritePrivateProfileString(AppName, KeyName, vbNullString, FileName)
End Function

Public Function GetIniAllApps(ByVal FileName As String) As String()
   Dim RetVal           As String
   Dim tmp              As String, tmp2 As String
   Dim Prv As Long, I As Long, J As Long
   Dim Apps() As String
   
   RetVal = String$(1024, 0)
   Prv = GetPrivateProfileString(vbNullString, vbNullString, "", RetVal, Len(RetVal), FileName)
   RetVal = Left(RetVal, Prv)
   For I = 1 To Len(RetVal)
      tmp2 = Mid(RetVal, I, 1)
      If tmp2 <> Chr(0) Then
         tmp = tmp & tmp2
      ElseIf tmp2 = Chr(0) And tmp <> "" Then
         ReDim Preserve Apps(J)
         Apps(J) = tmp
         J = J + 1
         tmp = ""
      End If
   Next
   GetIniAllApps = Apps
End Function

Public Function GetIniAppAllKeys(ByVal FileName As String, ByVal AppName As String) As String()
   Dim RetVal As String
   Dim tmp As String, tmp2 As String
   Dim Prv As Long, I As Long, J As Long
   Dim Keys() As String
   
   RetVal = String$(1024, 0)
   Prv = GetPrivateProfileString(AppName, vbNullString, "", RetVal, Len(RetVal), FileName)
   RetVal = Left(RetVal, Prv)
   For I = 1 To Len(RetVal)
      tmp2 = Mid(RetVal, I, 1)
      If tmp2 <> Chr(0) Then
         tmp = tmp & tmp2
      ElseIf tmp2 = Chr(0) And tmp <> "" Then
         ReDim Preserve Keys(J)
         Keys(J) = tmp
         J = J + 1
         tmp = ""
      End If
   Next
   GetIniAppAllKeys = Keys
End Function
