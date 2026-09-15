Attribute VB_Name = "modGestioneIni"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Base On Daniel Cohen Gindi Sample                          '
'http://www.vbdiamond.com                                   '
'dcgsmix@surfree.net.il                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetINI(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal Default As String, _
                       Optional ByVal Min As Variant, Optional Max As Variant) As String
   
   
   Dim Prv As Long
   Dim RetVal As String
   
   RetVal = String$(255, 0)
   Prv = GetPrivateProfileString(AppName, KeyName, "", RetVal, Len(RetVal), FileName)
   If Prv = 0 Then
      GetINI = Default
   Else
      GetINI = Left$(RetVal, InStr(RetVal, Chr(0)) - 1)
   End If
   If IsNumeric(GetINI) Then
        If Not IsMissing(Min) Then
            If CLng(GetINI) < Min Then
                GetINI = CStr(Min)
            End If
        End If
        If Not IsMissing(Max) Then
            If CLng(GetINI) > Max Then
                GetINI = CStr(Max)
            End If
        End If
   End If
End Function

Public Function SetINI(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String) As Long
   SetINI = WritePrivateProfileString(AppName, KeyName, KeyValue, FileName)
End Function

Public Function DelINIApp(ByVal FileName As String, ByVal AppName As String) As Long
   DelINIApp = WritePrivateProfileString(AppName, vbNullString, "", FileName)
End Function

Public Function DelINIKey(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String) As Long
   DelINIKey = WritePrivateProfileString(AppName, KeyName, vbNullString, FileName)
End Function

Public Function GetINIAllApps(ByVal FileName As String) As String()
   Dim RetVal As String
   Dim TMP As String, TMP2 As String
   Dim Prv As Long, I As Long, J As Long
   Dim Apps() As String
   
   RetVal = String$(1024, 0)
   Prv = GetPrivateProfileString(vbNullString, vbNullString, "", RetVal, Len(RetVal), FileName)
   RetVal = Left$(RetVal, Prv)
   For I = 1 To Len(RetVal)
      TMP2 = Mid$(RetVal, I, 1)
      If TMP2 <> Chr(0) Then
         TMP = TMP & TMP2
      ElseIf TMP2 = Chr(0) And TMP <> "" Then
         ReDim Preserve Apps(J)
         Apps(J) = TMP
         J = J + 1
         TMP = ""
      End If
   Next
   GetINIAllApps = Apps
End Function

Public Function GetINIAppAllKeys(ByVal FileName As String, ByVal AppName As String) As String()
   Dim RetVal As String
   Dim TMP As String, TMP2 As String
   Dim Prv As Long, I As Long, J As Long
   Dim Keys() As String
   
   RetVal = String$(1024, 0)
   Prv = GetPrivateProfileString(AppName, vbNullString, "", RetVal, Len(RetVal), FileName)
   RetVal = Left$(RetVal, Prv)
   For I = 1 To Len(RetVal)
      TMP2 = Mid$(RetVal, I, 1)
      If TMP2 <> Chr(0) Then
         TMP = TMP & TMP2
      ElseIf TMP2 = Chr(0) And TMP <> "" Then
         ReDim Preserve Keys(J)
         Keys(J) = TMP
         J = J + 1
         TMP = ""
      End If
   Next
   GetINIAppAllKeys = Keys
End Function
