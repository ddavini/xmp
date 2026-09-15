Attribute VB_Name = "modAPL"
Option Explicit

'v0.0.1

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public RtString As String

Private Const APLStdSubDir = "\APL\"
Private Const APLStdWSCFle = "WINSOCK.DAT"

Public Function Win32FileExist(FullFilePath As String, _
                            Optional FileExistMode As Byte = 0, _
                            Optional TestFileLen As Long) As Boolean
                                
    Win32FileExist = False
    If FileExistMode = 0 Then
        If Dir(FullFilePath) > "" Then
            Win32FileExist = True
        End If
    Else
        If FileExist(FullFilePath) Then
            Win32FileExist = True
        End If
    End If
    If Win32FileExist Then
        Call RtStringHandler(LCase$(FullFilePath) & " " & "Exist")
        If TestFileLen > 0 Then
            If FileLen(FullFilePath) <> TestFileLen Then
                Call RtStringHandler(LCase$(FullFilePath) & " " & "Exist but FileLen is worng")
            End If
        End If
    Else
        Call RtStringHandler(LCase$(FullFilePath) & " " & "Not Exist")
    End If
End Function

Private Function FileExist(FileName As String) As Boolean
    Dim FiNu As Long
On Error Resume Next
    FiNu = FreeFile
    Open FileName For Input As #FiNu
        If Err.Number = 53 Then
            FileExist = False
        Else
            FileExist = True
        End If
    Close #FiNu
End Function

Private Sub RtStringHandler(sStr As String)

    RtString = RtString & sStr & vbCrLf
    
End Sub

Public Function ControlForAPLFile(Optional Action As Boolean = False) As Boolean
    Dim sStr As String * 255
    Dim WSCPath As String
    Call GetSystemDirectory(sStr, 255)
    WSCPath = Replace$(sStr, vbNullChar, "") & "\MSWINSCK.OCX"

    If Win32FileExist(WSCPath, 0, 109248) Then
        ControlForAPLFile = True
    Else
        ControlForAPLFile = False
        If Action Then
            Call FileCopy(App.Path & APLStdSubDir & APLStdWSCFle, WSCPath)
        End If
    End If
End Function

