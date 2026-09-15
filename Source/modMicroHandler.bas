Attribute VB_Name = "modMicroHandler"
Option Explicit

Public m_Sending As Boolean

Public Enum m_Enum_MHAction
    LOGIN = 0
    LOGOUT = 1
    SNDDATA = 2
    SHOWUI = 66
    HIDEUI = 99
End Enum

Public Function SendCommand(CMD As m_Enum_MHAction, _
                            Optional sData As String = "Nope.", _
                            Optional iDataType As Integer = 0) As String
    Dim sTI As Single
    Dim lStart As Long
    Dim lEnd As Long
    
    Const byTimeOut As Byte = 255

On Error GoTo ErrH
    
    If CBool(GetINI(cfgFile, "PREFERENCE", "MH", "FALSE")) Then
        If Not m_Sending Then
            m_Sending = True
    
            With xmp
                If .ctrlWsk.State <> sckConnected Then
                    .ctrlWsk.Protocol = sckTCPProtocol
                    .ctrlWsk.Connect "localhost", 8080
                End If
        
                gb_MHACTIONS = CMD
                gb_MHDATA = sData
                gb_MHDATATYPE = iDataType
                
                sTI = Timer()
                
                Do
                    Call Pause(DoEv:=True)
                Loop Until gb_MHRESPONS <> "" Or Timer() - sTI > byTimeOut
                
                
                If gb_MHRESPONS <> "" Then
                    Select Case CMD
                        Case Is = LOGIN
                            lStart = InStr(1, gb_MHRESPONS, "sessionid=") + 10
                            lEnd = InStr(lStart, gb_MHRESPONS, vbCrLf)
                            gb_MHSESSIONID = Mid$(gb_MHRESPONS, lStart, lEnd - lStart)
                        Case Is = LOGOUT
                        Case Is = SNDDATA
                        Case Is = SHOWUI
                        Case Is = HIDEUI
                    End Select
                Else
                    Call Err.Raise(666, "modMicroHandler", "Comunication Error or TimeOut")
                End If
                
                .ctrlWsk.Close
            End With
            
            SendCommand = gb_MHRESPONS
            gb_MHRESPONS = ""
            m_Sending = False
        Else
            Exit Function
        End If
    End If
    
    Exit Function
ErrH:
    If Err.Number <> 0 Then
        xmp.ctrlWsk.Close
        gb_MHRESPONS = ""
        m_Sending = False
        DisplayError (Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "modMicroHandler.SendCommand" & vbCrLf & Err.Description)
    End If
End Function

