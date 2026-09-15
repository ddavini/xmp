Attribute VB_Name = "modCallBack"
Option Explicit

Public Sub xmMP3_Proc(ByVal ID As Long, ByVal param As Long)
    Select Case ID
    Case MSG_ERROR
    Case MSG_PLAYING
        xmp.g_PlayDone = False
    Case MSG_PAUSING
        xmp.g_PlayDone = False
    Case MSG_STOPING
        xmp.g_PlayDone = False
    Case MSG_PLAYDONE
        xmp.g_PlayDone = True
    End Select
End Sub

