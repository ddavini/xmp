Attribute VB_Name = "modVariabiliGlobali"
Option Explicit

'''MH'''
Public gb_MHRESPONS As String
Public gb_MHACTIONS As m_Enum_MHAction
Public gb_MHSESSIONID As Long
Public gb_MHDATA As String
Public gb_MHDATATYPE As Integer
''''''''

Public Enum SuGiu
    Su = 1
    Giu = -1
End Enum

Public Enum TipoModo
    Mono = 0
    Stereo = 1
    XSound = 2
    Azzerato = -1
End Enum

Private Type gPos
    SamplePos As Long
    Mp3File As String
End Type

'Costante numero massimo effetti per pulsante Shock su xmp
Public Const MaxEffect = 3

Public gb_OldProcMouseWheel As Long

Public AgganciatoFlag As Boolean
'Public Directory() As String
Public Minimizzato As Boolean
Public IndiceGlobalissimo As Integer
Public STRM As Long
'0 mono 1 stereo -1 Azzerato
Public Modo As Integer
Public GlobalSpectrumeMode As SpectrumeMode
Public frmVisualizzazioniAttivo As Boolean
Public gPosForSave As gPos

'Public geVal(9) As Long
Public StepRate As Integer
Public SpecMode As Byte
Public FFTSAMPLE As Integer
Public cfgFile As String

Public Const gcDeactiveColor As Long = 6619080 'vbGrayText
Public Const gcActiveColor As Long = vbGreen

Public Const Mp3TmpFiles As String = "files.$$$"

Public gb_Data As InputInfo

'List Search
Public gb_LstShearchIndex As Integer
Public gb_LstShearchChr As String

'Analyzer
Public Gap As Byte
Public Barwidth As Byte
Public PointFalls As Byte
Public Tolleranza As Byte
Public FallsVel As Single
Public noFallsVel As Single


Public gb_InfoMp3Active As Boolean
Public gb_InShutdown As Boolean
Public gb_InfoLastLine As Byte

Public Function Estensioni() As Variant
On Error GoTo ErrH
    Dim Ext As String
    Dim Store As String
    Dim L As Integer
    Dim I As Integer
    Dim J As Integer
    Dim LocalArray() As String

    Ext = GetINI(cfgFile, "INSTALL", "EXT", "MP3;")
    I = 1
    Do
        L = InStr(I, Ext, ";", vbTextCompare) - I
        Store = Mid$(Ext, I, L)
        ReDim Preserve LocalArray(J)
        LocalArray(J) = "." + Store
        I = I + L + 1
        J = J + 1
    Loop Until I > Len(Ext)
    Estensioni = LocalArray
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "VariabiliGlobali.Estensioni" & vbCrLf & Err.Description)
    End If
End Function


