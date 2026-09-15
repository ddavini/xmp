VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsECG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ECG (Ethereal Chaos Generator)
' By: David Midkiff <mdj2023@hotmail.com>
'
' ECG is a strong, secure, and unpredictable pseudo-random
' number generator or Visual Basic. Read attached readme.txt,
' quicknotes.txt and prng intro.txt for full details.
'
' ECG was specifically designed for use in cryptographic algorithms.

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (pGUID As GUID, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long

Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

Private Type GUID
    Guid1 As Long
    Guid2 As Long
    Guid3 As Long
    Guid4(0 To 7) As Byte
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Function GSeed() As Double
    ' Retrieves the GUID value into a double variable
    Dim udtGUID As GUID, sGUID As String, lResult As Long

    lResult = CoCreateGuid(udtGUID)
    If lResult Then
        GSeed = 0
    Else
        With udtGUID
            GSeed = CInt(CLng((CDbl(.Guid1) + CDbl(.Guid2) + CDbl(.Guid3)) / CDbl(32766)) Mod 32766)
        End With
    End If
End Function
Private Function HASH(Data As String) As Double
    ' Simple, primitive yet fast 80-bit one-way hash function designed by
    ' David Midkiff for use only in Ethereal Chaos Generator. To use this
    ' elsewhere you must have David's permission.
    
    Dim I As Double, NBox As Long, OBox As Long, PBox As Long, QBox As Long, RBox As Long
    For I = 1 To Len(Data)
        NBox = NBox + Asc(Mid$(Data, I, 1))
        NBox = (NBox * 1765 + 1066) Mod 1048576
        OBox = OBox + Asc(Mid$(Data, I, 1))
        OBox = (OBox * 815 + 1801) Mod 1048576
        PBox = PBox + Asc(Mid$(Data, I, 1))
        PBox = (PBox * 1364 + 1083) Mod 1048576
        QBox = QBox + Asc(Mid$(Data, I, 1))
        QBox = (QBox * 1624 + 1564) Mod 1048576
        RBox = RBox + Asc(Mid$(Data, I, 1))
        RBox = (RBox * 1454 + 1905) Mod 1048576
    Next
    
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576
    
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576
    
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576
    
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576
    
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576

    HASH = NBox + OBox + PBox + QBox + RBox
End Function
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    ' Appends data into a memory variable providing much faster
    ' concatentation then the conventional & statement.
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Public Function RandomNumber(Lower As Double, Upper As Double, Optional CustomSeed As Double) As Double
    Dim HSeed As Double, SSeed As Double, RSeed As Double
    Dim TSeed As Double, USeed As Double, VSeed As Double
    Dim WSeed As Double, XSeed As Double, FinalSeed As Double

    Dim iLoop As Long, R As String
    Dim Mouse As POINTAPI, Dimensions As RECT
    
    ' Retrieves the mouse position.
    GetCursorPos Mouse
    
    ' Sets the current millisecond and mouse position as the initial seed.
    Randomize (GetTickCount + Mouse.X + Mouse.Y)
    
    ' HSeed is generated by hashing a random number generated from the
    ' above seed.
    HSeed = HASH(CStr((Int(Rnd * 1000000) + Mouse.X - Mouse.Y)))

    For iLoop = 1 To 1000
        ' Retrieves window dimensions into variable Dimensions.
        GetWindowRect iLoop, Dimensions
        
        ' RSeed is genereated by accumulating window caption lenghts.
        RSeed = RSeed + GetWindowTextLength(iLoop)
        
        ' XSeed is generated by accumulating window dimensions.
        XSeed = XSeed + Dimensions.Bottom + Dimensions.Top + Dimensions.Left + Dimensions.Right
    Next
    
    ' Retrieves the mouse position.
    GetCursorPos Mouse
    
    ' SSeed is generated from the product of Upper and Lower.
    SSeed = Upper + Lower
    
    ' TSeed and USeed are genereated using the mouse position.
    TSeed = Mouse.X
    USeed = Mouse.Y
    
    ' VSeed is generated from the millisecond clock.
    VSeed = GetTickCount
    
    ' Four of the above seeds and GetTickCount is used as the
    ' next seed to produce WSeed.
    Randomize (GetTickCount + RSeed + XSeed + TSeed + USeed)
    WSeed = Int(Rnd * 9999999999#) + 1
    
    ' The final seed is generated using all of the above seeds.
    FinalSeed = GSeed + HSeed + SSeed + RSeed + TSeed + USeed + VSeed + WSeed + XSeed + CustomSeed
    
    ' The final number is generated.
    Randomize FinalSeed
    RandomNumber = Int(Rnd * (Upper - Lower + 1)) + Lower
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function RandomString(Length As Double) As String
    Dim iLoop As Double
    Reset ' Reset the concatenation variable
    For iLoop = 1 To Length ' Loop to the given length
        Append Chr$(RandomNumber(1, 255)) ' Store a random character between 1-255 in the concatenation memory location
    Next
    RandomString = GData ' Retrieve the string from memory
    Reset ' Reset the concatenation variable
End Function
