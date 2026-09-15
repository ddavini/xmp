VERSION 5.00
Begin VB.UserControl xmDisplay 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   LockControls    =   -1  'True
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   114
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "xmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Mode
    A
    B
    C
    D
End Enum

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Siz As Integer
Private StrDisplay As String
Private Texture As IPictureDisp
Private mCenter As Boolean

Event Click()


Private Sub UserControl_InitProperties()

    StrDisplay = ""
    Siz = 5
    mCenter = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    StrDisplay = PropBag.ReadProperty("xCaption", "")
    Siz = PropBag.ReadProperty("Size", 5)
    mCenter = PropBag.ReadProperty("Center", False)
       
    xCaption = StrDisplay
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("xCaption", StrDisplay, "")
    Call PropBag.WriteProperty("Size", Siz, 5)
    Call PropBag.WriteProperty("Center", mCenter, False)
    
End Sub




Private Sub DrawChar(Carattere As String, picScreen As PictureBox, _
                    Posizione As Integer, Dimensione As Integer, _
                    Optional Color As Mode = A)
    Const xBitSize = 5
    Const yBitSize = 6
    
    Select Case UCase(Carattere)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            (Asc(UCase(Carattere)) - 48), 1, Color)
        Case "(", ")"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            (13 + Asc(Carattere) - 40), 1, Color)
        Case """"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            26, 0, Color)
        Case "@"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            27, 0, Color)
            Call PaintChar(picScreen, Dimensione, Posizione + 1, _
            28, 0, Color)
        Case " "
            Call PaintChar(picScreen, Dimensione, Posizione + 1, _
            29, 0, Color)
        Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            (Asc(UCase(Carattere)) - 65), 0, Color)
        Case "..."
            Call PaintChar(picScreen, Dimensione, Posizione, _
            10, 1, Color)
        Case ":"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            12, 1, Color)
        Case "-"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            15, 1, Color)
        Case "'"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            16, 1, Color)
        Case "!"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            17, 1, Color)
        Case "_"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            18, 1, Color)
        Case "+"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            19, 1, Color)
        Case "\"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            20, 1, Color)
        Case "/"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            21, 1, Color)
        Case "["
            Call PaintChar(picScreen, Dimensione, Posizione, _
            22, 1, Color)
        Case "]"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            23, 1, Color)
        Case "^"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            24, 1, Color)
        Case "&"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            25, 1, Color)
        Case "&"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            26, 1, Color)
        Case "."
            Call PaintChar(picScreen, Dimensione, Posizione, _
            27, 1, Color)
        Case "="
            Call PaintChar(picScreen, Dimensione, Posizione, _
            28, 1, Color)
        Case "$"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            29, 1, Color)
        Case "?"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            3, 2, Color)
        Case "*"
            Call PaintChar(picScreen, Dimensione, Posizione, _
            4, 2, Color)
        Case Else
            Call PaintChar(picScreen, Dimensione, Posizione, _
            11, 1, Color)
    End Select

End Sub

Public Property Get Size() As Long

    Size = Siz
    
End Property

Public Property Let Size(ByVal Value As Long)
    
    Picture1.Height = Value + 3
    Siz = Value

End Property

Public Property Get Center() As Boolean

    Center = mCenter
    
End Property

Public Property Let Center(ByVal Value As Boolean)
    
    mCenter = Value

End Property

Public Property Get xCaption() As String
    xCaption = StrDisplay
End Property

Public Property Let xCaption(ByVal sStr As String)
Dim aChar() As Byte
Dim I As Integer
On Error GoTo ErrH:
    sStr = Replace$(sStr, vbNullChar, " ")
    If sStr <> "" Then
        ReDim aChar(Len(sStr) - 1)
        Call CopyMemory(aChar(0), ByVal sStr, Len(sStr))
             
        Picture1.Cls
        StrDisplay = sStr
        For I = 0 To UBound(aChar)
            'Debug.Print Chr(aChar(I))
            Call DrawChar(Chr(aChar(I)), Picture1, I, Siz, C)
        Next I
    End If
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmDisplay.xCaption" & vbCrLf & Err.Description)
    End If
End Property

Private Sub PaintChar(picScreen As PictureBox, Dimension As Integer, Position As Integer, _
                   XChar As Integer, YChar As Integer, Color As Mode)

    Const xBitSize = 5
    Const yBitSize = 6
    Dim X1 As Long
On Error GoTo ErrH:
    If mCenter Then
        X1 = (UserControl.ScaleWidth - (Dimension * Len(StrDisplay))) / 2
    Else
        X1 = 0
    End If
    Call picScreen.PaintPicture(Texture _
     , (Dimension * Position) + X1, 0, Dimension, Dimension + 1, _
    XChar * xBitSize, (yBitSize * YChar) + (19 * Color), xBitSize, yBitSize)
ErrH:
    If Err.Number <> 0 Then
        Call DisplayError(Err.Number & vbCrLf & Err.Source & _
                vbCrLf & "xmDisplay.PaintChar" & vbCrLf & Err.Description)
    End If
End Sub


Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Siz = 5
    UserControl.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    Picture1.Height = Siz + 3
    UserControl.ScaleHeight = Picture1.Height
    Picture1.Width = UserControl.ScaleWidth
    Set Texture = LoadDataIntoFile("Display")
End Sub

Private Sub UserControl_Resize()
    Picture1.Height = Siz + 3
    UserControl.ScaleHeight = Picture1.Height
    Picture1.Width = UserControl.ScaleWidth
End Sub
