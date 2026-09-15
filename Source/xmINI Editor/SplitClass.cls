VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-------------------------------------------'
'          Made by Dennis Meelker           '
'                                           '
'You can freely distribute and use this code'
'-------------------------------------------'
Private mvarWordAmount As Integer
Private Wordarray As Variant
Public Function Parse(sIn As String, sDel As String) As Variant
    On Error Resume Next
    Dim i As Integer, x As Integer, s As Integer, t As Integer
    i = 1: s = 1: t = 1: x = 1
    ReDim tArr(1 To x) As Variant


    If InStr(1, sIn, sDel) <> 0 Then


        Do
            ReDim Preserve tArr(1 To x) As Variant
            tArr(i) = mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            x = x + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(1 To x) As Variant
        tArr(i) = mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(1) = sIn
    End If
    mvarWordAmount = i
    Wordarray = tArr
End Function

Public Property Let WordAmount(ByVal vData As Integer)
    mvarWordAmount = vData
End Property


Public Property Get WordAmount() As Integer
    WordAmount = mvarWordAmount
End Property


Public Property Get Word(index) As String
    On Error Resume Next
    Word = Wordarray((index))
End Property



