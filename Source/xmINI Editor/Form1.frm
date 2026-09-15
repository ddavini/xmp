VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "xmINI Editor"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTmp 
      Height          =   2175
      Left            =   5400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtIni 
      Height          =   7185
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   360
      Width           =   7935
   End
   Begin VB.PictureBox Toolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10860
      TabIndex        =   4
      Top             =   0
      Width           =   10860
      Begin xmINI.Button Button5 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "Form1.frx":0CCA
         LineColor2      =   8421504
      End
      Begin xmINI.Button Button3 
         Height          =   375
         Left            =   860
         TabIndex        =   7
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "Form1.frx":0DDC
         LineColor2      =   8421504
      End
      Begin xmINI.Button Button2 
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "Form1.frx":112E
         LineColor2      =   8421504
      End
      Begin xmINI.Button Button1 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "Form1.frx":1480
         LineColor2      =   8421504
      End
      Begin xmINI.Button Button4 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "Form1.frx":17D2
         LineColor2      =   8421504
      End
   End
   Begin VB.PictureBox split 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   2880
      ScaleHeight     =   7095
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.ListBox lstSections 
      Height          =   7080
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox mid 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   2760
      MouseIcon       =   "Form1.frx":1B24
      MousePointer    =   99  'Custom
      ScaleHeight     =   7095
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   360
      Width           =   105
   End
   Begin VB.ListBox lstTemp 
      Height          =   1815
      Left            =   5760
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditInsertFile 
         Caption         =   "&Insert File"
         Shortcut        =   ^{INSERT}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsRefresh 
         Caption         =   "&Update Section List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFonts 
         Caption         =   "&Fonts"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private YepResize As Boolean
Private SectionDeleted As Boolean
Private LastText As String
Public OpenedFile As String
Public Saved As Boolean

Dim Splitter As New SplitClass

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
    Case 1
        mnuFileNew_Click
    Case 3
        mnuOpen_Click
    Case 4
        mnuSave_Click
    Case 6
        mnuPrint_Click
    Case 8
        mnuEditCut_Click
    Case 9
        mnuEditCopy_Click
    Case 10
        mnuEditPaste_Click
End Select
End Sub

Private Sub Button1_Click()
mnuFileNew_Click
End Sub

Private Sub Button2_Click()
mnuOpen_Click
End Sub

Private Sub Button3_Click()
If OpenedFile = "" Then
        mnuSaveAs_Click
    Else
        mnuSave_Click
End If
End Sub

Private Sub Button4_Click()
mnuPrint_Click
End Sub

Private Sub Button5_Click()
mnuEditFind_Click
End Sub

Private Sub Form_Load()
Dim FF As Integer
Dim TextLine As String
Dim ReadString As String

FF = FreeFile
YepResize = True

If Len(Dir(AddSlash(App.Path) & "FontSettings.dat")) > 0 Then
    Open AddSlash(App.Path) & "FontSettings.dat" For Input As FF
        
        Line Input #FF, TextLine
        lstSections.FontBold = TextLine
        txtIni.Font.Bold = TextLine
        
        Line Input #FF, TextLine
        lstSections.FontItalic = TextLine
        txtIni.Font.Italic = TextLine
        
        Line Input #FF, TextLine
        lstSections.FontName = TextLine
        txtIni.Font.Name = TextLine
        
        Line Input #FF, TextLine
        lstSections.FontSize = TextLine
        txtIni.Font.Size = TextLine
        
    Close FF
End If

FF = FreeFile

If Len(Dir(AddSlash(App.Path) & "RecentFiles.dat")) > 0 Then
    mnuSep4.Visible = True
    Open AddSlash(App.Path) & "RecentFiles.dat" For Input As FF 'AddSlash(App.Path) & "RecentFiles.dat" For Input As FF
        Do Until EOF(FF)
            Line Input #FF, TextLine
            If Len(Dir(TextLine)) > 0 And Trim$(TextLine) <> "" Then
                Load mnuRecent(mnuRecent.count)
                mnuRecent(mnuRecent.count - 1).Caption = TextLine
                mnuRecent(mnuRecent.count - 1).Visible = True
            End If
        Loop
    Close FF
Else
    mnuSep4.Visible = False
End If

If mnuRecent.count = 1 Then mnuSep4.Visible = False

Saved = True
    
    
    If Trim$(Command$) <> "" Then
        OpenedFile = Command$
        Call LoadINI(Command$)
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim FF As Integer
Dim I As Integer

FF = FreeFile
Open AddSlash(App.Path) & "RecentFiles.dat" For Output As FF 'AddSlash(App.Path) & "RecentFiles.dat" For Input As FF
    For I = 1 To mnuRecent.count - 1
        Print #FF, mnuRecent(I).Caption
    Next I
Close FF

FF = FreeFile

Open AddSlash(App.Path) & "FontSettings.dat" For Output As FF
        
        Print #FF, lstSections.FontBold
        Print #FF, lstSections.FontItalic
        Print #FF, lstSections.FontName
        Print #FF, lstSections.FontSize

Close FF

For I = Forms.count - 1 To 0 Step -1
        Unload Forms(I)
Next

If Saved = False Then
    If MsgBox("Save before quit?", vbInformation + vbYesNo, frmMain.Caption) = vbYes Then
        If OpenedFile = "" Then
            mnuSaveAs_Click
        Else
            mnuSave_Click
        End If
    End If
End If
End Sub

Private Sub Form_Resize()
If YepResize Then
    ResizeControls
End If
End Sub

Private Sub lstSections_Click()
Dim FoundPos As Integer
Dim LastItem As String
Dim SelectedItem As String

SelectedItem = lstSections.List(lstSections.ListIndex)
LastItem = lstSections.List(lstSections.ListCount - 1)

FoundPos = InStr(1, txtIni.Text, LastItem, vbTextCompare)
txtIni.SelStart = FoundPos - 1

FoundPos = InStr(1, txtIni.Text, SelectedItem, vbTextCompare)

If FoundPos = 0 Then Exit Sub

txtIni.SelStart = FoundPos - 1
txtIni.SelLength = Len(SelectedItem)
'result = SendMessageBynum(txtIni.hwnd, EM_LINESCROLL, 0, txtIni.GetLineFromChar(txtIni.SelStart))
'Debug.Print result
txtIni.SetFocus

'txtIni.Find lstSections.List(lstSections.ListIndex), 0, 0, rtfNoHighlight
'dl& = SendMessageBynum(Text1.hwnd, EM_LINESCROLL, 0, CLng(VScroll1.Value - firstvisible%))

End Sub

Public Function RefreshSectionList()
Dim FF As Integer
Dim I As Integer
Dim TextLine As String

lstTemp.Clear

'txtIni.SaveFile AddSlash(App.Path) & "tmp.ini", rtfText

Call SaveFileINI(AddSlash(App.Path) & "tmp.ini", txtIni)


FF = FreeFile
Open AddSlash(App.Path) & "tmp.ini" For Input As FF
    Do Until EOF(FF)
        Line Input #FF, TextLine
        If Left(TextLine, 1) = "[" Then lstTemp.AddItem TextLine
    Loop
Close FF

Kill AddSlash(App.Path) & "tmp.ini"

For I = 0 To lstTemp.ListCount - 1
    If Not lstSections.List(I) = lstTemp.List(I) Then
        GoTo FillSectionList
    End If
Next I
Exit Function

FillSectionList:
lstSections.Clear
For I = 0 To lstTemp.ListCount - 1
    lstSections.AddItem lstTemp.List(I)
Next I
End Function

Private Sub mid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
split.Left = txtIni.Left - (mid.Width / 2)
split.Top = lstSections.Top
split.Visible = True
End Sub

Private Sub mid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And mid.Visible = True Then
    split.Move X + txtIni.Left - 115 ', lstsections.Top, 3, lstsections.Height
End If
End Sub

Private Sub mid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
split.Visible = False
lstSections.Width = split.Left - lstSections.Left
mid.Left = split.Left
txtIni.Left = split.Left + mid.Width
txtIni.Width = (Me.ScaleWidth - txtIni.Left)

End Sub

Private Sub mnuEdit_Click()
If txtIni.SelLength = 0 Then
    mnuEditCopy.Enabled = False
    mnuEditCut.Enabled = False
    mnuEditDelete.Enabled = False
Else
    mnuEditCopy.Enabled = True
    mnuEditCut.Enabled = True
    mnuEditDelete.Enabled = True
End If
End Sub

Private Sub mnuEditAll_Click()
txtIni.SelStart = 0
txtIni.SelLength = Len(txtIni.Text)
txtIni.SetFocus
End Sub

Private Sub mnuEditCopy_Click()
Clipboard.Clear
Clipboard.SetText txtIni.SelText
End Sub

Private Sub mnuEditCut_Click()
Clipboard.Clear
Clipboard.SetText txtIni.SelText
txtIni.SelText = ""
End Sub

Private Sub mnuEditDelete_Click()
txtIni.SelText = ""
End Sub

Private Sub mnuEditFind_Click()
frmFind.Show
End Sub

Private Sub mnuEditInsertFile_Click()
On Error GoTo Canceled

Dim sOpen As SelectedFile

    
    FileDialog.sDefFileExt = "Ini Files(*.ini)"
    FileDialog.sFilter = "Ini Files(*.ini)" + Chr$(0) + "*.ini" + Chr$(0) + _
                         "All Files(*.*)" + Chr$(0) + "*.*"
                         
    
    FileDialog.sInitDir = App.Path
    FileDialog.sDlgTitle = "Open"
    
    FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY 'Or OFN_ALLOWMULTISELECT
    
    sOpen = ShowOpen(Me.hWnd, False)


    lstSections.Clear
    lstTemp.Clear

    txtTmp.Text = ""
    
    Call LoadFileINI(sOpen.sLastDirectory + sOpen.sFiles(0), txtTmp)
    'txtTemp.LoadFile CDlg.FileName
    
    txtIni.SelText = txtTmp.Text


Canceled:
End Sub

Private Sub mnuEditPaste_Click()
txtIni.SelText = Clipboard.GetText
RefreshSectionList
End Sub



Private Sub mnuEditReplace_Click()
frmReplace.Show
End Sub

Private Sub mnuEditUndo_Click()
SendMessage txtIni.hWnd, EM_UNDO, 0, 0
End Sub

Private Sub mnuFileNew_Click()
If Saved = False Then
    If MsgBox("Do you want to save your changes to the current document?", vbCritical + vbYesNo, frmMain.Caption) = vbYes Then
        If OpenedFile = "" Then
            mnuSaveAs_Click
        Else
            mnuSave_Click
        End If
    End If
End If

txtIni.Text = ""
lstSections.Clear
lstTemp.Clear
OpenedFile = ""
Saved = True
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Canceled
Randomize
Dim Rand As Integer
Dim sOpen As SelectedFile

    
    FileDialog.sDefFileExt = "Ini Files(*.ini)"
    FileDialog.sFilter = "Ini Files(*.ini)" + Chr$(0) + "*.ini" + Chr$(0) + _
                         "All Files(*.*)" + Chr$(0) + "*.*"
                         
    
    FileDialog.sInitDir = App.Path
    FileDialog.sDlgTitle = "Open"
    
    FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY 'Or OFN_ALLOWMULTISELECT
    
    sOpen = ShowOpen(Me.hWnd, False)


    lstSections.Clear
    lstTemp.Clear
    
    OpenedFile = sOpen.sLastDirectory + sOpen.sFiles(0)
    Call LoadINI(OpenedFile)
    
Canceled:
    Exit Sub
End Sub

Private Sub mnuOptionsFonts_Click()
frmFonts.Show vbModal, Me
End Sub

Private Sub mnuOptionsRefresh_Click()
RefreshSectionList
End Sub

Private Sub mnuPrint_Click()
On Error GoTo Canceled


Call ShowPrinter(Me.hWnd, True)

Printer.Copies = PrintDialog.nCopies
Printer.FontBold = txtIni.Font.Bold
Printer.FontItalic = txtIni.Font.Italic
Printer.FontUnderline = txtIni.Font.Underline
Printer.FontSize = txtIni.Font.Size
Printer.FontName = txtIni.Font.Name

Printer.Print txtIni.Text
Printer.EndDoc

Canceled:
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuRecent_Click(index As Integer)
lstSections.Clear
lstTemp.Clear
Call LoadFileINI(mnuRecent(index).Caption, txtIni)
'txtIni.LoadFile mnuRecent(index).Caption
OpenedFile = mnuRecent(index).Caption

Saved = True

RefreshSectionList
End Sub

Private Sub mnuSave_Click()
    If SaveFileINI(OpenedFile, txtIni) Then
        Saved = True
        Call SaveBox
    End If
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo Canceled
Randomize
Dim Rand As Integer
Dim I As Integer
Dim sOpen As SelectedFile

    FileDialog.sDefFileExt = "Ini Files(*.ini)"
    FileDialog.sFilter = "Ini Files(*.ini)" + Chr$(0) + "*.ini" + Chr$(0) + _
                         "All Files(*.*)" + Chr$(0) + "*.*"
                         
    
    FileDialog.sInitDir = App.Path
    FileDialog.sDlgTitle = "Save as..."
    
    FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY 'Or OFN_ALLOWMULTISELECT
    
    sOpen = ShowSave(Me.hWnd, False)

    OpenedFile = sOpen.sLastDirectory + sOpen.sFiles(0)

    If SaveFileINI(sOpen.sLastDirectory + sOpen.sFiles(0), txtIni) Then
        Call SaveBox
    End If

    Saved = True

    'Add item to recent file list

    mnuSep4.Visible = True

    For I = 1 To mnuRecent.count - 1
        If mnuRecent(I).Caption = OpenedFile Then Exit Sub
    Next I
    
    If mnuRecent.count = 5 Then
        Rand = Int((4 * Rnd) + 1)
        mnuRecent(Rand).Caption = OpenedFile
        mnuRecent(Rand).Visible = True
    Else
        Load mnuRecent(mnuRecent.count)
        mnuRecent(mnuRecent.count - 1).Caption = OpenedFile
        mnuRecent(mnuRecent.count - 1).Visible = True
    End If
        
    RefreshSectionList
Canceled:
End Sub


Private Sub txtIni_Change()
Saved = False
End Sub

Private Sub txtIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
    If LastText = "]" Or LastText = "[" Then RefreshSectionList
End If
End Sub

Private Sub txtIni_KeyUp(KeyCode As Integer, Shift As Integer)
Dim FoundPos As Variant
Dim StringToFind As String
If KeyCode = 221 Then
    '"]" key
    RefreshSectionList
End If
End Sub
Public Function AddSlash(Path As String) As String
'On Error Resume Next
If Right(Path, 1) = "\" Or Right(Path, 1) = "/" Then
    AddSlash = Path
Else
    AddSlash = Path & "\"
End If
End Function


Public Function ResizeControls()
On Error Resume Next
YepResize = False
lstSections.Height = Me.Height - (HScrollBar * 5)
txtIni.Width = Me.ScaleWidth - lstSections.Width - 100
txtIni.Height = lstSections.Height 'Me.ScaleHeight - 400
mid.Height = txtIni.Height
split.Height = txtIni.Height
Me.Height = lstSections.Height + (HScrollBar * 5)
YepResize = True
End Function
