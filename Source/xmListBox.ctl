VERSION 5.00
Begin VB.UserControl xmListBox 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   2505
   Begin XmP1.xmSlide ScrollBar 
      Height          =   2175
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3836
   End
End
Attribute VB_Name = "xmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
''''''''''''''''''''''''''''''
'Based on carles_pv@terra.es '
'Sample Code                 '
''''''''''''''''''''''''''''''

'
'Ex Modulo -> Tutto Pubblico ovviamente

Private g_FR As RECT
Private g_FC As RECT

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' DrawFrameControl flags
'Private Const DFC_CAPTION = 1                    'Title bar
'Private Const DFC_MENU = 2                       'Menu
'Private Const DFC_SCROLL = 3                     'Scroll bar
'Private Const DFC_BUTTON = 4                     'Standard button
'
'Private Const DFCS_CAPTIONCLOSE = &H0            'Close button
'Private Const DFCS_CAPTIONMIN = &H1              'Minimize button
'Private Const DFCS_CAPTIONMAX = &H2              'Maximize button
'Private Const DFCS_CAPTIONRESTORE = &H3          'Restore button
'Private Const DFCS_CAPTIONHELP = &H4             'Help button
'
'Private Const DFCS_MENUARROW = &H0               'Submenu arrow
'Private Const DFCS_MENUCHECK = &H1               'Check mark
'Private Const DFCS_MENUBULLET = &H2              'Bullet
'Private Const DFCS_MENUARROWRIGHT = &H4          '
'
'Private Const DFCS_SCROLLUP = &H0                'Up arrow of scroll bar
'Private Const DFCS_SCROLLDOWN = &H1              'Down arrow of scroll bar
'Private Const DFCS_SCROLLLEFT = &H2              'Left arrow of scroll bar
'Private Const DFCS_SCROLLRIGHT = &H3             'Right arrow of scroll bar
'
'Private Const DFCS_SCROLLCOMBOBOX = &H5          'Combo box scroll bar
'Private Const DFCS_SCROLLSIZEGRIP = &H8          'Size grip
'Private Const DFCS_SCROLLSIZEGRIPRIGHT = &H10    'Size grip in bottom-right corner of window
'
'Private Const DFCS_BUTTONCHECK = &H0             'Check box
'Private Const DFCS_BUTTONRADIO = &H4             'Radio button
'Private Const DFCS_BUTTON3STATE = &H8            'Three-state button
'Private Const DFCS_BUTTONPUSH = &H10             'Push button
'Private Const DFCS_INACTIVE = &H100              'Button is inactive (grayed)
'Private Const DFCS_PUSHED = &H200                'Button is pushed
'Private Const DFCS_CHECKED = &H400               'Button is checked
'Private Const DFCS_ADJUSTRECT = &H2000           'Bounding rectangle is adjusted to exclude the surrounding edge of the push button
'Private Const DFCS_FLAT = &H4000                 'Button has a flat border
'Private Const DFCS_MONO = &H8000                 'Button has a monochrome border

' g_list() Item Type

Private Type Item
    Item As String
    Selected As Boolean
End Type



' Enumerated constants

Enum Alignment
    [AlignLeft]
    [AlignCenter]
    [AlignRight]
End Enum

Enum Appearance
    [Flat]
    [3D]
End Enum

Enum BorderStyle
    [None]
    [Fixed Single]
End Enum

Enum OrderType
    [Ascendent]
    [Descendent]
End Enum

Enum SelectMode
    [SingleStandard]
    [SingleGraphical]
    [MultiStandard]
    [MultiGraphical]
End Enum

' Private Variables :

Dim WithEvents g_Font As StdFont
Attribute g_Font.VB_VarHelpID = -1

Dim g_FontColor As OLE_COLOR
Dim g_MousePointer As MousePointerConstants

Dim g_List() As Item
Dim g_Last As Integer
Dim g_VisibleRows As Integer
Dim g_RowHeight As Integer
Dim g_IncTextPos As Integer

Dim g_LeftTab As Integer, g_RightTab As Integer

Dim g_Y As Single
Dim g_SelectModeAppearance As Long

' Default Property Values

Const m_def_Alignment = 0
Const m_def_AlignTab = 30
Const m_def_ListCount = 0
Const m_def_ListIndex = -1
Const m_def_Focus = True
Const m_def_HoverSelection = False
Const m_def_OrderType = 0
Const m_def_ScrollBarWidth = 210
Const m_def_SelectedCount = 0
Const m_def_SelectedUnderlined = False
Const m_def_SelectMode = 0
Const m_def_SelectModeAppearance = 0
Const m_def_TopIndex = -1

' Property Variables

Dim m_Alignment As Alignment
Dim m_AlignTab As Integer
Dim m_Apeareance
Dim m_ListCount As Integer
Dim m_ListIndex As Integer
Dim m_Focus As Boolean
Dim m_HoverSelection As Boolean
Dim m_OrderType As OrderType
Dim m_ScrollBarWidth As Integer
Dim m_SelectedCount As Integer
Dim m_SelectedUnderlined As Boolean
Dim m_SelectMode As SelectMode
Dim m_SelectModeAppearance As Appearance
Dim m_TopIndex As Integer

' Event Declarations

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ListIndexChange()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Scroll()
Event TopIndexChange()


' ___________________________________________________________________________
' I n i t / R e a d / W r i t e   p r o p e r t i e s :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub UserControl_InitProperties()

    m_Alignment = m_def_Alignment
    m_AlignTab = m_def_AlignTab
    m_Focus = m_def_Focus
    m_HoverSelection = m_def_HoverSelection
    m_ListCount = m_def_ListCount
    m_ListIndex = m_def_ListIndex
    m_OrderType = m_def_OrderType
    m_ScrollBarWidth = m_def_ScrollBarWidth
    m_SelectedCount = m_def_SelectedCount
    m_SelectedUnderlined = m_def_SelectedUnderlined
    m_SelectMode = m_def_SelectMode
    m_SelectModeAppearance = m_def_SelectModeAppearance
    m_TopIndex = m_def_TopIndex
    
    Set UserControl.Font = Ambient.Font
    Set g_Font = Ambient.Font
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_AlignTab = PropBag.ReadProperty("AlignTab", m_def_AlignTab)
    m_Focus = PropBag.ReadProperty("Focus", m_def_Focus)
    m_HoverSelection = PropBag.ReadProperty("HoverSelection", m_def_HoverSelection)
    m_ListCount = PropBag.ReadProperty("ListCount", m_def_ListCount)
    m_ListIndex = PropBag.ReadProperty("ListIndex", m_def_ListIndex)
    m_ScrollBarWidth = PropBag.ReadProperty("ScrollBarWidth", m_def_ScrollBarWidth)
    m_SelectedCount = PropBag.ReadProperty("SelectedCount", m_def_SelectedCount)
    m_SelectedUnderlined = PropBag.ReadProperty("SelectedUnderlined", m_def_SelectedUnderlined)
    m_SelectMode = PropBag.ReadProperty("SelectMode", m_def_SelectMode)
    m_SelectModeAppearance = PropBag.ReadProperty("SelectModeAppearance", m_def_SelectModeAppearance)
    m_OrderType = PropBag.ReadProperty("OrderType", m_def_OrderType)
    m_TopIndex = PropBag.ReadProperty("TopIndex", m_def_TopIndex)
    
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", Verdadero)
    UserControl.ForeColor = PropBag.ReadProperty("FontColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    g_FontColor = UserControl.ForeColor
    g_MousePointer = UserControl.MousePointer
    g_SelectModeAppearance = (1 - m_SelectModeAppearance) * DFCS_FLAT
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("AlignTab", m_AlignTab, m_def_AlignTab)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, Verdadero)
    Call PropBag.WriteProperty("Focus", m_Focus, m_def_Focus)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
    Call PropBag.WriteProperty("ListCount", m_ListCount, m_def_ListCount)
    Call PropBag.WriteProperty("ListIndex", m_ListIndex, m_def_ListIndex)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OrderType", m_OrderType, m_def_OrderType)
    Call PropBag.WriteProperty("ScrollBarWidth", m_ScrollBarWidth, m_def_ScrollBarWidth)
    Call PropBag.WriteProperty("SelectedCount", m_SelectedCount, m_def_SelectedCount)
    Call PropBag.WriteProperty("SelectedUnderlined", m_SelectedUnderlined, m_def_SelectedUnderlined)
    Call PropBag.WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
    Call PropBag.WriteProperty("SelectModeAppearance", m_SelectModeAppearance, m_def_SelectModeAppearance)
    Call PropBag.WriteProperty("TopIndex", m_TopIndex, m_def_TopIndex)
    
End Sub


' ___________________________________________________________________________
' U s e r C o n t r o l :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub UserControl_Initialize()
            
    'UserControl.KeyPreview = False

    ScrollBar.Visible = False
    ScrollBar.xMax = 0
    g_RightTab = 0
    
    Erase g_List
    ReDim g_List(0)
    g_List(0).Item = "<Nothing>"
    g_Last = -1
    g_Y = -1
    ListIndex = -1
    
    Set g_Font = New StdFont
    g_Font = GetINI(cfgFile, "VISUALIZATION", "FONTNAME", "TAHOMA")
    
    Cls
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    g_RowHeight = TextHeight("Font")
    
    If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then
    
        g_IncTextPos = -15
        
        If g_RowHeight < 225 Then
            g_IncTextPos = (225 - TextHeight("Font")) / 2
            If g_IncTextPos Mod 2 <> 0 Then g_IncTextPos = g_IncTextPos - 15
            g_RowHeight = 225
        End If
        
    End If
    
    ' Get visible rows a readjust UserControl.Height
    g_VisibleRows = Int(ScaleHeight / g_RowHeight)
    UserControl.Height = (g_VisibleRows) * g_RowHeight + (UserControl.Height - ScaleHeight)
    
    ' Locate & resize ScrollBar
    ScrollBar.Move ScaleWidth - ScrollBar.Width, 0, ScrollBar.Width, ScaleHeight
    
    ' Define g_LeftTab
    If SelectMode = SingleGraphical Or SelectMode = MultiGraphical Then g_LeftTab = 225 Else g_LeftTab = 0
    
End Sub

Private Sub UserControl_Paint()

    If Not Ambient.UserMode Then
        
        Cls
        
        If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then
            
            DrawBoxChecked True
            
            CurrentX = 225 + AlignTab
            
            If TextHeight(Ambient.DisplayName) < 225 Then
                CurrentY = (225 - TextHeight(Ambient.DisplayName)) / 2
                If CurrentY Mod 2 <> 0 Then CurrentY = CurrentY - 15
            End If
            
        Else
        
            Select Case Alignment
                Case 0: CurrentX = AlignTab
                Case 1: CurrentX = (ScaleWidth - TextWidth(Ambient.DisplayName)) / 2
                Case 2: CurrentX = (ScaleWidth - TextWidth(Ambient.DisplayName) - AlignTab)
            End Select
            
        End If
        
        Print Ambient.DisplayName
        
    Else
                    
        DrawList
        
    End If
    
End Sub

Private Sub UserControl_Terminate()

    Erase g_List
    
End Sub

' ___________________________________________________________________________
' S c r o l l B a r :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub ScrollBar_Change()
    
    RaiseEvent Scroll
    If ListIndex = g_Last Then DrawList
    
End Sub

Private Sub ScrollBar_Scroll()

    RaiseEvent Scroll
    ScrollBar_Change
    
End Sub

' Value *********************************************************
Public Property Get ScrollValue() As Long

    ScrollValue = ScrollBar.xValue
    
End Property

Public Property Let ScrollValue(ByVal Valore As Long)
   
    ScrollBar.xValue = Valore
    
End Property
' Max *********************************************************
Public Property Get ScrollMax() As Long

    ScrollMax = ScrollBar.xMax
    
End Property

' ___________________________________________________________________________
' U s e r C o n t r o l   P r o p e r t i e s :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' A l i g n m e n t *********************************************************
Public Property Get Alignment() As Alignment

    Alignment = m_Alignment
    
End Property

Public Property Let Alignment(ByVal New_Alignment As Alignment)

    If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then Exit Property
    
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    
    Refresh
    
End Property

' A l i g n T a b ***********************************************************
Public Property Get AlignTab() As Integer

    AlignTab = m_AlignTab
    
End Property

Public Property Let AlignTab(ByVal New_AlignTab As Integer)

    If New_AlignTab > ScaleWidth / 2 Then
        m_AlignTab = ScaleWidth / 2
    Else
        m_AlignTab = New_AlignTab
    End If
    PropertyChanged "AlignTab"
    
    Refresh
    
End Property

' A p p e a r a n c e *******************************************************
Public Property Get Appearance() As Appearance

    Appearance = UserControl.Appearance
    
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearance)

    UserControl.Appearance() = New_Appearance
    BackColor = vbWindowBackground
    PropertyChanged "Appearance"
    
End Property

' B a c k C o l o r *********************************************************
Public Property Get BackColor() As OLE_COLOR

    BackColor = UserControl.BackColor
    
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
     
    Refresh
    
End Property

' B o r d e r S t y l e *****************************************************
Public Property Get BorderStyle() As BorderStyle

    BorderStyle = UserControl.BorderStyle
    
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)

    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    
End Property

' E n a b l e d *************************************************************
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"

    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
End Property

' F o c u s *****************************************************************
Public Property Get Focus() As Boolean

    Focus = m_Focus
    
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)

    m_Focus = New_Focus
    PropertyChanged "Focus"
    
    If New_Focus Then DrawFocus ListIndex Else DrawItem ListIndex
    
End Property

' F o n t *******************************************************************
Public Property Get Font() As Font

    Set Font = g_Font
    
End Property

Public Property Set Font(ByVal New_Font As Font)

    With g_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline ' Not operate
        .Strikethrough = New_Font.Strikethrough
    End With
    PropertyChanged "Font"
    
End Property

Private Sub g_Font_FontChanged(ByVal PropertyName As String)

    Set UserControl.Font = g_Font
    
    UserControl_Resize
    Refresh
    
End Sub

' F o n t C o l o r *********************************************************
Public Property Get FontColor() As OLE_COLOR

    FontColor = UserControl.ForeColor
    
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)

    UserControl.ForeColor() = New_FontColor
    g_FontColor = New_FontColor
    PropertyChanged "FontColor"
    
    Refresh
    
End Property

' H o v e r S e l e c t i o n ***********************************************
Public Property Get HoverSelection() As Boolean

    HoverSelection = m_HoverSelection
    
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    
    m_HoverSelection = New_HoverSelection
    PropertyChanged "HoverSelection"
    
    If New_HoverSelection = True Then
        MousePointer = vbUpArrow
        If (SelectMode = SingleStandard Or SelectMode = SingleGraphical) And ListIndex <> -1 Then
            DrawItem ListIndex
            DrawFocus ListIndex
        End If
    Else
        MousePointer = g_MousePointer
        If ListIndex <> -1 Then DrawFocus ListIndex
    End If
    
End Property

' L i s t C o u n t *********************************************************
Public Property Get ListCount() As Integer

    ListCount = UBound(g_List)
    
End Property

' L i s t I n d e x *********************************************************
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"

    ListIndex = m_ListIndex
    
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    
    If New_ListIndex < -1 Or New_ListIndex > UBound(g_List) - 1 Then Err.Raise 380
    
    If New_ListIndex < 0 Or UBound(g_List) = 0 Then
        m_ListIndex = -1
        g_Y = -1
    Else
        m_ListIndex = New_ListIndex
    End If
    PropertyChanged "ListIndex"
    
    RaiseEvent ListIndexChange
    
    ' Unselect last / Select actual (SingleStandard/SingleGraphical modes)
    If SelectMode = SingleStandard Or SelectMode = SingleGraphical Then
        If g_Last > -1 Then g_List(g_Last).Selected = False
        If m_ListIndex > -1 Then g_List(m_ListIndex).Selected = True
    End If
    
    ' Draw last (delete Focus) ...
    DrawItem g_Last
    g_Last = m_ListIndex
    ' ... and draw actual (draw Focus)
    DrawItem m_ListIndex
    DrawFocus m_ListIndex

    ' TopIndex
        
    If m_ListIndex < ScrollBar.xValue And m_ListIndex >= 0 Then
        ScrollBar.xValue = m_ListIndex
    ElseIf m_ListIndex >= ScrollBar.xValue + g_VisibleRows Then
        ScrollBar.xValue = m_ListIndex - g_VisibleRows + 1
    End If

End Property

' M o u s e I c o n *********************************************************
Public Property Get MouseIcon() As Picture

    Set MouseIcon = UserControl.MouseIcon
    
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)

    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
    
End Property

' M o u s e P o i n t e r ***************************************************
Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer
    
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    
End Property

' O r d e r T y p e *********************************************************
Public Property Get OrderType() As OrderType

    OrderType = m_OrderType
    
End Property

Public Property Let OrderType(ByVal New_OrderType As OrderType)

    m_OrderType = New_OrderType
    PropertyChanged "OrderType"
    
End Property

' S c r o l l B a r W i d t h ***********************************************
Public Property Get ScrollBarWidth() As Integer

    ScrollBarWidth = m_ScrollBarWidth
    
End Property

Public Property Let ScrollBarWidth(ByVal New_ScrollBarWidth As Integer)

    If New_ScrollBarWidth < 150 Then                ' Min value width
        m_ScrollBarWidth = 150
        ScrollBar.Width = 150
    ElseIf New_ScrollBarWidth > ScaleWidth / 2 Then ' Max value width
        m_ScrollBarWidth = ScaleWidth / 2
        ScrollBar.Width = ScaleWidth / 2
    Else
        m_ScrollBarWidth = New_ScrollBarWidth
        ScrollBar.Width = New_ScrollBarWidth
    End If
    PropertyChanged "ScrollBarWidth"
    
    UserControl_Resize
    
    ScrollBar.Visible = False
    ReadjustScrollBar
    
End Property

' S e l e c t e d C o u n t *************************************************
Public Property Get SelectedCount() As Integer
    
    Dim I As Integer
    
    SelectedCount = 0
    For I = 0 To UBound(g_List)
        If g_List(I).Selected Then SelectedCount = SelectedCount + 1
    Next I
    
End Property

' S e l e c t e d U n d e r l i n e d ***************************************
Public Property Get SelectedUnderlined() As Boolean
Attribute SelectedUnderlined.VB_ProcData.VB_Invoke_Property = "General"

    SelectedUnderlined = m_SelectedUnderlined
    
End Property

Public Property Let SelectedUnderlined(ByVal New_SelectedUnderlined As Boolean)

    m_SelectedUnderlined = New_SelectedUnderlined
    PropertyChanged "SelectedUnderlined"
    
    Refresh
    
End Property

' S e l e c t M o d e *******************************************************
Public Property Get SelectMode() As SelectMode

    SelectMode = m_SelectMode
    
End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectMode)
    
    m_SelectMode = New_SelectMode
    PropertyChanged "SelectMode"
   
    UserControl_Resize
    
    If New_SelectMode = SingleGraphical Or New_SelectMode = MultiGraphical Then
        Font.Underline = False
        m_Alignment = AlignLeft
    End If
    
    If Ambient.UserMode Then
    
        UserControl.ForeColor = g_FontColor
        
        If New_SelectMode = SingleStandard Or New_SelectMode = SingleGraphical Then
        
            Dim I As Integer
            
            If ListIndex <> -1 Then
                For I = LBound(g_List) To UBound(g_List)
                   If I <> ListIndex Then g_List(I).Selected = False
                Next I
                g_List(ListIndex).Selected = True
                DrawItem ListIndex
                DrawFocus ListIndex
            End If
            
        End If
        
   End If
   
   ReadjustScrollBar
   
   Refresh
    
End Property

' S e l e c t M o d e A p p e a r a n c e ***********************************
Public Property Get SelectModeAppearance() As Appearance

    SelectModeAppearance = m_SelectModeAppearance
    
End Property

Public Property Let SelectModeAppearance(ByVal New_SelectModeAppearance As Appearance)
    
    m_SelectModeAppearance = New_SelectModeAppearance
    PropertyChanged "SelectModeAppearance"
    
    g_SelectModeAppearance = (1 - m_SelectModeAppearance) * DFCS_FLAT
    
    Refresh
    
End Property

' T o p I n d e x ***********************************************************
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"

    TopIndex = ScrollBar.xMax
    
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)

    If New_TopIndex < 0 Or New_TopIndex > ScrollBar.xValue Then Err.Raise 380

    m_TopIndex = New_TopIndex
    PropertyChanged "TopIndex"
    
    RaiseEvent TopIndexChange
    
    ScrollBar.xValue = New_TopIndex
        
End Property


' ___________________________________________________________________________
' U s e r C o n t r o l   E v e n t s :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' C l i c k *****************************************************************
Private Sub UserControl_Click()
  
    If ListIndex <> -1 Then RaiseEvent Click
    
End Sub

' D b l C l i c k ***********************************************************
Private Sub UserControl_DblClick()

    If ListIndex <> -1 Then RaiseEvent DblClick
    
End Sub

' K e y D o w n *************************************************************
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If UBound(g_List) = 0 Or ListIndex = -1 Then Exit Sub

    Select Case KeyCode
    
        Case 38 'Up arrow
            If ListIndex > 0 Then ListIndex = ListIndex - 1
            
        Case 40 'Down arrow
            If ListIndex < UBound(g_List) - 1 Then ListIndex = ListIndex + 1
            
        Case 33 'PageDown
            If ListIndex > g_VisibleRows Then
                ListIndex = ListIndex - g_VisibleRows
            Else
                ListIndex = 0
            End If
            
        Case 34 'PageUp
            If ListIndex < UBound(g_List) - g_VisibleRows - 1 Then
                ListIndex = ListIndex + g_VisibleRows
            Else
                ListIndex = UBound(g_List) - 1
            End If
            
        Case 36 'Start
            ListIndex = 0
            
        Case 35 'End
            ListIndex = UBound(g_List) - 1
            
        Case 32 'Space: select on/off
            If SelectMode <> SingleStandard And SelectMode <> SingleGraphical And ListIndex <> -1 Then
                g_List(ListIndex).Selected = Not g_List(ListIndex).Selected
                DrawItem ListIndex
                DrawFocus ListIndex
            End If
            
        Case Else
            
            Exit Sub
            
    End Select
        
    If ListIndex < TopIndex Then
        ScrollBar.xValue = ListIndex
    ElseIf ListIndex > TopIndex + g_VisibleRows - 1 Then
        ScrollBar.xValue = ListIndex - g_VisibleRows + 1
    End If
    
End Sub

' K e y P r e s s ***********************************************************
Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
    
End Sub

' K e y U p *****************************************************************
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub

' M o u s e D o w n *********************************************************
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim SelectedItemIndex As Integer
    SelectedItemIndex = ScrollBar.xValue + Int(Y / g_RowHeight)
    
    If SelectedItemIndex < UBound(g_List) Then
    
        Select Case SelectMode
        
            Case 0, 1 'SingleStandard & SingleGraphical
                g_List(SelectedItemIndex).Selected = True
                
            Case 2, 3 'MultiStandard & MultiGraphical
                g_List(SelectedItemIndex).Selected = Not g_List(SelectedItemIndex).Selected
                
        End Select
        
        g_Y = Y
        
        ListIndex = SelectedItemIndex
        
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

' M o u s e M o v e *********************************************************
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Y < 0 Or Y > ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        Exit Sub
    End If
    
    If HoverSelection And Int(Y / g_RowHeight) <> Int(g_Y / g_RowHeight) Then
    
        If Int(Y / g_RowHeight) < UBound(g_List) Then
            UserControl_MouseDown vbLeftButton, Shift, X, Y
        End If
        
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

' M o u s e U p *************************************************************
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

' ___________________________________________________________________________
' U s e r C o n t r o l   M e t h o d s :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' A d d *********************************************************************
Public Sub AddItem(Item) '    0, ..., n-1. n = ListCount

        g_List(UBound(g_List)).Item = ElaboraNomeFile(CStr(Item))
    
        ReDim Preserve g_List(0 To UBound(g_List) + 1)
    
    ReadjustScrollBar
    
    If UBound(g_List) < g_VisibleRows + 1 Then DrawItem UBound(g_List) - 1
    
End Sub

' I n s e r t ***************************************************************
Public Sub InsertItem(Index As Integer, Item)
     
    If UBound(g_List) = 0 Or _
       Index > UBound(g_List) Then Err.Raise 381 ' Empty List or "Item added"
    
        ReDim Preserve g_List(UBound(g_List) + 1)
    
        Dim I As Integer
        
        For I = UBound(g_List) - 1 To Index Step -1
           g_List(I + 1) = g_List(I)
        Next I
         
        g_List(Index).Item = CStr(Item)
        g_List(Index).Selected = False
    
    ReadjustScrollBar
    
    If ListIndex <> -1 Then ListIndex = ListIndex + 1
    
    Refresh
    
End Sub

' M o d i f y ***************************************************************
Public Sub ModifyItem(Index As Integer, NewItem, Optional SelectItem As Boolean)
    
    If UBound(g_List) = 0 Then Err.Raise 381 ' Empty List
    
        g_List(Index).Item = CStr(NewItem)
    
    DrawItem Index
    DrawFocus ListIndex

End Sub

' R e m o v e ***************************************************************
Public Sub RemoveItem(Index As Integer)

    If UBound(g_List) = 0 Or Index > UBound(g_List) - 1 Then Err.Raise 381 ' Empty List
        
        If Index < UBound(g_List) - 1 Then
            Dim I As Integer
            For I = Index To UBound(g_List) - 1
               g_List(I) = g_List(I + 1)
            Next I
        End If
         
        ReDim Preserve g_List(UBound(g_List) - 1)
        
        ReadjustScrollBar
    
        If Index < ListIndex Then
            If ListIndex <> -1 Then ListIndex = ListIndex - 1
        ElseIf Index = ListIndex Then
            ListIndex = -1
        End If
    
    Refresh
    
End Sub

' G e t I t e m *************************************************************
Public Function GetItem(Index As Integer) As String
    
    If Index = -1 Then Err.Raise 381
    GetItem = g_List(Index).Item
    
End Function

' G e t L i s t *************************************************************
'Public Function GetList() As Item()
    
'    If UBound(g_List) > 0 Then ReDim Preserve g_List(UBound(g_List) - 1)
'    GetList = g_List

'End Function

' I s S e l e c t e d *******************************************************
Public Function IsSelected(Index As Integer) As Boolean

    If Index = -1 Then Err.Raise 381
    IsSelected = g_List(Index).Selected
   
End Function

' F i n d F i r s t *********************************************************
Public Function FindFirst(fndString As String, StartIndex As Integer, StartWith As Boolean) As Integer
    
    If UBound(g_List) = 0 Then Err.Raise 2, , "Empty list": Exit Function
    
    Dim I As Integer
    For I = StartIndex To UBound(g_List)
        If StartWith Then
            If InStr(1, UCase$(g_List(I).Item), UCase$(fndString)) = 1 Then FindFirst = I: Exit Function
        Else
            If InStr(1, UCase$(g_List(I).Item), UCase$(fndString)) > 1 Then FindFirst = I: Exit Function
        End If
    Next I
    
    FindFirst = -1 ' fndString not founded

End Function

' C l e a r *****************************************************************
Public Sub Clear()
      
    Erase g_List
    ReDim g_List(0)
    g_List(0).Item = "<Nothing>"
    g_Last = -1
    ListIndex = -1
    
    ScrollBar.Visible = False
    ScrollBar.xMax = 0
    g_RightTab = 0
       
    Cls
        
End Sub

' O r d e r *****************************************************************
Public Sub Order()

    If UBound(g_List) < 2 Then Exit Sub
    
    If SelectMode = SingleStandard Or SelectMode = SingleGraphical Then
        If ListIndex <> -1 Then g_List(ListIndex).Selected = False
    End If
    
    Dim Index As Integer, Index2 As Integer
    Dim FirstItem As Integer, NumberOfItems As Integer
    Dim Distance As Integer, Value As Item
    Dim Desc As Boolean
    
    FirstItem = LBound(g_List)
    NumberOfItems = UBound(g_List)
    If OrderType = Descendent Then Desc = True
    
    Do
        Distance = Distance * 3 + 1
    Loop Until Distance > NumberOfItems
    
    Do
        Distance = Distance \ 3
        
        For Index = Distance + FirstItem To NumberOfItems + FirstItem - 1
        
            Value = g_List(Index)
            Index2 = Index
            
            Do While (g_List(Index2 - Distance).Item > Value.Item) Xor Desc
                g_List(Index2) = g_List(Index2 - Distance)
                Index2 = Index2 - Distance
                If Index2 - Distance < FirstItem Then Exit Do
            Loop
            
            g_List(Index2) = Value
            
        Next
        
    Loop Until Distance = 1
    
    ListIndex = -1
    ScrollBar.xValue = 0
    
    Refresh
    
End Sub

' R e f r e s h *************************************************************
Public Sub Refresh() 'UserControl

    UserControl.Refresh
    
End Sub

' S e l e c t I t e m  /  U n s e l e c t I t e m ***************************
Public Sub SelectItem(Index As Integer)
    
    If SelectMode = SingleStandard Or SelectMode = SingleGraphical Then
        ListIndex = Index
    Else
        g_List(Index).Selected = True
        DrawItem Index
        If Index = ListIndex Then DrawFocus Index
    End If
    
End Sub

Public Sub UnselectItem(Index As Integer)
    
    If SelectMode = SingleStandard Or SelectMode = SingleGraphical Then
        ' Nothing to do ...
    Else
        g_List(Index).Selected = False
        DrawItem Index
        If Index = ListIndex Then DrawFocus Index
    End If

End Sub


' ___________________________________________________________________________
' P r i v a t e   S u b s :
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' D r a w L i s t ***********************************************************
Private Sub DrawList()
    
    If Not Extender.Visible Then Exit Sub

    ' Fill visible List
    Dim I As Integer
    
        For I = ScrollBar.xValue To ScrollBar.xValue + g_VisibleRows
            If I = UBound(g_List) Then Exit For
            DrawItem I
            CurrentY = CurrentY + g_RowHeight
        Next I
    
    ' Draw focus
    DrawFocus ListIndex
    
End Sub

' D r a w B o x *************************************************************
Private Sub DrawBoxChecked(Selected As Boolean)
    
    ' Set position & dimensions
    g_FC.Top = (CurrentY / 15) + ((g_RowHeight - 15) / 15 - 14) / 2 + 1
    g_FC.Bottom = g_FC.Top + 13
    g_FC.Left = 1
    g_FC.Right = g_FC.Left + 13
    
    ' Draw it
    If Selected Then
        If SelectMode = MultiGraphical Then
            DrawFrameControl hdc, g_FC, DFC_BUTTON, DFCS_BUTTONCHECK Or g_SelectModeAppearance Or DFCS_CHECKED
        Else
            DrawFrameControl hdc, g_FC, DFC_BUTTON, DFCS_BUTTONRADIO Or g_SelectModeAppearance Or DFCS_CHECKED
        End If
    Else
        If SelectMode = MultiGraphical Then
            DrawFrameControl hdc, g_FC, DFC_BUTTON, DFCS_BUTTONCHECK Or g_SelectModeAppearance
        Else
            DrawFrameControl hdc, g_FC, DFC_BUTTON, DFCS_BUTTONRADIO Or g_SelectModeAppearance
        End If
    End If
    
End Sub

' D r a w I t e m ***********************************************************
Private Sub DrawItem(Index As Integer)

    RaiseEvent TopIndexChange

    ' Item out of area ?
    If Index < ScrollBar.xValue Or _
       Index > ScrollBar.xValue + g_VisibleRows Then Exit Sub  'Item out of area
    
        ' Text CurrentY
        CurrentY = (Index - ScrollBar.xValue) * g_RowHeight
        
        ' Selected Item ?
        If g_List(Index).Selected Then
        
            If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then
                Line (ScaleWidth, CurrentY + g_RowHeight - 15)-(g_LeftTab, CurrentY), BackColor, BF
                DrawBoxChecked True
            ElseIf SelectedUnderlined Then
                Line (ScaleWidth, CurrentY + g_RowHeight - 15)-(0, CurrentY), BackColor, BF
                UserControl.Font.Underline = True
            Else
                Line (ScaleWidth, CurrentY + g_RowHeight - 15)-(0, CurrentY), BackColor, BF
            End If
            
        Else
        
            Line (ScaleWidth, CurrentY + g_RowHeight - 15)-(g_LeftTab, CurrentY), BackColor, BF
            
            If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then DrawBoxChecked False
            
            UserControl.Font.Underline = False
    
        End If
        
        ' If text height < graph selection height : center text on graph selection
        If SelectMode = MultiGraphical Or SelectMode = SingleGraphical Then CurrentY = CurrentY + g_IncTextPos

        ' Text Alignment CurrentX
        Select Case Alignment
        
            Case 0: CurrentX = g_LeftTab + AlignTab
            Case 1: CurrentX = (ScaleWidth - TextWidth(g_List(Index).Item) - g_RightTab) / 2
            Case 2: CurrentX = (ScaleWidth - TextWidth(g_List(Index).Item) - g_RightTab - AlignTab)
        
        End Select
        
        ' Draw Item
        If g_List(Index).Selected And SelectMode <> MultiGraphical And SelectMode <> SingleGraphical And SelectedUnderlined = False Then
            UserControl.ForeColor = vbHighlightText
        Else
            UserControl.ForeColor = g_FontColor
        End If
        
        TextOut hdc, CurrentX / 15, CurrentY / 15, g_List(Index).Item, Len(g_List(Index).Item)
        
End Sub

' D r a w R e c t ***********************************************************
Private Sub DrawFocus(Index As Integer)
    
    If Not Focus Then Exit Sub
    
    ' Item out of area ?
    If Index < ScrollBar.xValue Or _
       Index > ScrollBar.xValue + g_VisibleRows Then Exit Sub  'Item out of area
       
        ' Set position & dimensions
        g_FR.Top = (Index - ScrollBar.xValue) * g_RowHeight / 15
        g_FR.Bottom = g_FR.Top + g_RowHeight / 15
        g_FR.Left = g_LeftTab / 15
        g_FR.Right = (ScaleWidth - g_RightTab) / 15
    
        ' Draw it
        UserControl.ForeColor = g_FontColor
        DrawFocusRect hdc, g_FR

End Sub

' R e a d j u s t S c r o l l B a r *****************************************
Private Sub ReadjustScrollBar()
     
    If UBound(g_List) > g_VisibleRows Then
     
        If Not ScrollBar.Visible Then
        
            ScrollBar.Visible = True
            'ScrollBar.LargeChange = g_VisibleRows
            g_RightTab = ScrollBar.Width
            DrawList
            
        End If
        
    Else
     
       ScrollBar.Visible = False
       g_RightTab = 0
       
    End If
     
    ScrollBar.xMax = UBound(g_List) - g_VisibleRows

End Sub


