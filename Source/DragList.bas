Attribute VB_Name = "mDragList"
'drg DragList
'
'(c) 1999 Joe Hart
'
'This is a subclassed listbox control that allows dragging 'items from one location to another
'
'Thanks to Eduardo Morcillo (E-Mail: edanmo@geocities.com) for providing 'the program
'of which this control is based.
'
'You MAY use this control in your own projects. Be sure to also include the 'drgDragList.ctl
'file that accompanies this module.
'
'If you compile this contol into an OCX, please RENAME IT before you distribute it
'due to version compatibilty issues.
'
'If you make modifications to this draglist, please send them to me at bghost@ti.cz

'This is the .BAS module for the drgDraglist.  It is needed for the subclassing.

Option Explicit

'API Declarations used for subclassing.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Constants for GetWindowLong() and SetWindowLong() API's.
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

'used as the variable to hold the pointer to the draglist
Private ctlShadowControl As DragList

'Used as a pointer to the UserData section of a window.
Dim ptrObject As Long

'The address of this function is used for subclassing.
'Messages will be sent here and then forwarded to the
'UserControl's WindowProc function. The HWND determines
'to which control the message is sent.
Public Function SubWinProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   On Error Resume Next

   'Get pointer to the control's VTable from the
   'window's UserData section. The VTable is an internal
   'structure that contains pointers to the methods and
   'properties of the control.
   ptrObject = GetWindowLong(hWnd, GWL_USERDATA)

   'Copy the memory that points to the VTable of our original
   'control to the shadow copy of the control you use to
   'call the original control's WindowProc Function.
   'This way, when you call the method of the shadow control,
   'you are actually calling the original controls' method.
   CopyMemory ctlShadowControl, ptrObject, 4
   
   'Call the WindowProc function in the instance of the UserControl.
   SubWinProc = ctlShadowControl.WinProc(hWnd, msg, wParam, lParam)
   
   'Destroy the Shadow Control Copy
   CopyMemory ctlShadowControl, 0&, 4
   Set ctlShadowControl = Nothing
End Function



