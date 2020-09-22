Attribute VB_Name = "modToolTip"
Option Explicit
'-------------------------------------------------------------------------
'This module is needed because it provides a WinProc used for subclassing.  Tooltips
'must be provided in this manner because VB5's tooltips for an intrinsic
'control does not work with the SetCapture API in use.  Also, VB's provided
'tooltip is container provided.  Therefore, if a control is used in a
'container that does not provide a tooltiptext property on the extender
'object, the tooltip would not be provided.  The tooltip in this control is
'provided regardless of the container hosting it.  The
'set capture is necessary to know precisely when the mouse moves over the control
'and then back off the control.
'-------------------------------------------------------------------------
Public gHWndToolTip As Long                 'Hwnd of Tooltip created by this object
Public gbToolTipsInstanciated As Boolean    'If true the ToolTip class window has been created
Public glToolsCount As Long                 'The number of controls using tool tips
                                            
#If DEBUGSUBCLASS Then
    Public goWindowProcHookCreator As Object
#End If

Public Function SubWndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '-------------------------------------------------------------------------
    'Purpose:   Address used for subclassing.  Calls instance of uctlSoftButton whose hWnd is stored in USERDATA of window matching passed hWnd
    '-------------------------------------------------------------------------
    On Error Resume Next
    SubWndProc = ConvertUserDataToButton(hwnd).WindowProc(hwnd, MSG, wParam, lParam)
End Function

Private Function ConvertUserDataToButton(hwnd As Long) As Button
    '-------------------------------------------------------------------------
    'Purpose:   Gets the hWnd of a uctlSoftButton object, and converts it to a reference to the uctlSoftButton object without increasing
    '                   VB's ref count of that object
    '-------------------------------------------------------------------------
    Dim Obj As Button
    Dim pObj As Long
    pObj = GetWindowLong(hwnd, GWL_USERDATA)
    CopyMemory Obj, pObj, 4
    Set ConvertUserDataToButton = Obj
    CopyMemory Obj, 0&, 4
End Function

