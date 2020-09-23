Attribute VB_Name = "modFormMinMaxResize"
Option Explicit


' Private types
Private Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

' min/max form sizes
Private Const MIN_WIDTH = 10
Private Const MAX_WIDTH = 310
Private Const MIN_HEIGHT = 10
Private Const MAX_HEIGHT = 10000

' private consts
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WINDOWPOSCHANGED = &H47

' handle to the old win proc
Public OldWindowProc As Long

' public consts
Public Const GWL_WNDPROC = (-4)

' api declares
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


' sub class procdure
Public Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
    ' check if window resizing
    If msg = WM_WINDOWPOSCHANGING Then
        ' keep window size within limits
        If lParam.cx < MIN_WIDTH Then lParam.cx = MIN_WIDTH
        If lParam.cx > MAX_WIDTH Then lParam.cx = MAX_WIDTH
        If lParam.cy < MIN_HEIGHT Then lParam.cy = MIN_HEIGHT
        If lParam.cy > MAX_HEIGHT Then lParam.cy = MAX_HEIGHT
    End If
    
    ' Continue processing !DON'T REMOVE THIS LINE!
    WindowProc = CallWindowProc( _
        OldWindowProc, hwnd, msg, wParam, _
        lParam)
End Function
