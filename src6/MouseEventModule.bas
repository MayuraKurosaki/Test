Attribute VB_Name = "MouseEventModule"
Option Explicit
Option Private Module
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
'  (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Const WM_MOUSEWHEEL = &H20A
Const MK_CONTROL = &H8
Const MK_SHIFT = &H4
Const WM_DESTROY = &H2
Const WM_DROPFILES = &H233
Public gparam As Collection
Private dform As MouseEventForm

'メッセージフック
Function FormHook(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    On Error Resume Next
    Set dform = gparam(CStr(hwnd))
    Select Case uMsg
        Case WM_MOUSEWHEEL
            Debug.Print "FormHook:MouseWheel"
            dform.EventMouseWheel dform.Control, wParam, IIf(wParam And MK_SHIFT, 1, 0) + IIf(wParam And MK_CONTROL, 2, 0)
        Case WM_DROPFILES
            Debug.Print "FormHook:DropFiles"
            dform.EventDropFiles hwnd, wParam
            Exit Function
        Case WM_DESTROY
            dform.Terminate
    End Select
    FormHook = CallWindowProc(dform.Hook, hwnd, uMsg, wParam, lParam)
End Function

