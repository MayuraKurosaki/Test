Attribute VB_Name = "MouseWheelModule"
Option Explicit

Private Const WH_MOUSE_LL          As Long = 14
Private Const WH_MOUSE              As Long = 7
Private Const WM_MOUSEWHEEL        As Long = &H20A
Private Const HC_ACTION            As Long = 0
Private Const GWL_HINSTANCE        As Long = (-6)
Private Const WHEEL_DOWN           As LongPtr = 7864320
Private Const WHEEL_UP             As LongPtr = -7864320
Private Const WM_DROPFILES = &H233

'Private Const WHEEL_UP             As LongPtr = 4287102976#

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long

#If Win64 Then
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
#Else
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
#End If

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As LongPtr, ByVal fAccept As Long)

'Private Type POINTAPI
'    XY As LongPtr
'End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As LongPtr
    wHitTestCode As Long '????
    dwExtraInfo As LongPtr
End Type
 
Private HookPtr As LongPtr, EventControl As Object, EventPtr As LongPtr
 
'------------------
'Hook, Proc, UnHook
'------------------
 
Public Sub HookControl(NewEventControl As Object)
    If HookPtr = 0 Then
       Set EventControl = NewEventControl
'       EventControl.BackColor = vbRed 'Test
       EventPtr = CurserPtr
       HookPtr = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, FormPtr, 0)
'       HookPtr = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, FormPtr, 0)
    End If
End Sub
 
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
    On Error GoTo 1
    MouseProc = CallNextHookEx(HookPtr, nCode, wParam, ByVal lParam)

    Dim WheelScrool As Variant
    Select Case True
           Case EventControl Is Nothing: UnHookControl
           Case EventPtr <> CurserPtr: UnHookControl
           Case HookPtr = 0
           Case nCode <> HC_ACTION
           Case wParam <> WM_MOUSEWHEEL
           Case lParam.hwnd = WHEEL_DOWN: WheelScrool = EventControl.ListIndex - 1
           Case lParam.hwnd = WHEEL_UP:   WheelScrool = EventControl.ListIndex + 1
           Case wParam = WM_DROPFILES:  Debug.Print "DropFile"
    End Select
    If Not IsEmpty(WheelScrool) Then
       WheelScrool = IIf(WheelScrool < 0, 0, WheelScrool)
       WheelScrool = IIf(WheelScrool > EventControl.ListCount - 1, EventControl.ListCount - 1, WheelScrool)
'       If EventControl.BackColor <> vbYellow Then EventControl.BackColor = vbYellow 'Test
       EventControl.ListIndex = WheelScrool
    End If
    Exit Function

1:  UnHookControl
End Function

Public Sub UnHookControl()
    If HookPtr <> 0 Then
       UnhookWindowsHookEx HookPtr
       HookPtr = 0
'       EventControl.BackColor = vbGreen 'Test
       Set EventControl = Nothing
    End If
End Sub

'---------------------------
'Status query (not required)
'---------------------------

Public Property Get IsHookControl() As Boolean
    IsHookControl = (HookPtr <> 0)
End Property

'------------------
'Pointer Functionen
'------------------

Public Function CurserPtr() As LongPtr
    Dim tPT As POINTAPI: GetCursorPos tPT
'    CurserPtr = WindowFromPoint(tPT.XY)
    CurserPtr = WindowFromPoint(tPT.x, tPT.y)
End Function

Private Function FormPtr() As LongPtr
    Dim fHw As LongPtr: fHw = FindWindow("ThunderDFrame", EventControl.Parent.Caption)
'    DragAcceptFiles fHw, True
    FormPtr = GetWindowLongPtr(fHw, GWL_HINSTANCE)
End Function
