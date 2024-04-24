Attribute VB_Name = "MouseWheel"
Option Explicit

#If Win64 Then
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long

#Else
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14

Private Const WM_MOUSEMOVE As Long = &H200

Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203

Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_MOUSEWHEEL As Long = &H20A

Private Const WM_KEYDOWN As Long = &H100
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28

Private hHook As LongPtr

Private comBox As MSForms.ComboBox
Private listBox As MSForms.listBox

Type POINT
    X As Long
    Y As Long
End Type

Type MSLLHOOKSTRUCT
    pt As POINT
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private Function MouseProcOnComBox(ByVal ncode As Long, ByVal wParam As LongPtr, ByRef lParam As MSLLHOOKSTRUCT) As LongPtr
    On Error GoTo ErrLine

    If comBox Is Nothing Then
        Exit Function
    End If
    If 0 = hHook Then
        Exit Function
    End If

    Dim currentIdx As Long
'    currentIdx = comBox.ListIndex
    currentIdx = comBox.TopIndex
    
    If ncode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
'            If comBox.ListCount > currentIdx Then comBox.ListIndex = currentIdx + 1
            If comBox.ListCount > currentIdx Then comBox.TopIndex = currentIdx + 1
        Else
'            If currentIdx > 0 Then comBox.ListIndex = currentIdx - 1
            If currentIdx > 0 Then comBox.TopIndex = currentIdx - 1
        End If
    End If

    MouseProcOnComBox = CallNextHookEx(hHook, ncode, wParam, lParam)

    Exit Function
ErrLine:
    Debug.Print "MouseProcOnComBox is called"
    Debug.Print "error is: "
    Debug.Print Err.Description
End Function

Private Function MouseProcOnListBox(ByVal ncode As Long, ByVal wParam As LongPtr, ByRef lParam As MSLLHOOKSTRUCT) As LongPtr
    On Error GoTo ErrLine

    If listBox Is Nothing Then
        Exit Function
    End If
    If 0 = hHook Then
        Exit Function
    End If

    Dim currentIdx As Long
'    currentIdx = listBox.ListIndex
    currentIdx = listBox.TopIndex

    If ncode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
'            If listBox.ListCount > currentIdx Then listBox.ListIndex = currentIdx + 1
            If listBox.ListCount > currentIdx Then listBox.TopIndex = currentIdx + 1
        Else
'            If currentIdx > 0 Then listBox.ListIndex = currentIdx - 1
            If currentIdx > 0 Then listBox.TopIndex = currentIdx - 1
        End If
    End If

    MouseProcOnListBox = CallNextHookEx(hHook, ncode, wParam, lParam)

    Exit Function
ErrLine:
    Debug.Print "MouseProcOnListBox is called"
    Debug.Print "error is: "
    Debug.Print Err.Description
End Function

Public Function ChooseHook_ComBox(ByRef Box As MSForms.ComboBox)
    If 0 <> hHook Then
        uHook
    End If

    If 0 = hHook Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnComBox, 0, 0)

        If 0 <> hHook Then
            Set comBox = Box
        End If
    End If
End Function

Public Function ChooseHook_ListBox(ByRef Box As MSForms.listBox)
    If 0 <> hHook Then
        uHook
    End If

    If 0 = hHook Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnListBox, 0, 0)

        If 0 <> hHook Then
            Set listBox = Box
        End If
    End If
End Function

Public Function uHook()
    If 0 <> hHook Then
        UnhookWindowsHookEx hHook

        hHook = 0

        Set comBox = Nothing

        Set listBox = Nothing
    End If
End Function
