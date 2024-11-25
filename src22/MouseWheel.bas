Attribute VB_Name = "MouseWheel"
Option Explicit

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long

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

'Private comboBox As MSForms.comboBox
'Private ListBox As MSForms.ListBox
'Private Frame As MSForms.Frame
Private ControlEvt As ControlEvent

Type POINT
    x As Long
    y As Long
End Type

Type MSLLHOOKSTRUCT
    pt As POINT
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private Function MouseProcOnComboBox(ByVal ncode As Long, ByVal wParam As LongPtr, ByRef lParam As MSLLHOOKSTRUCT) As LongPtr
    On Error GoTo ErrLine

'    If comboBox Is Nothing Then
    If ControlEvt Is Nothing Then
        Exit Function
    End If
    If 0 = hHook Then
        Exit Function
    End If

    With ControlEvt.Control
        Dim currentIdx As Long
    '    currentIdx = comboBox.TopIndex
        currentIdx = .TopIndex
        
        If ncode >= 0 And wParam = WM_MOUSEWHEEL Then
            If lParam.mouseData < 0 Then
                If .ListCount > currentIdx Then .TopIndex = currentIdx + 1
            Else
                If currentIdx > 0 Then .TopIndex = currentIdx - 1
            End If
        End If
    End With

    MouseProcOnComboBox = CallNextHookEx(hHook, ncode, wParam, lParam)

    Exit Function
ErrLine:
    Debug.Print "MouseProcOnComboBox is called"
    Debug.Print "error is: "
    Debug.Print Err.Description
End Function

Private Function MouseProcOnListBox(ByVal ncode As Long, ByVal wParam As LongPtr, ByRef lParam As MSLLHOOKSTRUCT) As LongPtr
    On Error GoTo ErrLine

'    If ListBox Is Nothing Then
    If ControlEvt Is Nothing Then
        Exit Function
    End If
    If 0 = hHook Then
        Exit Function
    End If


    If ncode >= 0 And wParam = WM_MOUSEWHEEL Then
        With ControlEvt
            Dim currentIdx As Long
            currentIdx = .Control.TopIndex
            Dim pos As POINTAPI
            pos = PointToPixcel(GetControlPosition(.Control, TopLeft))
            Dim posPt As POINTF
            
            If lParam.mouseData < 0 Then
                If .Control.ListCount > currentIdx Then .Control.TopIndex = currentIdx + 1
            Else
                If currentIdx > 0 Then .Control.TopIndex = currentIdx - 1
            End If
            pos.x = lParam.pt.x - pos.x
            pos.y = lParam.pt.y - pos.y
            posPt = PixcelToPoint(pos)
            
            Debug.Print "MouseWheel:TopIndex:" & .Control.TopIndex & " / MousePointer:(" & posPt.x & "," & posPt.y & ")"
    '        Debug.Print "MouseWheel:TopIndex:" & listBox.TopIndex & " / MousePointer:(" & lParam.pt.X & "px," & lParam.pt.Y & "px)"
        End With
    End If

    MouseProcOnListBox = CallNextHookEx(hHook, ncode, wParam, lParam)

    Exit Function
ErrLine:
    Debug.Print "MouseProcOnListBox is called"
    Debug.Print "error is: "
    Debug.Print Err.Description
End Function

Private Function MouseProcOnFrame(ByVal ncode As Long, ByVal wParam As LongPtr, ByRef lParam As MSLLHOOKSTRUCT) As LongPtr
    On Error GoTo ErrLine

'    If Frame Is Nothing Then
    If ControlEvt Is Nothing Then
        Exit Function
    End If
    If 0 = hHook Then
        Exit Function
    End If

    With ControlEvt.Control
        Dim currentIdx As Long
        currentIdx = .ScrollTop
    
        If ncode >= 0 And wParam = WM_MOUSEWHEEL Then
            If lParam.mouseData < 0 Then
                If .ScrollHeight > currentIdx Then .ScrollTop = currentIdx + 10
            Else
                If currentIdx > 0 Then .ScrollTop = currentIdx - 10
            End If
        End If
    End With

    MouseProcOnFrame = CallNextHookEx(hHook, ncode, wParam, lParam)

    Exit Function
ErrLine:
    Debug.Print "MouseProcOnFrame is called"
    Debug.Print "error is: "
    Debug.Print Err.Description
End Function

'Public Function HookComboBox(ByRef Box As MSForms.comboBox)
'    If 0 <> hHook Then
'        UnHook
'    End If
'
'    If 0 = hHook Then
'        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnComboBox, 0, 0)
'
'        If 0 <> hHook Then
'            Set comboBox = Box
'        End If
'    End If
'End Function
'
'Public Function HookListBox(ByRef Box As MSForms.ListBox)
'    If 0 <> hHook Then
'        UnHook
'    End If
'
'    If 0 = hHook Then
'        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnListBox, 0, 0)
'
'        If 0 <> hHook Then
'            Set ListBox = Box
'        End If
'    End If
'End Function
'
'Public Function HookFrame(ByRef Box As MSForms.Frame)
'    If 0 <> hHook Then
'        UnHook
'    End If
'
'    If 0 = hHook Then
'        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnFrame, 0, 0)
'
'        If 0 <> hHook Then
'            Set Frame = Box
'        End If
'    End If
'End Function

Public Function HookControl(ByRef CtrlEvt As ControlEvent)
    If 0 <> hHook Then
        UnHook
    Else
        Select Case TypeName(CtrlEvt.Control)
            Case "Frame"
                hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnFrame, 0, 0)
            Case "ListBox"
                hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnListBox, 0, 0)
            Case "ComboBox"
                hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnComboBox, 0, 0)
            Case Else
                Exit Function
        End Select
        Set ControlEvt = CtrlEvt
    End If
End Function

Public Function UnHook()
    If 0 <> hHook Then
        UnhookWindowsHookEx hHook

        hHook = 0

'        Set comboBox = Nothing
'        Set ListBox = Nothing
'        Set Frame = Nothing
        Set ControlEvt = Nothing
    End If
End Function

