Attribute VB_Name = "ListViewEditModule"
Option Explicit

'Private Const WM_SETFOCUS = 7&
'Private Const WM_NOTIFY = &H4E&
'Private Const WM_MOUSEACTIVATE = &H21&
'Private Const WM_NCDESTROY = &H82&
'Private Const WM_SETFONT = &H30
'Private Const WM_VSCROLL = &H115

'Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
'Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
'Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Public Sub ListViewEditTest()
    ListViewEditView.Show
End Sub

Public Function Redirect(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
    Redirect = lv.WndProc(hWnd, uMsg, wParam, lParam)
End Function

Public Function RedirectFrm(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
    RedirectFrm = lv.WndProcFrm(hWnd, uMsg, wParam, lParam)
End Function

Public Function RedirectLV(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
    RedirectLV = lv.WndProcLV(hWnd, uMsg, wParam, lParam)
End Function

Public Sub ShowMessage(prm As LongPtr)
    ListViewEditView.Label2.Caption = prm
End Sub

Public Sub ShowMessage2(prm As LongPtr)
    ListViewEditView.Label3.Caption = prm
End Sub

Public Sub ShowMessage3(prm As Long)
    ListViewEditView.Label4.Caption = prm
End Sub

'Public Function WndProcLV(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
''    If hWnd = This.hwndLV Then
'        Select Case uMsg
'
''            Case WM_MOUSEACTIVATE
'''                If GetFocus() <> hWnd Then
'''                    acc.accSelect 1&
'''                    WndProc = MA_NOACTIVATE
'''                    Exit Function
'''                End If
'''
''            Case WM_NCDESTROY
''                RemoveSubClass
''                Exit Function
'''           '----------------------------
''            Case WM_KEYDOWN
'''                Select Case wParam
'''                    Case VK_RETURN
'''                        Exit Function
'''                    Case VK_UP, VK_DOWN
'''                        WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'''                        Exit Function
'''                    Case Else
'''                End Select
'''                Exit Function
''       '-----------------------------
'            Case WM_VSCROLL
'                Debug.Print "WM_VSCROLL"
'            Case Else
'        End Select
''    End If
'    WndProcLV = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'
'End Function
'
'Public Function WndProcFrm(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
''    If hWnd = This.hWndFrm Then
''            Call ShowMessage(hWnd)
''        Call ShowMessage2(This.hWndFrm)
'        Call ShowMessage3(uMsg)
'        Debug.Print "wndproc"
'        Select Case uMsg
'
'            Case WM_SETFOCUS
''                SetFocus This.hwndLV
'                Call ShowMessage(lParam)
'                Exit Function
'
'            Case WM_VSCROLL
'                Debug.Print "WM_VSCROLL"
'
'            Case WM_NCDESTROY
'                Debug.Print "WM_NCDESTROY"
'            Case WM_NOTIFY
''    ''            MoveMemory iCode, ByVal lParam + 8, 4
''    '
''    ''            Select Case iCode
'                Debug.Print "WM_NOTIFY"
'                Call ShowMessage(lParam)
'                Select Case lParam
''
'''                    Case LVN_ITEMCHANGED 'ItemSelected
'''                        MoveMemory nlv, ByVal lParam, LenB(nlv)
'''                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
'''                        Exit Function
'''
'''                    Case NM_CLICK 'ItemClick
'''                        MoveMemory nia, ByVal lParam, LenB(nia)
'''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'''                        Exit Function
'''
'''                    Case NM_DBLCLK
'''                        MoveMemory nia, ByVal lParam, LenB(nia)
'''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'''                        Exit Function
''
''                    'Case ‚¢‚ë‚¢‚ë...
''                    '      :
''                    '      :
'                    Case Else
'
'                End Select
''
'        End Select
''
''    End If
'    WndProcFrm = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'
'End Function

'Private Function GetPtr(ByVal ptr As LongPtr) As LongPtr
'    GetPtr = ptr
'End Function

'Private Sub RemoveSubClass()
'    With This
'        If .hwndLV Then
'            RemoveWindowSubclass .hwndLV, .pfnLV, .hwndLV
'            .hwndLV = 0
'        End If
'
'        If .hWndFrm Then
'            RemoveWindowSubclass .hWndFrm, .pfnFrm, .hWndFrm
'            .hWndFrm = 0
'        End If
'    End With
'End Sub
