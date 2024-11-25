VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim fr As IControl

Event ItemSelected(ByVal iItem As Long, ByVal iSubItem As Long)
Event ItemClick(ByVal iItem As Long, ByVal iSubItem As Long)
'Event ItemDClick(ByVal iItem&, ByVal iSubItem&)

Public Sub Init(ByVal Ctl As IOptionFrame)
    Dim p&, I&, j&, mWS&
'    Dim ii&(1)
    Dim ii As InitCommonControlsExType
    Dim wd As Long, ht As Long
    
    Dim flbkColor As Long
    Dim fnt As LongPtr
    Dim buf As Long
    
    If TT.hParent Then Exit Sub
    Set fr = Ctl
       
    TT.hParent = fr.[_GethWnd]
    
    If TT.hParent = 0 Then
        Err.Raise 91
    End If
    
    Set acc = fr
    acc.accLocation 0, 0, wd, ht
    
'    ii(0) = 8
'    ii(1) = 1   'ICC_LISTVIEW_CLASSES
'    InitCommonControlsEx VarPtr(ii(0))
    ii.dwSize = 8
    ii.dwICC = 1   'ICC_LISTVIEW_CLASSES
    InitCommonControlsEx ii
    
    With TT
        .pfn = GetPtr(AddressOf Redirect)
        SetWindowSubclass .hParent, .pfn, .hParent, ObjPtr(Me)
        
        '�R�����R���g���[���̃E�B���h�E�X�^�C��
        mWS = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
        mWS = mWS Or &O1 Or LVS_SHOWSELALWAYS
        'LVS_REPORT��&01�@LVS_LIST��&03
       
        .hChild = CreateWindowExW(WS_EX_CLIENTEDGE, _
                            StrPtr("SysListView32"), 0, mWS, _
                            0, 0, wd, ht, .hParent, 0, 0, 0)
        SetWindowSubclass .hChild, .pfn, .hChild, ObjPtr(Me)
                                    
        '�g���X�^�C��
        mWS = LVS_EX_GRIDLINES Or LVS_EX_CHECKBOXES Or LVS_EX_FULLROWSELECT
'        LVS_EX_FULLROWSELECT��1�s�I��
'        LVS_EX_GRIDLINES���O���b�h���C��
'        LVS_EX_CHECKBOXES ��ChekBox�t��

        SendMessageW .hChild, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, mWS

       '------<color>----------
'       flbkColor = GetSysColor(vbButtonFace And &HFF)
'       SendMessageW .hChild, LVM_SETTEXTBKCOLOR, 0, flbkColor
'       SendMessageW .hChild, LVM_SETBKCOLOR, 0, flbkColor
'       SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(255, 0, 0)
  
       '------<font>----------
       fnt = CreateFont(16, _
                         FW_NORMAL, _
                         0, _
                         0, _
                         0, _
                         False, _
                         False, _
                         False, _
                         DEFAULT_CHARSET, _
                         OUT_DEFAULT_PRECIS, _
                         CLIP_DEFAULT_PRECIS, _
                         DEFAULT_QUALITY, _
                         DEFAULT_PITCH, _
                         "MS UI Gothic")

'                        int nHeight,               '// �t�H���g�̍���
'                        int nWidth,                '// ���ϕ�����
'                        int nEscapement,           '// ������������̊p�x
'                        int nOrientation,          '// �x�[�X���C���̊p�x
'                        int fnWeight,              '// �t�H���g�̑���
'                        DWORD fdwItalic,           '// �Α̂ɂ��邩�ǂ���
'                        DWORD fdwUnderline,        '// ������t���邩�ǂ���
'                        DWORD fdwStrikeOut,        '// ����������t���邩�ǂ���
'                        DWORD fdwCharSet,          '// �����Z�b�g�̎��ʎq
'                        DWORD fdwOutputPrecision,  '// �o�͐��x
'                        DWORD fdwClipPrecision,    '// �N���b�s���O���x
'                        DWORD fdwQuality,          '// �o�͕i��
'                        DWORD fdwPitchAndFamily,   '// �s�b�`�ƃt�@�~��
'                        LPCTSTR lpszFace           '// �t�H���g��
        SendMessageW .hChild, WM_SETFONT, fnt, ByVal 1&
       
    End With
End Sub

Public Function InsertColumn(Title As String, ByVal iCol As LongPtr, ByVal Width As Long) As LongPtr
    Dim lvc As LVCOLUMN
    With lvc
        .mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_SUBITEM
        .cx = Width
        .pszText = StrPtr(Title)
    End With
    With TT
        If .hChild = 0 Then Exit Function
        InsertColumn = SendMessageW(.hChild, LVM_INSERTCOLUMN, iCol, VarPtr(lvc))
    End With
    
End Function

Public Function InsertItem(sItem As String, ByVal iItem As Long) As LongPtr
    Dim lvi As LVITEM
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .pszText = StrPtr(sItem)
    End With
    With TT
        If .hChild = 0 Then Exit Function
        InsertItem = SendMessageW(.hChild, LVM_INSERTITEM, 0, VarPtr(lvi))
    End With
    
End Function

Public Function SetItem(sItem As String, ByVal iItem As Long, ByVal iSubItem As Long) As LongPtr
    Dim lvi As LVITEM
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .iSubItem = iSubItem
        .pszText = StrPtr(sItem)
    End With
    With TT
        If .hChild = 0 Then Exit Function
        SetItem = SendMessageW(.hChild, LVM_SETITEM, 0, VarPtr(lvi))
    End With
End Function

Public Function DeleteAllItems() As LongPtr
    DeleteAllItems = SendMessageW(TT.hChild, &H1009, 0, 0)
End Function

Public Function GetItemCount() As LongPtr
'    Dim Cnt As Integer
'    Cnt = SendMessageW(TT.hChild, &H1004, 0, 0)
    GetItemCount = SendMessageW(TT.hChild, &H1004, 0, 0)
End Function

Public Function LabelText(ByVal iItem As Long, ByVal iSubItem As Long) As String
    Dim lv As LVITEM
    Dim buf(255) As Integer
    
    With lv
        .cchTextMax = 256
        .iSubItem = iSubItem
        .pszText = VarPtr(buf(0))
    End With
    SendMessageW TT.hChild, LVM_GETITEMTEXT, iItem, VarPtr(lv)
    With lv
        If .pszText = 0 Then Exit Function
        CopyMemory ByVal VarPtr(LabelText), SysAllocString(.pszText), 4
    End With
    
End Function

Public Function GetCheckState(ByVal ii As Long) As Boolean
' �`�F�b�N��Ԃ��擾�iii = �C���f�b�N�X�j...
    ' LVM_GETITEMSTATE...
    ii = SendMessageW(TT.hChild, &H102C, ii, &H1000&)
    GetCheckState = ii  '= 0

End Function

Public Function SetCheckState(ByVal ii As Long, ByVal F As Boolean)
' �`�F�b�N��Ԃ��Z�b�g�iii = �C���f�b�N�X)
' ii���u-1�v�̎��͑S�Ă̍��ڂ��Ώ�
    Dim buf As LVITEM
    With buf
        .stateMask = &HF000&    'LVIS_STATEIMAGEMASK
        .State = &H1000&
        If F Then .State = &H2000&
    End With
    ' LVM_SETITEMSTATE
    SendMessageW TT.hChild, &H102B, ii, VarPtr(buf)
End Function

Friend Function WndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim iCode As Long
    Dim lvi As LVITEM
    Dim nlv As NMLISTVIEW
    Dim nia As NMITEMACTIVATE  ' ���ǉ�

    
    If hwnd = TT.hParent Then
    
        Select Case uMsg
        
        Case WM_SETFOCUS
            SetFocus TT.hChild
            Exit Function

        Case WM_NOTIFY
            CopyMemory iCode, ByVal lParam + 8, 4
            
            Select Case iCode
                    
            Case LVN_ITEMCHANGED 'ItemSelected
                CopyMemory nlv, ByVal lParam, LenB(nlv)
                RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
                Exit Function
                
            Case NM_CLICK 'ItemClick
                CopyMemory nia, ByVal lParam, LenB(nia)
                RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
                Exit Function
'
'            Case NM_DBLCLK
'                CopyMemory nia, ByVal lParam, LenB(nia)
'                RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'                Exit Function
            
  
            'Case ���낢��...
            '      :
            '      :
            End Select
        
        End Select
    
    Else
        Select Case uMsg

        Case WM_MOUSEACTIVATE
            If GetFocus() <> hwnd Then
                acc.accSelect 1&
                WndProc = MA_NOACTIVATE
                Exit Function
            End If

        Case WM_NCDESTROY
            RemoveSubClass
            Exit Function
       '----------------------------
        Case WM_KEYDOWN
            Select Case wParam
                Case VK_RETURN
                    Exit Function
                Case VK_UP, VK_DOWN
                    WndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
                    Exit Function
            End Select
            Exit Function
       '-----------------------------
        End Select
    End If
    WndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
    
End Function

Private Function GetPtr(ByVal ptr As LongPtr) As LongPtr
    GetPtr = ptr
End Function

Private Sub RemoveSubClass()
    With TT
        If .hChild Then
            RemoveWindowSubclass .hChild, .pfn, .hChild
            .hChild = 0
        End If
        
        If .hParent Then
            RemoveWindowSubclass .hParent, .pfn, .hParent
            .hParent = 0
        End If
    End With
End Sub

Private Sub Class_Terminate()
    RemoveSubClass
End Sub
