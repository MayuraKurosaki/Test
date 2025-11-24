VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ICC_LISTVIEW_CLASSES = &H1          ' listview, header

Dim fr As IControl
Dim fnt As LongPtr

Event ItemSelected(ByVal iItem As Long, ByVal iSubItem As Long)
Event ItemClick(ByVal iItem As Long, ByVal iSubItem As Long)
'Event ItemDClick(ByVal iItem As Long, ByVal iSubItem As Long)

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    DeleteObject fnt
    RemoveSubClass
End Sub

Public Function Init(ByVal Ctl As IOptionFrame) As Boolean
    Init = False
    Dim mWS As Long
    
'    Dim ii&(1)
'    Dim ii As InitCommonControlsExStruct
'    Dim iccex As InitCommonControlsExStruct
    Dim iccex As InitCommonControlsExType
    Dim wd As Long, ht As Long
    
    Dim flbkColor As Long
'    Dim fnt As LongPtr
    Dim buf As Long
    
    If TT.hParent Then Exit Function
    Set fr = Ctl
       
    TT.hParent = fr.[_GethWnd]
    
'    Call WindowFromAccessibleObject(fr, TT.hParent)
    
    If TT.hParent = 0 Then
        Err.Raise 91
    End If
    
    Set acc = fr
    acc.accLocation 0, 0, wd, ht
    
'    ii(0) = 8
'    ii(1) = 1   'ICC_LISTVIEW_CLASSES
'    InitCommonControlsEx VarPtr(ii(0))
    iccex.lngSize = 8
    iccex.lngICC = ICC_LISTVIEW_CLASSES
'    InitCommonControlsEx VarPtr(ii)
    InitCommonControlsEx iccex
    
    With TT
        .pfn = GetPtr(AddressOf Redirect)
        SetWindowSubclass .hParent, .pfn, .hParent, ObjPtr(Me)
        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hParent), "Unicode", "ANCI")
        
        
        'コモンコントロールのウィンドウスタイル
        mWS = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
'        mWS = mWS Or &O3 Or LVS_SHOWSELALWAYS
        mWS = mWS Or LVS_REPORT Or LVS_SHOWSELALWAYS
        'LVS_REPORT→&01　LVS_LIST→&03
       
        .hChild = CreateWindowExW(WS_EX_CLIENTEDGE, _
                            StrPtr("SysListView32"), 0, mWS, _
                            0, 0, wd, ht, .hParent, 0, 0, 0)
'        .hChild = CreateWindowEx(WS_EX_CLIENTEDGE, _
                            "SysListView32", 0, mWS, _
                            0, 0, wd, ht, .hParent, 0, 0, 0)
        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hChild), "Unicode", "ANCI")
        Debug.Print "ｈChild" & .hChild
        SetWindowSubclass .hChild, .pfn, .hChild, ObjPtr(Me)
                                    
        '拡張スタイル
'       mWS = LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_REPORT
        mWS = LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_CHECKBOXES
'        mWS = LVS_EX_CHECKBOXES '<ChekBox付き>
        SendMessageW .hChild, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal mWS
'        SendMessage .hChild, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal mWS

'       '------<color>----------
'       flbkColor = GetSysColor(vbButtonFace And &HFF)
'       SendMessageW .hChild, LVM_SETTEXTBKCOLOR, 0, flbkColor
'       SendMessageW .hChild, LVM_SETBKCOLOR, 0, flbkColor
''       SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(255, 0, 0)
  
       '------<font>----------
       fnt = CreateFontW(16, _
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
                         StrPtr("Yu Gothic UI"))
'                         StrPtr("MS UI Gothic"))
'
'                        int nHeight,               '// フォントの高さ
'                        int nWidth,                '// 平均文字幅
'                        int nEscapement,           '// 文字送り方向の角度
'                        int nOrientation,          '// ベースラインの角度
'                        int fnWeight,              '// フォントの太さ
'                        DWORD fdwItalic,           '// 斜体にするかどうか
'                        DWORD fdwUnderline,        '// 下線を付けるかどうか
'                        DWORD fdwStrikeOut,        '// 取り消し線を付けるかどうか
'                        DWORD fdwCharSet,          '// 文字セットの識別子
'                        DWORD fdwOutputPrecision,  '// 出力精度
'                        DWORD fdwClipPrecision,    '// クリッピング精度
'                        DWORD fdwQuality,          '// 出力品質
'                        DWORD fdwPitchAndFamily,   '// ピッチとファミリ
'                        LPCTSTR lpszFace           '// フォント名
        
'        SendMessageW .hChild, WM_SETFONT, fnt, ByVal 1&
        
'        SendMessage .hChild, WM_SETFONT, fnt, ByVal 1&

'        SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
''    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]
       
    End With
    Init = True
End Function

Public Function SetWindowFont() As Long
    
End Function

Public Function InsertColumn(Title As String, ByVal iCol As Long, ByVal Width As Long) As Long
    Dim lvc As LVCOLUMNW
    With lvc
        .mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_SUBITEM
        .CX = Width
        .pszText = StrPtr(Title)
'        .pszText = Title
    End With
    With TT
        If .hChild = 0 Then Exit Function
        InsertColumn = CLng(SendMessageW(.hChild, LVM_INSERTCOLUMNW, ByVal iCol, lvc))
'        InsertColumn = CLng(SendMessageW(.hChild, LVM_INSERTCOLUMN, ByVal iCol, ByVal VarPtr(lvc)))
'        InsertColumn = CLng(SendMessage(.hChild, LVM_INSERTCOLUMNW, ByVal iCol, lvc))
    End With
'    Debug.Print "InsertColumn"
End Function

Public Function InsertItem(sItem As String, ByVal iItem As Long) As Long
    Dim lvi As LVITEMW
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .pszText = StrPtr(sItem)
'        .pszText = sItem
    End With
    With TT
        If .hChild = 0 Then Exit Function
        InsertItem = CLng(SendMessageW(.hChild, LVM_INSERTITEMW, 0, lvi))
'        InsertItem = CLng(SendMessageW(.hChild, LVM_INSERTITEMW, 0, ByVal VarPtr(lvi)))
'        InsertItem = CLng(SendMessage(.hChild, LVM_INSERTITEMW, 0, lvi))
    End With
'    Debug.Print "InsertItem"
End Function

Public Function SetItem(sItem As String, ByVal iItem As Long, ByVal iSubItem As Long) As Long
    Dim lvi As LVITEMW
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .iSubItem = iSubItem
        .pszText = StrPtr(sItem)
'        .pszText = sItem
    End With
    With TT
        If .hChild = 0 Then Exit Function
        SetItem = CLng(SendMessageW(.hChild, LVM_SETITEMW, 0, lvi))
'        SetItem = CLng(SendMessageW(.hChild, LVM_SETITEMW, 0, ByVal VarPtr(lvi)))
'        SetItem = CLng(SendMessage(.hChild, LVM_SETITEMW, 0, lvi))
    End With
'    Debug.Print "SetItem"
End Function

Public Function DeleteAllItems() As Long
    DeleteAllItems = CLng(SendMessageW(TT.hChild, &H1009, 0, 0))
'    DeleteAllItems = CLng(SendMessage(TT.hChild, &H1009, 0, 0))
'    Debug.Print "DeleteAllItems"
End Function

Public Function GetItemCount() As Long
'    Dim Cnt As Integer
'    Cnt = SendMessageW(TT.hChild, &H1004, 0, 0)
    GetItemCount = CLng(SendMessageW(TT.hChild, &H1004, 0, 0))
'    GetItemCount = CLng(SendMessage(TT.hChild, &H1004, 0, 0))
'    Debug.Print "GetItemCount"
End Function

Public Function LabelText(ByVal iItem As Long, ByVal iSubItem As Long) As String
    Dim lv As LVITEM
    Dim buf(255) As Integer
    
    With lv
        .cchTextMax = 256
        .iSubItem = iSubItem
        .pszText = VarPtr(buf(0))
    End With
    SendMessageW TT.hChild, LVM_GETITEMTEXTW, iItem, VarPtr(lv)
'    SendMessage TT.hChild, LVM_GETITEMTEXT, iItem, VarPtr(lv)
    With lv
        If .pszText = 0 Then Exit Function
        MoveMemory ByVal VarPtr(LabelText), SysAllocString(.pszText), 4
    End With
'    Debug.Print "LabelText"
End Function

Public Function GetCheckState(ByVal ii As Long) As Long 'Boolean
' チェック状態を取得（ii = インデックス）...
    ' LVM_GETITEMSTATE...
    ii = CLng(SendMessageW(TT.hChild, &H102C, ii, &H1000&))
'    ii = CLng(SendMessage(TT.hChild, &H102C, ii, &H1000&))
    GetCheckState = ii  '= 0
'    Debug.Print "GetCheckState"
End Function

Public Function SetCheckState(ByVal ii As Long, ByVal F As Boolean)
' チェック状態をセット（ii = インデックス)
' iiが「-1」の時は全ての項目が対象
    Dim buf As LVITEMW
    With buf
        .stateMask = &HF000&    'LVIS_STATEIMAGEMASK
        .state = &H1000&
        If F Then .state = &H2000&
    End With
    ' LVM_SETITEMSTATE
    SendMessageW TT.hChild, &H102B, ii, VarPtr(buf)
'    SendMessage TT.hChild, &H102B, ii, VarPtr(buf)
'    Debug.Print "SetCheckState"
End Function

Public Function SetTextColor(ClrText As Long) As Boolean
  Dim ret As Long
  ret = CLng(SendMessageW(TT.hChild, LVM_SETTEXTCOLOR, 0, ByVal ClrText))
'  ret = CLng(SendMessageW(TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)))
  If ret <> 0 Then
      SetTextColor = True
      Debug.Print "SetTextColor"
  Else
      SetTextColor = False
  End If
End Function

Public Function SetTextBkColor(ClrTextBk As Long) As Boolean
  Dim ret As Long
    ret = CLng(SendMessageW(TT.hChild, LVM_SETTEXTBKCOLOR, 0, ClrTextBk))
  If ret <> 0 Then
      SetTextBkColor = True
      Debug.Print "SetTextBkColor"
  Else
      SetTextBkColor = False
  End If
End Function

Public Function RedrawItems(iFirst As Long, iLast As Long) As Boolean
  Dim ret As Long
    ret = CLng(SendMessageW(TT.hChild, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast))
  If ret <> 0 Then
      RedrawItems = True
      Debug.Print "RedrawItems"
  Else
      RedrawItems = False
  End If
End Function

Public Function Update(i As Long) As Boolean
  Dim ret As Long
    ret = CLng(SendMessageW(TT.hChild, LVM_UPDATE, ByVal i, 0))
  If ret <> 0 Then
      Update = True
      Debug.Print "Update"
  Else
      Update = False
  End If
End Function

Public Function SetColumnWidth(iCol As Long, CX As Long) As Boolean
  Dim ret As Long
    ret = CLng(SendMessageW(TT.hChild, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(CX, 0)))
  If ret <> 0 Then
      SetColumnWidth = True
      Debug.Print "Update"
  Else
      SetColumnWidth = False
  End If
End Function

Friend Function WndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim iCode As Long
'    Dim lvi As LVITEMW
    Dim nlv As NMLISTVIEW
    Dim nia As NMITEMACTIVATE  ' ←追加

    
    If hwnd = TT.hParent Then
    
        Select Case uMsg
        
            Case WM_SETFOCUS
                SetFocus TT.hChild
'                Debug.Print "Frame:WM_SETFOCUS"
                Exit Function
    
            Case WM_NOTIFY
                MoveMemory iCode, ByVal lParam + 8, 4

                Select Case iCode

                    Case LVN_ITEMCHANGED
                        MoveMemory nlv, ByVal lParam, LenB(nlv)
                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
                        Exit Function
    
                    Case NM_CLICK
                        MoveMemory nia, ByVal lParam, LenB(nia)
                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
                        Exit Function
    
'                    Case NM_DBLCLK
'                        MoveMemory nia, ByVal lParam, LenB(nia)
'                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'                        Exit Function
    
    
'                    Case いろいろ...
':
':
                    Case Else

                End Select
            Case Else
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
            Case Else

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

