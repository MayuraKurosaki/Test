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
Private Const IDD_LISTVIEW = 101
'Private Const WC_LISTVIEW = "SysListView32"

Private Type Fields
    Frame As IControl
    hFont As LongPtr
    ColumnCount As Long
    ListItemsCount As Long
    hParent As LongPtr
    hChild As LongPtr
    pfn As LongPtr
    ColumnHeaders As ColumnHeaders
    ListItems As ListItems
End Type
'Private Type TT
'    hParent As LongPtr
'    hChild As LongPtr
'    pfn As LongPtr
'End Type
'''
'Private TT As TT
Private Type CSV_Strunt
    FilePath As String
    CharSet As String
    IncludeHeader As Boolean
    Header() As String
    QuateCount As Long
End Type


'Private VLItems() As VListItem
Private VLItems As Dictionary


Private CSV As CSV_Strunt
Private acc As IAccessible

Private This As Fields


'--------------Compatible Events----------------
'Occurs after a user edits the label of the currently selected Node or ListItem object.
Public Event AfterLabelEdit(Cancel As Integer, NewString As String)
'Occurs when a user attempts to edit the label of the currently selected ListItem or Node object.
Public Event BeforeLabelEdit(Cancel As Integer)
'Occurs when the user presses and then releases a mouse button over an object.
Public Event Click()
'Occurs when a ColumnHeader object in a ListView control is clicked.
Public Event ColumnClick(ByVal ColumnHeader As ColumnHeader)
'Occurs when you press and release a mouse button and then press and release it again over an object.
Public Event DblClick()



'Occurs when a ListSubItem object is checked
Public Event ItemCheck(ByVal Item As ListItem)
'Occurs when a ListItem object is clicked or selected
'Public Event ItemClick(ByVal Item As ListItem)
'Occurs when the user presses a key while an object has the focus.
Public Event KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'Occurs when the user presses and releases an ANSI key.
Public Event KeyPress(KeyAscii As Integer)
'Occurs when the user releases a key while an object has the focus.
Public Event KeyUp(KeyCode As Integer, ByVal Shift As Integer)
'Occurs when the user presses the mouse button while an object has the focus.
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'Occurs when the user moves the mouse.
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'Occurs when the user releases the mouse button while an object has the focus.
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'OLECompleteDrag event
Public Event OLECompleteDrag(Effect As Long)
'OLEDragDrop event
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'OLEDragOver event
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'OLEGiveFeedback event
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'OLESetData event
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
'OLEStartDrag event
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)





Event ItemSelected(ByVal iItem As Long, ByVal iSubItem As Long)
Event ItemClick(ByVal iItem As Long, ByVal iSubItem As Long)
'Event ItemDClick(ByVal iItem As Long, ByVal iSubItem As Long)

Public Property Get Header() As String()
    Header = CSV.Header
End Property

Public Property Get VList() As Dictionary
    Set VList = VLItems
End Property

Private Sub Class_Initialize()
    With This
        .ColumnCount = 0
        .ListItemsCount = 0
        Set .ColumnHeaders = New ColumnHeaders
        Set .ListItems = New ListItems
    End With
'    Set VLItems = New Dictionary
'    Debug.Print VLItems.Count
End Sub

Private Sub Class_Terminate()
    DeleteObject This.hFont
    RemoveSubClass
    With This
        Set .ColumnHeaders = Nothing
        Set .ListItems = Nothing
    End With
'    Set VLItems = Nothing
End Sub

Public Function Init(ByVal Ctl As IOptionFrame) As Boolean
    Init = False
    Dim mWS As Long
    
    Dim iccex As InitCommonControlsExType
    Dim wd As Long, ht As Long
    
    Dim flbkColor As Long
    Dim buf As Long
    
    If This.hParent Then Exit Function
    Set This.Frame = Ctl
       
    This.hParent = This.Frame.[_GethWnd]
    
'    Call WindowFromAccessibleObject(fr, This.hParent)
    
    If This.hParent = 0 Then
        Err.Raise 91
    End If
    
    Set acc = This.Frame
    acc.accLocation 0, 0, wd, ht
    
    iccex.dwSize = Len(iccex)
    Debug.Print iccex.dwSize
    iccex.dwICC = ICC_LISTVIEW_CLASSES
    InitCommonControlsEx iccex
    
    With This
'        .pfn = GetPtr(AddressOf ListViewModule.Redirect)
        .pfn = GetPtr(AddressOf Redirect)
        Debug.Print "hParent:" & .hParent & " / pfn:" & .pfn
        If SetWindowSubclass(.hParent, .pfn, .hParent, ObjPtr(Me)) = 0 Then Debug.Print "FrameSubclassError": Exit Function
               
        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hParent), "Unicode", "ANCI")
        
        'コモンコントロールのウィンドウスタイル
        mWS = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
        mWS = mWS Or LVS_REPORT Or LVS_SHOWSELALWAYS Or LVS_OWNERDATA
        
        On Error GoTo ErrorHandler
        
        .hChild = CreateWindowEx(WS_EX_CLIENTEDGE, _
                            WC_LISTVIEW, 0, mWS, _
                            0, 0, wd, ht, .hParent, IDD_LISTVIEW, 0, 0)
        If .hChild = 0 Then Exit Function
        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hChild), "Unicode", "ANCI")
        Debug.Print "ｈChild" & .hChild
        If SetWindowSubclass(.hChild, .pfn, .hChild, ObjPtr(Me)) = 0 Then Debug.Print "ListViewSubclassError": Exit Function
                                    
        '拡張スタイル
        mWS = LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_CHECKBOXES Or LVS_EX_DOUBLEBUFFER
        SendMessage .hChild, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal mWS

'       '------<color>----------
'       flbkColor = GetSysColor(vbButtonFace And &HFF)
'       SendMessageW .hChild, LVM_SETTEXTBKCOLOR, 0, flbkColor
'       SendMessageW .hChild, LVM_SETBKCOLOR, 0, flbkColor
''       SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(255, 0, 0)
  
       '------<font>----------
'       .hFont = CreateFontW(16, _
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
       .hFont = CreateFont(16, _
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
                         "Yu Gothic UI")
        
'        SendMessageW .hChild, WM_SETFONT, fnt, ByVal 1&
        
'        SendMessage .hChild, WM_SETFONT, This.hFont, ByVal 1&
'        SendMessage hChild, WM_SETFONT, This.hFont, ByVal 1&

'        SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
''    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]
       
    End With
    Init = True
    Exit Function
    
ErrorHandler:
    Debug.Print GetAPIError(Err.LastDllError)
    Call RemoveSubClass
End Function

Public Function SetWindowFont() As Long
    
End Function

'Public Function SetReportHeader(Items() As String) As Long
'    Dim i As Long
'    This.ColumnCount = UBound(Items)
'    For i = LBound(Items) To This.ColumnCount
'        Call InsertColumn(Items(i), i, 0)
'    Next i
'End Function
Public Function SetHeader() As Long
'    Dim Item As String
    Dim i As Long
'    i = 1
    For i = 1 To This.ColumnHeaders.Count
        Call InsertColumn(This.ColumnHeaders(i), i, 0)
        Debug.Print "SerHeader:" & i & ":" & This.ColumnHeaders(i)
    Next i
'    Dim i As Long
'    This.ColumnCount = UBound(Items)
'    For i = 1 To This.ColumnHeaders.Count
'        Call InsertColumn(Items(i), i, 0)
'    Next i
End Function

'Public Function SetReportRecord(Items() As String) As Long
'    Dim i As Long
'    Call InsertItem(Items(0), This.ListItemsCount)
'    If This.ColumnCount < UBound(Items) Then
'        For i = This.ColumnCount + 1 To UBound(Items)
'            Call InsertColumn("", i, 0)
'        Next i
'        This.ColumnCount = UBound(Items)
'    End If
'    For i = 1 To This.ColumnCount
'        Call SetItem(Items(i), This.ListItemsCount, i)
'    Next i
'    This.ListItemsCount = This.ListItemsCount + 1
'End Function

Public Function SetReportRecord(Records As Dictionary) As Long
    Dim i As Long
    Call InsertItem(Records.Item(i), This.ListItemsCount)
    If This.ColumnCount < Records.Count Then
        For i = This.ColumnCount + 1 To Records.Count
            Call InsertColumn("", i, 0)
        Next i
        This.ColumnCount = Records.Count
    End If
    For i = 1 To This.ColumnCount
        Call SetItem(Records.Item(i), This.ListItemsCount, i)
    Next i
    This.ListItemsCount = This.ListItemsCount + 1
End Function

Public Function InsertColumn(Title As String, ByVal iCol As Long, ByVal Width As Long) As Long
    Dim lvc As LVCOLUMNA
    With lvc
        .mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_SUBITEM
        .CX = Width
        .pszText = Title
    End With
    With This
        If .hChild = 0 Then Exit Function
        InsertColumn = CLng(SendMessage(.hChild, LVM_INSERTCOLUMN, ByVal iCol, lvc))
    End With
End Function

Public Function InsertItem(sItem As String, ByVal iItem As Long) As Long
    Dim lvi As LVITEMA
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .pszText = sItem
    End With
    With This
        If .hChild = 0 Then Exit Function
        InsertItem = CLng(SendMessage(.hChild, LVM_INSERTITEM, 0, lvi))
    End With
End Function

Public Function SetItem(sItem As String, ByVal iItem As Long, ByVal iSubItem As Long) As Long
    Dim lvi As LVITEMA
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .iSubItem = iSubItem
        .pszText = sItem
    End With
    With This
        If .hChild = 0 Then Exit Function
        SetItem = CLng(SendMessage(.hChild, LVM_SETITEM, 0, lvi))
    End With
End Function

Public Function DeleteAllItems() As Long
    DeleteAllItems = CLng(SendMessage(This.hChild, LVM_DELETEALLITEMS, 0, 0))
End Function

Public Function GetItemCount() As Long
    GetItemCount = CLng(SendMessage(This.hChild, LVM_GETITEMCOUNT, 0, 0))
End Function

Public Sub SetItemCount(cItems As Long)
    Debug.Print cItems
    SendMessage This.hChild, LVM_SETITEMCOUNT, ByVal cItems, 0
End Sub

Public Function LabelText(ByVal iItem As Long, ByVal iSubItem As Long) As String
'    Debug.Print "Item:" & iItem & " / SubItem:" & iSubItem
    Dim lv As LVITEMA
    Dim buf As String * 255
'    buf = String(255, vbNullChar)
'    Debug.Print "Len(buf):" & Len(buf)
    
    With lv
        .cchTextMax = 256
        .iSubItem = iSubItem
        .pszText = buf
    End With
    SendMessage This.hChild, LVM_GETITEMTEXT, iItem, lv
    
    LabelText = lv.pszText
'    Debug.Print "Len(LabelText):" & Len(LabelText)
'    Debug.Print "LabelText:" & LabelText
End Function

Public Function GetCheckState(ByVal iItem As Long) As Boolean
' チェック状態を取得（ii = インデックス）...
    ' LVM_GETITEMSTATE...
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_GETITEMSTATE, iItem, &H1000&))
    If ret <> 0 Then
        GetCheckState = True
    Else
        GetCheckState = False
    End If
'    ii = CLng(SendMessage(This.hChild, LVM_GETITEMSTATE, ii, &H1000&))
'    GetCheckState = ii  '= 0
End Function

Public Sub SetCheckState(ByVal iItem As Long, ByVal F As Boolean)
' チェック状態をセット（ii = インデックス)
' iiが「-1」の時は全ての項目が対象
    Dim buf As LVITEMA
    With buf
        .stateMask = &HF000&    'LVIS_STATEIMAGEMASK
        .State = &H1000&
        If F Then .State = &H2000&
    End With
    ' LVM_SETITEMSTATE
    SendMessage This.hChild, LVM_SETITEMSTATE, iItem, VarPtr(buf)
End Sub

Public Function SetTextColor(ClrText As Long) As Boolean
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_SETTEXTCOLOR, 0, ByVal ClrText))
    If ret <> 0 Then
        SetTextColor = True
    Else
        SetTextColor = False
    End If
End Function

Public Function SetTextBkColor(ClrTextBk As Long) As Boolean
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_SETTEXTBKCOLOR, 0, ClrTextBk))
    If ret <> 0 Then
        SetTextBkColor = True
    Else
        SetTextBkColor = False
    End If
End Function

Public Function RedrawItems(iFirst As Long, iLast As Long) As Boolean
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast))
    If ret <> 0 Then
        RedrawItems = True
    Else
        RedrawItems = False
    End If
End Function

Public Function Update(i As Long) As Boolean
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_UPDATE, ByVal i, 0))
    If ret <> 0 Then
        Update = True
    Else
        Update = False
    End If
End Function

Public Function SetColumnWidth(iCol As Long, CX As Long) As Boolean
    Dim ret As Long
    ret = CLng(SendMessage(This.hChild, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(CX, 0)))
    If ret <> 0 Then
        SetColumnWidth = True
    Else
        SetColumnWidth = False
    End If
End Function

Public Function SetAllColumnWidthAuto() As Boolean
    SetAllColumnWidthAuto = False
    Dim i As Long
    For i = 0 To This.ColumnCount - 1
        SetAllColumnWidthAuto = SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER)
        If SetAllColumnWidthAuto = False Then Exit Function
    Next i
End Function

Friend Function WndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim iCode As Long
    Dim nlv As NMLISTVIEW
    Dim nia As NMITEMACTIVATE  ' ←追加

    
    If hwnd = This.hParent Then
    
        Select Case uMsg
        
            Case WM_SETFOCUS
                SetFocus This.hChild
                Exit Function
    
            Case WM_NOTIFY
'                Debug.Print "WM_NOTIFY:"
                Dim rtn As Long
                If (wParam = IDD_LISTVIEW) Then
                    rtn = DoGVNotify(hwnd, lParam)
                End If
                If rtn Then
                  WndProc = rtn
                  Exit Function
                End If
                
'                MoveMemory iCode, ByVal lParam + 8, 4
'
'                Select Case iCode
'
'                    Case LVN_ITEMCHANGED
'                        Debug.Print "LVN_ITEMCHANGED"
'                        MoveMemory nlv, ByVal lParam, LenB(nlv)
'                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
'                        Exit Function
'
'                    Case NM_CLICK
'                        Debug.Print "NM_CLICK"
'                        MoveMemory nia, ByVal lParam, LenB(nia)
'                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'                        Exit Function
'
''                    Case NM_DBLCLK
''                        MoveMemory nia, ByVal lParam, LenB(nia)
''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
''                        Exit Function
'
'
''                    Case いろいろ...
'':
'':
'                    Case Else
'
'                End Select
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

Public Function DoGVNotify(hwnd As LongPtr, lParam As LongPtr) As Long
    Dim sText As String, sSubText As String
    Dim tNMH As NMHDR
    MoveMemory tNMH, ByVal lParam, LenB(tNMH)
    
    Dim nlv As NMLISTVIEW
    Dim nia As NMITEMACTIVATE
    
    Dim tmp As Long
    tmp = tNMH.Code
    
    Select Case tNMH.Code

        Case LVN_ITEMCHANGED
            MoveMemory nlv, ByVal lParam, LenB(nlv)
            RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
            Exit Function
    
        Case NM_CLICK
            MoveMemory nia, ByVal lParam, LenB(nia)
            RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
            Exit Function
            
        Case LVN_GETDISPINFOW
            Dim LVDIW As NMLVDISPINFOW
            MoveMemory ByVal VarPtr(LVDIW), ByVal lParam, LenB(LVDIW)
            With LVDIW.Item
            Debug.Print "GetDispInfoW"
                If (.mask And LVIF_TEXT) Then
                    Debug.Print "LVIF_TEXT:Item:" & .iItem & " SubItem:" & .iSubItem
                    Select Case .iSubItem
                        Case 0
'                            .sText = ItemTxt(.iItem)
'                                .cchTextMax = Len(sText)
'                            .pszText = StrPtr(VLItems(.iItem).sText)
                            Debug.Print "Item:" & This.ListItems(.iItem)
'                            .pszText = StrPtr(This.ListItems(.iItem).Text)
                            .pszText = This.ListItems(.iItem).TextPtr
                        Case Else
                            'sSubText = "Subitem " & .iSubItem
'                                .cchTextMax = Len(sSubText)
'                            .pszText = StrPtr(VLItems(.iItem).sSubItems(.iSubItem))
                            Debug.Print "SubItem:" & This.ListItems(.iItem).ListSubItems(.iSubItem)
'                            .pszText = StrPtr(This.ListItems(.iItem).ListSubItems(.iSubItem).Text)
                            .pszText = This.ListItems(.iItem).ListSubItems(.iSubItem).TextPtr
'                                Debug.Print "subitemtext=" & StrPtr(sSubText)
                    End Select
                End If
                
                If (.mask And LVIF_IMAGE) Then
                    Debug.Print "LVIF_IMAGE"
                    Select Case .iSubItem
                        Case 0
'                            .iImage = VLItems(.iItem).iImage
'                            .iImage = This.ListItems(.iItem).Icon
                    End Select
                End If
                MoveMemory ByVal lParam, ByVal VarPtr(LVDIW), LenB(LVDIW)
            End With
           
        Case LVN_GETDISPINFOA
            Debug.Print "GetDispInfoA"
            Dim LVDI As NMLVDISPINFOA
            MoveMemory ByVal VarPtr(LVDI), ByVal lParam, LenB(LVDI)
            With LVDI.Item
                
                If (.mask And LVIF_TEXT) Then
                    Debug.Print "LVIF_TEXT"
                    Select Case .iSubItem
                        Case 0
'                                sText = ItemTxt(.iItem)
'                                .cchTextMax = Len(sText)
                            .pszText = StrPtr(VLItems(.iItem).sText)
                        Case Else
                            'sSubText = "Subitem " & .iSubItem
'                                .cchTextMax = Len(sSubText)
                            .pszText = StrPtr(VLItems(.iItem).sSubItems(.iSubItem))
'                                Debug.Print "subitemtext=" & StrPtr(sSubText)
                    End Select
                End If
                
                If (.mask And LVIF_IMAGE) Then
                    Debug.Print "LVIF_IMAGE"
                    Select Case .iSubItem
                        Case 0
                            .iImage = VLItems(.iItem).iImage
                    End Select
                End If
                MoveMemory ByVal lParam, ByVal VarPtr(LVDI), LenB(LVDI)
            End With
        End Select
    Exit Function
End Function

Private Function GetPtr(ByVal ptr As LongPtr) As LongPtr
    GetPtr = ptr
End Function

Private Sub RemoveSubClass()
    With This
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

' Write To CSV
Public Sub ListViewToCSV(Optional ByVal FileName As String, Optional ByVal CharSet As String = "SHIFT-JIS")

End Sub

' Read From CSV
Public Sub CsvToVListView(ByVal FileName As String, Optional ByVal HasHeader As Boolean = True, Optional ByVal CharSet As String = "Auto")
    With CSV
        .FilePath = FileName
        .CharSet = CharSet
        .IncludeHeader = HasHeader
    End With
    
    'readCsvでCSVを読み込み
    Dim CSVText As String
    CSVText = readCsv
 
     
'    ReDim VLItems(0 To 0)
    'CsvToJaggedで行・フィールドに分割してジャグ配列に
    Call CsvToJagged(CSVText)
    
    SetItemCount This.ListItems.Count
    This.ColumnCount = This.ColumnHeaders.Count
End Sub

Private Sub CsvToJagged(ByVal CSVText As String)
'    Dim FieldArray() As String 'フィールド
    Dim FieldText As String '1フィールド文字列
    Dim blnCrLf As Boolean '改行判定
    Dim i As Long '行位置
    Dim j As Long '列位置
    Dim k As Long '文字位置
    Dim ItemText As String
    Dim ProcHeader As Boolean

'    ReDim FieldArray(0 To 0)
    
'    Dim SubItem As ListSubItem
'    Set SubItem = New ListSubItem
    
    ProcHeader = CSV.IncludeHeader
    
    i = 0 'シートの1行目から出力
    j = 0 '列位置はputChildArrayでカウントアップ
    CSV.QuateCount = 0 'ダブルクォーテーションの数
    FieldText = ""
    ItemText = ""
    For k = 1 To Len(CSVText)
        Select Case Mid(CSVText, k, 1)
            Case vbLf, vbCr '「"」が偶数なら改行、奇数ならただの文字
                If CSV.QuateCount Mod 2 = 0 Then
                    blnCrLf = False
                    If k > 1 Then '改行のCrLfはCrで改行判定済なので無視する
                        If Mid(CSVText, k - 1, 2) = vbCrLf Then
                            blnCrLf = True
                        End If
                    End If
                    If blnCrLf = False Then
                        'これが改行となる
'                        Debug.Print "改行:" & i
                        If ProcHeader Then
'                            Debug.Print "ヘッダ最後"
                            This.ColumnHeaders.Add j + 1, , FieldText, 0
                            FieldText = ""
                            
'                            CSV.Header = FieldArray
'                            Call SetReportHeader(FieldArray)
                            Call SetHeader
                            ProcHeader = False
                            i = 0
                        Else
                            Call SetFieldText(i, j, FieldText)
'                            Call SetRecord(ItemText, FieldArray, i, j, FieldText)
                            i = i + 1 '列位置
'                            Dim Item As VListItem
'                            VLItems.Add i, Item
'                            ReDim Preserve VLItems(0 To i)
                        End If
'                        ReDim FieldArray(0 To 0)
                        j = 0 '列位置
                        CSV.QuateCount = 0 'ダブルクォーテーション数
                    End If
                Else
                    FieldText = FieldText & Mid(CSVText, k, 1)
                End If
            Case ",", vbTab '「"」が偶数なら区切り、奇数ならただの文字
                If CSV.QuateCount Mod 2 = 0 Then
                    If ProcHeader Then
'                        Debug.Print "ヘッダ"
                        This.ColumnHeaders.Add j + 1, , FieldText, 0
'                    ElseIf j = 0 Then
'                        Call AddRecord(i, FieldText)
'                        ItemText = FieldText
                    Else
                        Call SetFieldText(i, j, FieldText)
                    End If
                    FieldText = ""
                    CSV.QuateCount = 0
                    j = j + 1
                Else
                    FieldText = FieldText & Mid(CSVText, k, 1)
                End If
            Case """" '「"」のカウントをとる
                CSV.QuateCount = CSV.QuateCount + 1
                FieldText = FieldText & Mid(CSVText, k, 1)
            Case Else
                FieldText = FieldText & Mid(CSVText, k, 1)
        End Select
    Next

    '最終行の最終列の処理
    If j > 0 And FieldText <> "" Then
'        Call SetRecord(ItemText, FieldArray, i, j, FieldText)
        Call SetFieldText(i, j, FieldText)
    End If
End Sub

'レコードを登録
'Private Sub SetRecord(ByRef ItemText As String, ByRef Fields() As String, ByRef i As Long, ByRef j As Long, ByRef FieldText As String)
Private Sub AddRecord(ByRef i As Long, ByRef FieldText As String)
        
'    If i > VLItems.Count Then '常に成立するが一応記述
'        Dim Item As VListItem
'        VLItems.Add i, Item
''        ReDim Preserve VLItems(0 To i)
'    End If
'    Dim Item As VListItem
    With This.ListItems.Add(i, "", FieldText)
'        .ListSubItems.Add j, "", Fields(j)
    End With
'    VLItems.Add i, Item
'    VLItems(i).sText = ItemText
'    VLItems(i).sSubItems = Fields
    
'    i = i + 1 '列位置
'    j = 0 '列位置
    CSV.QuateCount = 0 'ダブルクォーテーション数
    FieldText = "" '1フィールド文字列
'    ReDim Fields(0 To 0)
End Sub
'Private Sub SetRecord(ByRef ItemText As String, ByRef Fields() As String, ByRef i As Long, ByRef j As Long, ByRef FieldText As String)
'
''    If i > VLItems.Count Then '常に成立するが一応記述
''        Dim Item As VListItem
''        VLItems.Add i, Item
'''        ReDim Preserve VLItems(0 To i)
''    End If
'    Dim Item As VListItem
'    With Item
'        .sText = ItemText
'        .sSubItems = Fields
'    End With
'    VLItems.Add i, Item
''    VLItems(i).sText = ItemText
''    VLItems(i).sSubItems = Fields
'
'    i = i + 1 '列位置
'    j = 0 '列位置
'    CSV.QuateCount = 0 'ダブルクォーテーション数
'    FieldText = "" '1フィールド文字列
'    ReDim Fields(0 To 0)
'End Sub

'1フィールドごとにセルに出力
Private Sub SetFieldText(ByRef i As Long, ByRef j As Long, ByRef FieldText As String)
'    Static RecordIndex As Long
'    Debug.Print "RecordIndex:" & RecordIndex & " / i:" & i
    
    '「""」を「"」で置換
    FieldText = Replace(FieldText, """""", """")
    '前後の「"」を削除
    If Left(FieldText, 1) = """" And Right(FieldText, 1) = """" Then
        If Len(FieldText) <= 2 Then
            FieldText = ""
        Else
            FieldText = Mid(FieldText, 2, Len(FieldText) - 2)
        End If
    End If
'    i = i + 1
'    If RecordIndex < i Then
    If j = 0 Then
        This.ListItems.Add Index:=i, Key:="", Text:=FieldText
'        RecordIndex = i
    Else
        This.ListItems(i).ListSubItems.Add Index:=j, Key:="", Text:=FieldText
    End If
    
    j = j + 1
    FieldText = ""
    CSV.QuateCount = 0
End Sub

'文字コードを自動判別し、全行をCrLf区切りに統一してStringに入れる
Private Function readCsv() As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim inTS As Object
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
 
    Dim strRec As String
    Dim i As Long
    Dim aryRec() As String

    If CSV.CharSet = "Auto" Then CSV.CharSet = getCharSet
    Debug.Print "CharSet:" & CSV.CharSet
    Select Case UCase(CSV.CharSet)
        Case "UTF-8", "UTF-8N"
            'ADOを使って読込、その後の処理を統一するため全レコードをCrLfで結合
            'Set inTS = objFSO.OpenTextFile(strFile, ForAppending)
            Set inTS = objFSO.OpenTextFile(CSV.FilePath, 8)
            i = inTS.Line - 1
            inTS.Close
            Debug.Print i & "行"
            ReDim aryRec(i)
            With adoSt
                '.Type = adTypeText
                .type = 2
                .CharSet = "UTF-8"
                .Open
                .LoadFromFile CSV.FilePath
                i = 0
                Do While Not (.EOS)
                    'aryRec(i) = .ReadText(adReadLine)
                    aryRec(i) = .ReadText(-2)
                    i = i + 1
                Loop
                .Close
                strRec = Join(aryRec, vbCrLf)
            End With
        Case "UTF-16 LE", "UTF-16 BE"
            'Set inTS = objFSO.OpenTextFile(strFile, , , TristateTrue)
            Set inTS = objFSO.OpenTextFile(CSV.FilePath, , , -1)
            strRec = inTS.ReadAll
            inTS.Close
        Case "SHIFT-JIS"
            Set inTS = objFSO.OpenTextFile(CSV.FilePath)
            strRec = inTS.ReadAll
            inTS.Close
        Case Else
            'EUC-JP、UTF-32については未テスト
            MsgBox "文字コードを確認してください。" & vbLf & CSV.CharSet
            Stop
    End Select
    Set inTS = Nothing
    Set objFSO = Nothing
    readCsv = strRec
End Function

'文字コードの自動判別
Private Function getCharSet() As String
    Dim bytes() As Byte
    Dim intFileNo As Integer
    ReDim bytes(FileLen(CSV.FilePath))
    intFileNo = FreeFile
    Open CSV.FilePath For Binary As #intFileNo
    Get #intFileNo, , bytes
    Close intFileNo
 
    'BOMによる判断
    getCharSet = getCharFromBOM(bytes)
 
    'BOMなしをデータの文字コードで判別
    If getCharSet = "" Then
        getCharSet = getCharFromCode(bytes)
    End If
 
    Debug.Print CSV.FilePath & " : " & getCharSet
End Function

'BOMによる判断
Private Function getCharFromBOM(ByRef bytes() As Byte) As String
    getCharFromBOM = ""
    If UBound(bytes) < 3 Then Exit Function
 
    Select Case True
        Case bytes(0) = &HEF And _
             bytes(1) = &HBB And _
             bytes(2) = &HBF
            getCharFromBOM = "UTF-8"
            Exit Function
        Case bytes(0) = &HFF And _
             bytes(1) = &HFE
             If bytes(2) = &H0 And _
                bytes(3) = &H0 Then
                getCharFromBOM = "UTF-32 LE"
                Exit Function
            End If
            getCharFromBOM = "UTF-16 LE"
            Exit Function
        Case bytes(0) = &HFE And _
             bytes(1) = &HFF
            getCharFromBOM = "UTF-16 BE"
            Exit Function
        Case bytes(0) = &H0 And _
             bytes(1) = &H0 And _
             bytes(2) = &HFE And _
             bytes(3) = &HFF
            getCharFromBOM = "UTF-32 BE"
            Exit Function
    End Select
End Function

'BOMなしをデータの文字コードで判別
Private Function getCharFromCode(ByRef bytes() As Byte) As String
    Const bEscape As Byte = &H1B
    Const bAt As Byte = &H40
    Const bDollar As Byte = &H24
    Const bAnd As Byte = &H26
    Const bOpen As Byte = &H28
    Const bB As Byte = &H42
    Const bD As Byte = &H44
    Const bJ As Byte = &H4A
    Const bI As Byte = &H49

    Dim bLen As Long: bLen = UBound(bytes)
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
    Dim isBinary As Boolean: isBinary = False
    Dim i As Long
 
    For i = 0 To bLen - 1
        b1 = bytes(i)
        If b1 <= &H6 Or b1 = &H7F Or b1 = &HFF Then
            isBinary = True
            If b1 = &H0 And i < bLen - 1 And bytes(i + 1) <= &H7F Then
                getCharFromCode = "SHIFT-JIS"
                Exit Function
            End If
        End If
    Next
    If isBinary Then
        getCharFromCode = ""
        Exit Function
    End If

    For i = 0 To bLen - 3
        b1 = bytes(i)
        b2 = bytes(i + 1)
        b3 = bytes(i + 2)

        If b1 = bEscape Then
            If b2 = bDollar And b3 = bAt Then
                getCharFromCode = "SHIFT-JIS"
                Exit Function
            ElseIf b2 = bDollar And b3 = bB Then
                getCharFromCode = "SHIFT-JIS"
                Exit Function
            ElseIf b2 = bOpen And (b3 = bB Or b3 = bJ) Then
                getCharFromCode = "SHIFT-JIS"
                Exit Function
            ElseIf b2 = bOpen And b3 = bI Then
                getCharFromCode = "SHIFT-JIS"
                Exit Function
            End If
            If i < bLen - 3 Then
                b4 = bytes(i + 3)
                If b2 = bDollar And b3 = bOpen And b4 = bD Then
                    getCharFromCode = "SHIFT-JIS"
                    Exit Function
                End If
                If i < bLen - 5 And _
                    b2 = bAnd And b3 = bAt And b4 = bEscape And _
                    bytes(i + 4) = bDollar And bytes(i + 5) = bB Then
                    getCharFromCode = "SHIFT-JIS"
                    Exit Function
                End If
            End If
        End If
    Next

    Dim sjis As Long: sjis = 0
    Dim euc As Long: euc = 0
    Dim utf8 As Long: utf8 = 0
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If ((&H81 <= b1 And b1 <= &H9F) Or (&HE0 <= b1 And b1 <= &HFC)) And _
           ((&H40 <= b2 And b2 <= &H7E) Or (&H80 <= b2 And b2 <= &HFC)) Then
            sjis = sjis + 2
            i = i + 1
        End If
    Next
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If ((&HA1 <= b1 And b1 <= &HFE) And _
            (&HA1 <= b2 And b2 <= &HFE)) Or _
            (b1 = &H8E And (&HA1 <= b2 And b2 <= &HDF)) Then
            euc = euc + 2
            i = i + 1
        ElseIf i < bLen - 2 Then
            b3 = bytes(i + 2)
            If b1 = &H8F And (&HA1 <= b2 And b2 <= &HFE) And _
                (&HA1 <= b3 And b3 <= &HFE) Then
                euc = euc + 3
                i = i + 2
            End If
        End If
    Next
    For i = 0 To bLen - 2
        b1 = bytes(i)
        b2 = bytes(i + 1)
        If (&HC0 <= b1 And b1 <= &HDF) And _
            (&H80 <= b2 And b2 <= &HBF) Then
            utf8 = utf8 + 2
            i = i + 1
        ElseIf i < bLen - 2 Then
            b3 = bytes(i + 2)
            If (&HE0 <= b1 And b1 <= &HEF) And _
                (&H80 <= b2 And b2 <= &HBF) And _
                (&H80 <= b3 And b3 <= &HBF) Then
                utf8 = utf8 + 3
                i = i + 2
            End If
        End If
    Next
 
    Select Case True
        Case euc > sjis And euc > utf8
            getCharFromCode = "EUC-JP"
        Case utf8 > euc And utf8 > sjis
            getCharFromCode = "UTF-8N"
        Case sjis > euc And sjis > utf8
            getCharFromCode = "SHIFT-JIS"
        Case Else '判定できず
            getCharFromCode = ""
    End Select
End Function



