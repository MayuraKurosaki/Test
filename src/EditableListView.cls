VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditableListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' messages
Private Const LVM_FIRST = &H1000
Private Const LVM_HITTEST = (LVM_FIRST + 18)
Private Const LVM_SCROLL = (LVM_FIRST + 20)
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)

'API定義 [ ConnectToConnectionPoint ]
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
 
Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" _
         (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, _
         ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, _
         ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
'Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long

'' SetWindowPos() hwndInsertAfter values
'Const HWND_TOPMOST = -1
'' SetWindowPos Flags
'Const SWP_NOSIZE = &H1
'Const SWP_NOMOVE = &H2
'Const SWP_NOZORDER = &H4

'GetDeviceCapsのnIndex設定値
Private Const LOGPIXELSX As Long = 88   ' Logical pixels/inch in X
Private Const LOGPIXELSY As Long = 90   ' Logical pixels/inch in Y

Private Enum LISTVIEWITEMRECT
    LVIR_BOUNDS = 0
    LVIR_ICON = 1
    LVIR_LABEL = 2
    LVIR_SELECTBOUNDS = 3
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type POINTF
    x As Single
    y As Single
End Type

' LVM_HITTEST lParam
Private Type LVHITTESTINFO   ' was LV_HITTESTINFO
    pt As POINTAPI
    Flags As LVHT_FLAGS
    iItem As Long
'#If (WIN32_IE >= &H300) Then
    iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
'#End If
'#If (WIN32_IE >= &H600) then
    iGroup As Long
'#End If
End Type

Private Enum LVHT_FLAGS
     LVHT_NOWHERE = &H1   ' in LV client area, but not over item
     LVHT_ONITEMICON = &H2
     LVHT_ONITEMLABEL = &H4
     LVHT_ONITEMSTATEICON = &H8
     LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
    
    '  ' outside the LV's client area
     LVHT_ABOVE = &H8
     LVHT_BELOW = &H10
     LVHT_TORIGHT = &H20
     LVHT_TOLEFT = &H40
#If (WIN32_IE >= &H600) Then
    LVHT_EX_GROUP_HEADER = &H10000000
    LVHT_EX_GROUP_FOOTER = &H20000000
    LVHT_EX_GROUP_COLLAPSE = &H40000000
    LVHT_EX_GROUP_BACKGROUND = &H80000000
    LVHT_EX_GROUP_STATEICON = &H1000000
    LVHT_EX_GROUP_SUBSETLINK = &H2000000
    LVHT_EX_GROUP = (LVHT_EX_GROUP_BACKGROUND Or LVHT_EX_GROUP_COLLAPSE Or LVHT_EX_GROUP_FOOTER Or LVHT_EX_GROUP_HEADER Or LVHT_EX_GROUP_STATEICON Or LVHT_EX_GROUP_SUBSETLINK)
    LVHT_EX_ONCONTENTS = &H4000000          'On item AND not on the background
    LVHT_EX_FOOTER = &H8000000
#End If
End Enum

Private Enum EF_TYPE
    EF_TEXTBOX = 0
    EF_COMBOBOX = 1
End Enum

Private WithEvents EditableListView As MSComctlLib.ListView
Attribute EditableListView.VB_VarHelpID = -1
Private WithEvents FrameEditField As MSForms.Frame
Attribute FrameEditField.VB_VarHelpID = -1
Private WithEvents TextBoxEditField As MSForms.TextBox
Attribute TextBoxEditField.VB_VarHelpID = -1
Private WithEvents ComboBoxEditField As MSForms.ComboBox
Attribute ComboBoxEditField.VB_VarHelpID = -1

Private ListViewFieldText As String

Private Type Field
'    ListView As MSComctlLib.ListView
    ParentFrame As MSForms.Frame
    ItemIndex As Long
    SubItemIndex As Long
'    FrameEditField As MSForms.Frame
'    TextBox As MSForms.TextBox
'    ComboBox As MSForms.ComboBox
    FontSize As Integer
    hWndLV As LongPtr
'    hWndFrm As LongPtr
    lvx As Single
    lvy As Single
    EditArea As RECT
    Initialized As Boolean
    Cookie As Long
End Type

Private This As Field

Private Sub Class_Initialize()
'    Set This.ListView = New MSComctlLib.ListView
'    Set This.FrameEditField = New MSForms.Frame
'    Set This.TextBox = New MSForms.TextBox
'    Set This.ComboBox = New MSForms.ComboBox
    This.Initialized = False
End Sub

Private Sub Class_Terminate()
'    Set This.TextBox = Nothing
'    Set This.ComboBox = Nothing
'    Set This.FrameEditField = Nothing
'    Set This.ListView = Nothing
    If (This.Cookie <> 0) Then
        Call ConnectEvent(False)
    End If
    
    If This.Initialized Then
        FrameEditField.Controls.Remove "TextBoxEditField"
        FrameEditField.Controls.Remove "ComboBoxEditField"
        This.ParentFrame.Controls.Remove "FrameEditField"
        This.ParentFrame.Controls.Remove "EditableListView"
        Set This.ParentFrame = Nothing
    End If
    
    This.Initialized = False
End Sub

Public Property Let ItemIndex(RHS As Long)
    This.ItemIndex = RHS
End Property

Public Property Get ItemIndex() As Long
    ItemIndex = This.ItemIndex
End Property

Public Property Let SubItemIndex(RHS As Long)
    This.SubItemIndex = RHS
End Property

Public Property Get SubItemIndex() As Long
    SubItemIndex = This.SubItemIndex
End Property

Public Property Get ListView() As MSComctlLib.ListView
Attribute ListView.VB_UserMemId = 0
'Attribute ListView.VB_UserMemId = 0
    'Set ListView = This.ListView
    Set ListView = EditableListView
End Property

Public Sub Init(ParentFrame As MSForms.Frame, Optional FontSize As Integer)
    Set This.ParentFrame = ParentFrame
'    Set This.ListView = This.ParentFrame.Controls.Add("MSComctlLib.ListViewCtrl", "EditableListView")
    Set EditableListView = This.ParentFrame.Controls.Add("MSComctlLib.ListViewCtrl", "EditableListView")
    Set FrameEditField = This.ParentFrame.Controls.Add("Forms.Frame.1", "FrameEditField")
    Set TextBoxEditField = FrameEditField.Controls.Add("Forms.TextBox.1", "TextBoxEditField")
    Set ComboBoxEditField = FrameEditField.Controls.Add("Forms.ComboBox.1", "ComboBoxEditField")

    This.FontSize = FontSize
'    Call WindowFromAccessibleObject(FrameEditField, This.hWndFrm)
    
'    With This.ListView
    With EditableListView
        .Font.Size = FontSize
        .View = lvwReport           '表示
        .LabelEdit = lvwManual      'ラベルの編集
        .HideSelection = False      '選択の自動解除
        .AllowColumnReorder = False '列の移動禁止
        .FullRowSelect = True       '行全体を選択
        .Gridlines = True           'グリッド線
        .FlatScrollBar = False       'trueは無し
'        .FlatScrollBar = True       'trueは無し
        .Appearance = ccFlat
        .Top = 0
        .Left = 0
        .Height = This.ParentFrame.Height
        .Width = This.ParentFrame.Width
       
        '列見出し
        .ColumnHeaders.Add , "col1", "1"
        .ColumnHeaders.Add , "col2", "2"
        .ColumnHeaders.Add , "col3", "3"
        .ColumnHeaders.Add , "col4", "4"
        .ColumnHeaders.Add , "col5", "5"
        .ColumnHeaders.Add , "col6", "6"
        This.hWndLV = .hwnd
    End With
    
    With FrameEditField
        .Font.Size = FontSize
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbRed
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
        .Visible = False '最初は非表示
'        Call SetWindowPos(This.hWndFrm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER)
    End With
    
    With TextBoxEditField
        .Font.Size = FontSize
        .BorderStyle = fmBorderStyleNone
        .Text = ""
        .SpecialEffect = fmSpecialEffectFlat
        .Height = FrameEditField.Height
        .Width = FrameEditField.Width
        .Top = 0
        .Left = 0
        .Visible = False '最初は非表示
    End With
    
    With ComboBoxEditField
        .Font.Size = FontSize
        .BorderStyle = fmBorderStyleNone
        .Text = ""
        .SpecialEffect = fmSpecialEffectFlat
        .Height = FrameEditField.Height
        .Width = FrameEditField.Width
        .Top = 0
        .Left = 0
        .Visible = False '最初は非表示
    End With
    
    This.Initialized = True
    Call EditableListView_Initialize
End Sub

Private Sub EditableListView_Click()
'Attribute EditableListView_Click.VB_UserMemId = -600

'    Dim pntPos As POINTAPI
'    Dim sglPosX, sglPosY As Single
'
'    Dim intColWid As Integer
'    Dim intColCnt As Integer
'
'    'sglLVposL Frame1の左端位置
'    'sglLVposR Frame1の右端位置
'    'sglLVposA セル列の左端
'    'sglLVposB セル列の右端
'    Dim sglLVposL, sglLVposR, sglLVposA, sglLVposB As Single
'
'    Dim i As Long
'
'    'スクロールバーを一度消さないと、なぜかバーが初期値になる←を解除
'    This.ParentFrame.ScrollBars = fmScrollBarsBoth
'
'    Call GetCursorPos(pntPos)
'
'    'ピクセルからポイントへ変換
'    sglPosX = pntPos.x * 0.75
'    sglPosY = pntPos.y * 0.75
'
'    intColCnt = ListViewEdit.ColumnHeaders.Count
'
'    'Frame1内の左端位置（最後の +5 は枠幅）
'    sglLVposL = FormListViewEdit.Left + FrameListView.Left + ListViewEdit.Left + 5
'
'    'Frame1内の右端位置（最後の -5 は枠幅 -13はスクロールバー）
'    sglLVposR = sglLVposL + FrameListView.Width - 5 - 13
'
'    '1列目のセル左端セット（スクロールに対応）
'    sglLVposA = sglLVposL - FrameListView.ScrollLeft
'
'    '列幅合計用
'    intColWid = 0
'
'    For i = 1 To intColCnt
'        sglLVposB = sglLVposA + ListViewEdit.ColumnHeaders(i).Width
'
'        'マウス座標が列の幅内か判断
'        If (sglLVposA < sglPosX) And (sglLVposB > sglPosX) Then
'
'            If sglLVposA < sglLVposL Then
'                'スクロールでクリックセルが左端よりでてる
'                FrameListView.ScrollLeft = FrameListView.ScrollLeft - (sglLVposL - sglLVposA)
'            ElseIf (sglLVposB > sglLVposR) And ((sglLVposB - sglLVposA) < (sglLVposR - sglLVposL)) Then
'                'スクロールでクリックセルが右端よりでて、列幅がフレーム幅より小さい
'                FrameListView.ScrollLeft = FrameListView.ScrollLeft + (sglLVposB - sglLVposR)
'            End If
'
'            '入力用テキストボックスサイズ
'            TextBoxListViewEdit.Width = sglLVposB - sglLVposA
'            FrameTextBox.Width = sglLVposB - sglLVposA
'
'            'セル番地をグローバル変数に格納
'            GLB_INT_COL = i - 1
'            GLB_INT_ROW = ListViewEdit.SelectedItem.index
'
'            '列の座標（最後の +2 は微調整）
'            FrameTextBox.Left = intColWid + 2
'            '行の座標　見出厚+3
'            FrameTextBox.Top = 3 + (ListViewEdit.SelectedItem.index * (CST_INT_FSIZE + CST_INT_FCFCT))
'
'            '入力用ボックスの値セット
'            If GLB_INT_COL = 0 Then
'                TextBoxListViewEdit.Text = CStr(ListViewEdit.ListItems(GLB_INT_ROW).Text)
'            Else
'                TextBoxListViewEdit.Text = CStr(ListViewEdit.ListItems(GLB_INT_ROW).SubItems(GLB_INT_COL))
'            End If
'
'            '入力用ボックス表示
'            FrameTextBox.Visible = True
'        End If
'        sglLVposA = sglLVposB
'        intColWid = intColWid + ListViewEdit.ColumnHeaders(i).Width
'    Next i
'    '列幅、行追加対応
'    Call subFrmSclSet

'    With ListView_HitTest
'        This.ItemIndex = .iItem
'        This.SubItemIndex = .iSubItem
'    End With
'
'    This.ParentFrame.Parent.LabelRowMon = This.ItemIndex
'    This.ParentFrame.Parent.LabelColumnMon = This.SubItemIndex
'
'    If ListView_GetSubItemRect(This.ItemIndex, This.SubItemIndex, LVIR_BOUNDS, This.EditArea) <> 0 Then
'        This.ParentFrame.Parent.LabelTopMon = This.EditArea.Top
'        This.ParentFrame.Parent.LabelLeftMon = This.EditArea.Left
'        This.ParentFrame.Parent.LabelHeightMon = This.EditArea.Bottom - This.EditArea.Top
'        This.ParentFrame.Parent.LabelWidthMon = This.EditArea.Right - This.EditArea.Left
'    End If
'    Call ShowEditField(This.EditArea, EF_TYPE.EF_TEXTBOX)
    Call ShowEditField(EF_TYPE.EF_TEXTBOX)
'    Call ShowEditField(This.EditArea, EF_TYPE.EF_COMBOBOX)
End Sub

Private Sub EditableListView_Initialize()
    Dim i As Integer

    For i = 1 To 100
        '1列目
        EditableListView.ListItems.Add Text:=CStr(i)
        '2列目
        EditableListView.ListItems(i).SubItems(1) = "aaa"
        '3列目
        EditableListView.ListItems(i).SubItems(2) = "あああ"
'        '1列目
'        This.ListView.ListItems.Add Text:=CStr(i)
'        '2列目
'        This.ListView.ListItems(i).SubItems(1) = "aaa"
'        '3列目
'        This.ListView.ListItems(i).SubItems(2) = "あああ"

        '文字に色付ける場合
        'EditableListView.ListItems(i).ForeColor = "&H00FF00"
        'EditableListView.ListItems(i).ListSubItems(1).ForeColor = "&H00FF00"
        'EditableListView.ListItems(i).ListSubItems(2).ForeColor = "&H00FF00"
    Next

    '見出厚+3　　行数（+2は見出と1行余裕をみる）
'    EditableListView.Height = 3 + ((EditableListView.ListItems.Count + 2) * (This.FontSize + 0.5))
'    This.ListView.Height = 3 + ((This.ListView.ListItems.Count + 2) * (This.FontSize + 0.5))
    Call ConnectEvent(True)
End Sub

Private Sub ConnectEvent(ByVal Connect As Boolean)
    Dim IID_IDispatch As GUID

    ' GUID {00020400-0000-0000-C000000000000046}
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

'    Call ConnectToConnectionPoint(Me, _
'                                  IID_IDispatch, _
'                                  Connect, _
'                                  EditableListView, _
'                                  This.Cookie, _
'                                  0&)
    Call ConnectToConnectionPoint(Me, _
                                  IID_IDispatch, _
                                  Connect, _
                                  TextBoxEditField, _
                                  This.Cookie, _
                                  0&)
                                  
'    Call ConnectToConnectionPoint(Me, _
'                                  IID_IDispatch, _
'                                  Connect, _
'                                  This.ListView, _
'                                  This.Cookie, _
'                                  0&)
    If Connect Then Debug.Print "Connect Event"
End Sub

'----------------------------------
Private Sub EditableListView_AfterLabelEdit(Cancel As Integer, NewString As String)

End Sub

Private Sub EditableListView_AfterUpdate()

End Sub

Private Sub EditableListView_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub EditableListView_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub EditableListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Debug.Print "EditableListView_ColumnClick"
End Sub

Private Sub EditableListView_DblClick()

End Sub

Private Sub EditableListView_Enter()
    Debug.Print "EditableListView_Enter"
'    This.ParentFrame.ScrollBars = fmScrollBarsNone
End Sub

Private Sub EditableListView_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "EditableListView_Exit"
End Sub

Private Sub EditableListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub EditableListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Debug.Print "EditableListView_ItemClick"
End Sub

Private Sub EditableListView_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

End Sub

Private Sub EditableListView_KeyPress(KeyAscii As Integer)

End Sub

Private Sub EditableListView_KeyUp(KeyCode As Integer, ByVal Shift As Integer)

End Sub

Private Sub EditableListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub EditableListView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
   This.lvx = x
   This.lvy = y
End Sub

Private Sub EditableListView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub EditableListView_OLECompleteDrag(Effect As Long)

End Sub

Private Sub EditableListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub EditableListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

End Sub

Private Sub EditableListView_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

End Sub

Private Sub EditableListView_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)

End Sub

Private Sub EditableListView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

End Sub

'Private Sub EditableListView_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
'    Debug.Print "EditableListView_Scroll"
'End Sub
'
Public Sub FrameEditField_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute FrameEditField_Exit.VB_UserMemId = -2147384829
'Attribute FrameEditField_Exit.VB_UserMemId = -2147384829
    Debug.Print "FrameEditField_Exit"
    FrameEditField.Visible = False
End Sub

Public Sub FrameEditField_Enter()
Attribute FrameEditField_Enter.VB_UserMemId = -2147384830
'Attribute FrameEditField_Enter.VB_UserMemId = -2147384830
    Debug.Print "FrameEditField_Enter"
End Sub

Public Sub TextBoxEditField_Enter()
Attribute TextBoxEditField_Enter.VB_UserMemId = -2147384830
'Attribute TextBoxEditField_Enter.VB_UserMemId = -2147384830
    Debug.Print "TextBoxEditField_Enter"
End Sub

Public Sub TextBoxEditField_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute TextBoxEditField_Exit.VB_UserMemId = -2147384829
'Attribute TextBoxEditField_Exit.VB_UserMemId = -2147384829
    Debug.Print "TextBoxEditField_Exit"
End Sub

Private Sub TextBoxEditField_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If This.SubItemIndex = 0 Then
                EditableListView.ListItems(This.ItemIndex + 1).Text = ListViewFieldText
            Else
                EditableListView.ListItems(This.ItemIndex + 1).SubItems(This.SubItemIndex) = ListViewFieldText
            End If
            FrameEditField.Visible = False
        Case vbKeyReturn
            FrameEditField.Visible = False
        Case vbKeyLeft
        Case vbKeyUp
        Case vbKeyRight
        Case vbKeyDown
        
    End Select
End Sub

Private Sub TextBoxEditField_AfterUpdate()
'Attribute TextBoxEditField_AfterUpdate.VB_UserMemId = -2147384832
    Debug.Print "TextBoxEditField_AfterUpdate"
End Sub

Private Sub TextBoxEditField_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Attribute TextBoxEditField_BeforeUpdate.VB_UserMemId = -2147384831
    Debug.Print "TextBoxEditField_BeforeUpdate"
End Sub

'入力ボックスの入力イベント
Private Sub TextBoxEditField_Change()
    If This.SubItemIndex = 0 Then
        EditableListView.ListItems(This.ItemIndex + 1).Text = TextBoxEditField.Text
    Else
        EditableListView.ListItems(This.ItemIndex + 1).SubItems(This.SubItemIndex) = TextBoxEditField.Text
    End If
End Sub

Private Function ListView_HitTest() As LVHITTESTINFO
    Dim HTI As LVHITTESTINFO

    With HTI
        .pt.x = This.lvx
        .pt.y = This.lvy
        .Flags = LVHT_ONITEM
    End With

    Call SendMessage(This.hWndLV, LVM_SUBITEMHITTEST, 0, HTI)
    
    ListView_HitTest = HTI
End Function

'#define ListView_GetSubItemRect(hwnd, iItem, iSubItem, code, prc) \
'        (BOOL)SNDMSG((hwnd), LVM_GETSUBITEMRECT, (WPARAM)(int)(iItem), \
'                ((prc) ? ((((LPRECT)(prc))->top = (iSubItem)), (((LPRECT)(prc))->left = (code)), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))
Private Function ListView_GetSubItemRect(ByVal ItemIndex As Long, ByVal SubItemIndex As Long, ByVal Code As LISTVIEWITEMRECT, lpRect As RECT) As Long
    With lpRect
        .Top = SubItemIndex
        .Left = Code
    End With
    
    ListView_GetSubItemRect = CLng(SendMessage(This.hWndLV, LVM_GETSUBITEMRECT, CLngPtr(This.ItemIndex), lpRect))
End Function

'#define ListView_GetColumnWidth(hwnd, iCol) \
'    (int)SNDMSG((hwnd), LVM_GETCOLUMNWIDTH, (WPARAM)(int)(iCol), 0)
Private Function ListView_GetColumnWidth(ByVal SubItemIndex As Long) As Long
    ListView_GetColumnWidth = CLng(SendMessage(This.hWndLV, LVM_GETCOLUMNWIDTH, SubItemIndex, 0&))
End Function

'#define ListView_Scroll(hwndLV, dx, dy) \
'    (BOOL)SNDMSG((hwndLV), LVM_SCROLL, (WPARAM)(int)(dx), (LPARAM)(int)(dy))
Private Function ListView_Scroll(ByVal dx As Long, ByVal dy As Long) As Long
    ListView_Scroll = CLng(SendMessage(This.hWndLV, LVM_SCROLL, dx, dy))
End Function

'Private Sub ShowEditField(FieldRect As RECT, EditFiledType As EF_TYPE)
Private Sub ShowEditField(EditFiledType As EF_TYPE)
    With ListView_HitTest
        This.ItemIndex = .iItem
        This.SubItemIndex = .iSubItem
    End With
    
    This.ParentFrame.Parent.LabelRowMon = This.ItemIndex
    This.ParentFrame.Parent.LabelColumnMon = This.SubItemIndex

    If ListView_GetSubItemRect(This.ItemIndex, This.SubItemIndex, LVIR_BOUNDS, This.EditArea) <> 0 Then
        This.ParentFrame.Parent.LabelTopMon = This.EditArea.Top
        This.ParentFrame.Parent.LabelLeftMon = This.EditArea.Left
        This.ParentFrame.Parent.LabelHeightMon = This.EditArea.Bottom - This.EditArea.Top
        This.ParentFrame.Parent.LabelWidthMon = This.EditArea.Right - This.EditArea.Left
    End If
    
    Dim FieldTop As Single, FieldLeft As Single, FieldHeight As Single, FieldWidth As Single
    Dim FieldPosPixcel As POINTAPI
    FieldPosPixcel.x = This.EditArea.Left + 1
    FieldPosPixcel.y = This.EditArea.Top + 2
'    FieldPosPixcel.x = FieldRect.Left + 1
'    FieldPosPixcel.y = FieldRect.Top + 2
    With PixcelToPoint(FieldPosPixcel)
        FieldTop = .y
        FieldLeft = .x
    End With
    
    FieldPosPixcel.x = ListView_GetColumnWidth(This.SubItemIndex) + 1
    FieldPosPixcel.y = This.EditArea.Bottom + 1
'    FieldPosPixcel.y = FieldRect.Bottom + 1
'    This.ParentFrame.Parent.LabelColumnWidthMon = ListView_GetColumnWidth(This.SubItemIndex)
    With PixcelToPoint(FieldPosPixcel)
        FieldHeight = .y - FieldTop
        FieldWidth = .x
    End With
    
'    Dim FieldText As String
    If This.SubItemIndex = 0 Then
        ListViewFieldText = EditableListView.ListItems(This.ItemIndex + 1).Text
    ElseIf This.SubItemIndex > EditableListView.ListItems(This.ItemIndex + 1).ListSubItems.Count Then
        ListViewFieldText = ""
    Else
        ListViewFieldText = EditableListView.ListItems(This.ItemIndex + 1).ListSubItems(This.SubItemIndex).Text
    End If

    With FrameEditField
        .Top = FieldTop
        .Left = FieldLeft
        .Height = FieldHeight
        .Width = FieldWidth
        .Visible = True
'        .SetFocus
'        Debug.Print "FrameEditField_GotFocus"
    End With
    
    Select Case EditFiledType
        Case EF_TEXTBOX
            ComboBoxEditField.Visible = False
            With TextBoxEditField
                .Height = FieldHeight
                .Width = FieldWidth
                .Text = ListViewFieldText
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)

            End With
'            Debug.Print "TextBoxEditField_GotFocus"
        Case EF_COMBOBOX
            TextBoxEditField.Visible = False
            With ComboBoxEditField
                .Height = FieldHeight
                .Width = FieldWidth
                .Text = ListViewFieldText
                .Visible = True
'                .SetFocus
            End With
'            Debug.Print "TextBoxEditField_GotFocus"
    End Select

'    Call SetWindowPos(This.hWndFrm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    EditableListView.SetFocus
    SendKeys vbTab

'    With FrameEditField
'        .Top = FieldTop
'        .Left = FieldLeft
'        .Height = FieldHeight
'        .Width = FieldWidth
'        .Visible = True
''        .SetFocus
''        Debug.Print "FrameEditField_GotFocus"
'    End With
End Sub

'ピクセルをポイントに変換
Private Function PixcelToPoint(Pixcel As POINTAPI) As POINTF
    PixcelToPoint.x = (Pixcel.x * 72 / LogicalPixcel.x)
    PixcelToPoint.y = (Pixcel.y * 72 / LogicalPixcel.y)
End Function

'DPIを取得：ディスプレイの拡大率込
Private Function LogicalPixcel() As POINTAPI
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)

    LogicalPixcel.x = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    LogicalPixcel.y = GetDeviceCaps(hDCDesk, LOGPIXELSY)

    Call ReleaseDC(hWndDesk, hDCDesk)
End Function

