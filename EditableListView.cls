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

Private Const WM_SETFOCUS = 7&
Private Const WM_NOTIFY = &H4E&
Private Const WM_MOUSEACTIVATE = &H21&
Private Const WM_NCDESTROY = &H82&
Private Const WM_SETFONT = &H30
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_MOUSEWHEEL As Long = &H20A

Private Const WHEEL_DELTA = 120

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
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
'Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function InitCommonControlsEx Lib "COMCTL32" (ByRef LPINITCOMMONCONTROLSEX As InitCommonControlsExType) As Long
Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SysAllocString Lib "OleAut32.dll" (ByVal psz As LongPtr) As LongPtr

Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As LongPtr

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

Private Type InitCommonControlsExType
     dwSize As Long
     dwICC As Long
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
    hwndLV As LongPtr
    hWndFrm As LongPtr
    pfn As LongPtr
    pfnFrm As LongPtr
    pfnLV As LongPtr
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
    
    Call RemoveSubClass
    
    If This.Initialized Then
        FrameEditField.Controls.Remove "TextBoxEditField"
        FrameEditField.Controls.Remove "ComboBoxEditField"
        This.ParentFrame.Controls.Remove "FrameEditField"
        This.ParentFrame.Controls.Remove "EditableListView"
        Set This.ParentFrame = Nothing
    End If
    
    This.Initialized = False
End Sub
'--------------Enums----------------
Public Enum AppearanceConstants
    ccFlat = 0
    cc3D = 1
End Enum

Public Enum ListArrangeConstants
    lvwNone = 0
    lvwAutoLeft = 1
    lvwAutoTop = 2
End Enum

Public Enum BorderStyleConstants
    ccNone = 0
    ccFixedSingle = 1
End Enum

Public Enum ListLabelEditConstants
    lvwAutomatic = 0
    lvwManual = 1
End Enum

Public Enum MousePointerConstants
    ccDefault = 0
    ccArrow = 1
    ccCross = 2
    ccIBeam = 3
    ccIcon = 4
    ccSize = 5
    ccSizeNESW = 6
    ccSizeNS = 7
    ccSizeNWSE = 8
    ccSizeEW = 9
    ccUpArrow = 10
    ccHourglass = 11
    ccNoDrop = 12
    ccArrowHourglass = 13
    ccArrowQuestion = 14
    ccSizeAll = 15
    ccCustom = 99 
End Enum

Public Enum OLEDragConstants
    ccOLEDragManual = 0
    ccOLEDragAutomatic = 1
End Enum

Public Enum OLEDropConstants
    ccOLEDropNone = 0
    ccOLEDropManual = 1
End Enum

Public Enum ListPictureAlignmentConstants
    lvwTopLeft = 0
    lvwTopRight = 1
    lvwBottomLeft = 2
    lvwBottomRight = 3
    lvwCenter = 4
    lvwTile = 5
End Enum

Public Enum ListSortOrderConstants
    lvwAscending = 1
    lvwDescending = 2
End Enum

Public Enum ListTextBackgroundConstants
    lvwTransparent = 0
    lvwOpaque = 1
End Enum

Public Enum ListViewConstants
    lvwIcon = 0
    lvwSmallIcon = 1
    lvwList = 2
    lvwReport = 3
 End Enum










'--------------Compatible Properties----------------
'Returns/sets whether a user can reorder columns in report view.
Public Property Get AllowColumnReorder() As Boolean
    AllowColumnReorder = EditableListView.AllowColumnReorder
End Property
Public Property Let AllowColumnReorder(RHS As Boolean)
    EditableListView.AllowColumnReorder = RHS
End Property

'Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects.
Public Property Get Appearance() As AppearanceConstants
    Appearance = EditableListView.Appearance
End Property
Public Property Let Appearance(RHS As AppearanceConstants)
    EditableListView.Appearance = RHS
End Property

'Returns/sets how the icons in a ListView control's Icon or SmallIcon view are arranged.
Public Property Get Arrange() As ListArrangeConstants
    Arrange = EditableListView.Arrange
End Property
Public Property Let Arrange(RHS As ListArrangeConstants)
    EditableListView.Arrange = RHS
End Property

'Returns/sets the background color used to display text and graphics in an object.
Public Property Get BackColor() As OLE_COLOR
    BackColor = EditableListView.BackColor
End Property
Public Property Let BackColor(RHS As OLE_COLOR)
    EditableListView.BackColor = RHS
End Property

'Returns/sets the border style for an object.
Public Property Get BorderStyle As BorderStyleConstants
    BorderStyle = EditableListView.BorderStyle
End Property
Public Property Let BorderStyle(RHS As BorderStyleConstants)
    EditableListView.BorderStyle = RHS
End Property

'Returns/sets a value which determines if the control displays a checkbox next to each item in the list.
Public Property Get CheckBoxes() As Boolean
    CheckBoxes = EditableListView.CheckBoxes
End Property
Public Property Let CheckBoxes(RHS As Boolean)
    EditableListView.CheckBoxes = RHS
End Property

'Returns/sets the ImageList control to be used for ColumnHeader icons.
Public Property Get ColumnHeaderIcons() As Object
    Set ColumnHeaderIcons = EditableListView.ColumnHeaderIcons
End Property
Public Property Set ColumnHeaderIcons(RHS As Object)
    Set EditableListView.ColumnHeaderIcons = RHS
End Property

'Returns a reference to a collection of ColumnHeader objects.
Public Property Get ColumnHeaders() As ColumnHeaders
    Set ColumnHeaders = EditableListView.ColumnHeaders
End Property

'Returns a reference to a Node or ListItem object and highlights the object with the system highlight color.
Public Property Get DropHighlight() As ListItem
    Set DropHighlight = EditableListView.DropHighlight
End Property

'Returns/sets a value that determines whether a form or control can respond to user-generated events.
Public Property Get Enabled() As Boolean
    Enabled = EditableListView.Enabled
End Property
Public Property Let Enabled(RHS As Boolean)
    EditableListView.Enabled = RHS
End Property

'Returns/sets whether the scrollbars appear flat.
Public Property Get FlatScrollBar() As Boolean
    FlatScrollBar = EditableListView.FlatScrollBar
End Property
Public Property Let FlatScrollBar(RHS As Boolean)
    EditableListView.FlatScrollBar = RHS
End Property

'Returns a Font object.
Public Property Get Font() As StdFont
    Set Font = EditableListView.Font
End Property

'Returns/sets the background color used to display text and graphics in an object.
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = EditableListView.ForeColor
End Property
Public Property Let ForeColor(RHS As OLE_COLOR)
    EditableListView.ForeColor = RHS
End Property

'Returns/sets whether selecting a column highlights the entire row.
Public Property Get FullRowSelect() As Boolean
    FullRowSelect = EditableListView.FullRowSelect
End Property
Public Property Let FullRowSelect(RHS As Boolean)
    EditableListView.FullRowSelect = RHS
End Property

'Returns/sets whether grid lines appear between rows and columns
Public Property Get Gridlines() As Boolean
    Gridlines = EditableListView.Gridlines
End Property
Public Property Let Gridlines(RHS As Boolean)
    EditableListView.Gridlines = RHS
End Property

'Returns/sets whether or not a ListView control's column headers are hidden in Report view.
Public Property Get HideColumnHeaders() As Boolean
    HideColumnHeaders = EditableListView.HideColumnHeaders
End Property
Public Property Let HideColumnHeaders(RHS As Boolean)
    EditableListView.HideColumnHeaders = RHS
End Property

'Determines whether the selected item will display as selected when the ListView loses focus
Public Property Get HideSelection() As Boolean
    HideSelection = EditableListView.HideSelection
End Property
Public Property Let HideSelection(RHS As Boolean)
    EditableListView.HideSelection = RHS
End Property

'Returns/sets whether hot tracking is enabled.
Public Property Get HotTracking() As Boolean
    HotTracking = EditableListView.HotTracking
End Property
Public Property Let HotTracking(RHS As Boolean)
    EditableListView.HotTracking = RHS
End Property

'Returns/sets whether hover selection is enabled.
Public Property Get HoverSelection() As Boolean
    HoverSelection = EditableListView.HoverSelection
End Property
Public Property Let HoverSelection(RHS As Boolean)
    EditableListView.HoverSelection = RHS
End Property

'Returns a handle to a form or control.
Public Property Get hWnd() As LongPtr
    hWnd = EditableListView.hWnd
End Property

'Returns/sets the images associated with the Icon properties of a ListView control.
Public Property Get IconIcons() As Object
    Set IconIcons = EditableListView.IconIcons
End Property
Public Property Set IconIcons(RHS As Object)
    Set EditableListView.IconIcons = RHS
End Property

'Returns/sets a value that determines if a user can edit the label of a ListItem or Node object.
Public Property Get LabelEdit() As ListLabelEditConstants
    LabelEdit = EditableListView.LabelEdit
End Property
Public Property Let LabelEdit(RHS As ListLabelEditConstants)
    EditableListView.LabelEdit = RHS
End Property

'Returns or sets a value that determines if labels are wrapped when the ListView is in Icon view.
Public Property Get LabelWrap() As Boolean
    LabelWrap = EditableListView.LabelWrap
End Property
Public Property Let LabelWrap(RHS As Boolean)
    EditableListView.LabelWrap = RHS
End Property

'Returns a reference to a collection of ListItem objects in a ListView control.
Public Property Get ListItems() As ListItems
    Set ListItems = EditableListView.ListItems
End Property

'Sets a custom mouse icon.
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = EditableListView.MouseIcon
End Property
Public Property Set MouseIcon(RHS As StdPicture)
    Set EditableListView.MouseIcon = RHS
End Property

'Returns/sets the type of mouse pointer displayed when over part of an object.
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = EditableListView.MousePointer
End Property
Public Property Let MousePointer(RHS As MousePointerConstants)
    EditableListView.MousePointer = RHS
End Property

'Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made.
Public Property Get MultiSelect() As Boolean
    MultiSelect = EditableListView.MultiSelect
End Property
Public Property Let MultiSelect(RHS As Boolean)
    EditableListView.MultiSelect = RHS
End Property

'Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control.
Public Property Get OLEDragMode() As OLEDragConstants
    OLEDragMode = EditableListView.OLEDragMode
End Property
Public Property Let OLEDragMode(RHS As OLEDragConstants)
    EditableListView.OLEDragMode = RHS
End Property

'Returns/Sets whether this control can act as an OLE drop target.
Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = EditableListView.OLEDropMode
End Property
Public Property Let OLEDropMode(RHS As OLEDropConstants)
    EditableListView.OLEDropMode = RHS
End Property

'Returns/sets the background picture for the control.
Public Property Get Picture() As StdPicture
    Set Picture = EditableListView.Picture
End Property
Public Property Set Picture(RHS As StdPicture)
    Set EditableListView.Picture = RHS
End Property

'Returns/sets the picture alignment.
Public Property Get PictureAlignment() As ListPictureAlignmentConstants
    PictureAlignment = EditableListView.PictureAlignment
End Property
Public Property Let PictureAlignment(RHS As ListPictureAlignmentConstants)
    EditableListView.PictureAlignment = RHS
End Property

'Returns a reference to the currently selected ListItem or Node object.
Public Property Get SelectedItem() As ListItem
    Set SelectedItem = EditableListView.SelectedItem
End Property

'Returns/sets the images associated with the SmallIcons property of a ListView control.
Public Property Get SmallIcons() As Object
    Set SmallIcons = EditableListView.SmallIcons
End Property
Public Property Set SmallIcons(RHS As Object)
    Set EditableListView.SmallIcons = RHS
End Property

'Indicates whether the elements of a control are automatically sorted alphabetically.
Public Property Get Sorted() As Boolean
    Sorted = EditableListView.Sorted
End Property
Public Property Let Sorted(RHS As Boolean)
    EditableListView.Sorted = RHS
End Property

'Returns/sets the current sort key.
Public Property Get SortKey() As Integer
    SortKey = EditableListView.SortKey
End Property
Public Property Let SortKey(RHS As Integer)
    EditableListView.SortKey = RHS
End Property

'Returns/sets whether or not the ListItems will be sorted in ascending or descending order.
Public Property Get SortOrder() As ListSortOrderConstants
    SortOrder = EditableListView.SortOrder
End Property
Public Property Let SortOrder(RHS As ListSortOrderConstants)
    EditableListView.SortOrder = RHS
End Property

'Returns/sets a value that determines if the text background is transparent or uses the ListView background color
Public Property Get TextBackground() As ListTextBackgroundConstants
    TextBackground = EditableListView.TextBackground
End Property
Public Property Let TextBackground(RHS As ListTextBackgroundConstants)
    EditableListView.TextBackground = RHS
End Property

'Returns/sets the current view of the ListView control.
Public Property Get View() As ListViewConstants
    View = EditableListView.View
End Property
Public Property Let View(RHS As ListViewConstants)
    EditableListView.View = RHS
End Property

'--------------Extended Properties----------------
'Returns/sets the current item index.
Public Property Get ItemIndex() As Long
    ItemIndex = EditableListView.ItemIndex
End Property
Public Property Let ItemIndex(RHS As Long)
    EditableListView.ItemIndex = RHS
End Property

'Returns/sets the current subitem index.
Public Property Get SubItemIndex() As Long
    SubItemIndex = EditableListView.SubItemIndex
End Property
Public Property Let SubItemIndex(RHS As Long)
    EditableListView.SubItemIndex = RHS
End Property

'--------------Compatible Events----------------
'Occurs after a user edits the label of the currently selected Node or ListItem object.
Public Event AfterLabelEdit(ByVal Cancel As Boolean, NewString As String)
Private Sub EditableListView_AfterLabelEdit(ByVal Cancel As Boolean, NewString As String)
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

'Occurs when a user attempts to edit the label of the currently selected ListItem or Node object.
Public Event BeforeLabelEdit(ByVal Cancel As Boolean)
Private Sub EditableListView_BeforeLabelEdit(ByVal Cancel As Boolean)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

'Occurs when the user presses and then releases a mouse button over an object.
Public Event Click()
Private Sub EditableListView_Click()
    RaiseEvent Click()
End Sub

'Occurs when a ColumnHeader object in a ListView control is clicked.
Public Event ColumnClick(ByVal ColumnHeader As ColumnHeader)
Private Sub EditableListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim ch As ColumnHeader
    Set ch = New ColumnHeader
    Set ch = ColumnHeader
    RaiseEvent ColumnClick(ch)
End Sub

'Occurs when you press and release a mouse button and then press and release it again over an object.
Public Event DblClick()
Private Sub EditableListView_DblClick()
    RaiseEvent DblClick()
End Sub

'Occurs when a ListSubItem object is checked
Public Event ItemCheck(ByVal ListItem As ListItem)
Private Sub EditableListView_ItemCheck(ByVal ListItem As MSComctlLib.ListItem)
    Dim li As ListItem
    Set li = New ListItem
    Set li = ListItem
    RaiseEvent ItemCheck(li)
End Sub

'Occurs when a ListItem object is clicked or selected
Public Event ItemClick(ByVal ListItem As ListItem)
Private Sub EditableListView_ItemClick(ByVal ListItem As MSComctlLib.ListItem)
    Dim li As ListItem
    Set li = New ListItem
    Set li = ListItem
    RaiseEvent ItemClick(li)
End Sub

'Occurs when the user presses a key while an object has the focus.
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Private Sub EditableListView_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'Occurs when the user presses and releases an ANSI key.
Public Event KeyPress(ByVal KeyAscii As Integer)
Private Sub EditableListView_KeyPress(ByVal KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'Occurs when the user releases a key while an object has the focus.
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Private Sub EditableListView_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'Occurs when the user presses the mouse button while an object has the focus.
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Sub EditableListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'Occurs when the user moves the mouse.
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Sub EditableListView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'Occurs when the user releases the mouse button while an object has the focus.
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Sub EditableListView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'OLECompleteDrag event
Public Event OLECompleteDrag(ByVal Effect As Long)
Private Sub EditableListView_OLECompleteDrag(ByVal Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'OLEDragDrop event
Public Event OLEDragDrop(ByVal Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Sub EditableListView_OLEDragDrop(ByVal Data As MSForms.DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim dobj As DataObject
    Set dobj = New DataObject
    Set dobj = Data
    RaiseEvent OLEDragDrop(dobj, Effect, Button, Shift, x, y)
End Sub

'OLEDragOver event
Public Event OLEDragOver(ByVal Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Sub EditableListView_OLEDragOver(ByVal Data As MSForms.DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim dobj As DataObject
    Set dobj = New DataObject
    Set dobj = Data
    RaiseEvent OLEDragOver(dobj, Effect, Button, Shift, x, y)
End Sub

'OLEGiveFeedback event
Public Event OLEGiveFeedback(ByVal Effect As Long, ByVal DefaultCursors As Boolean)
Private Sub EditableListView_OLEGiveFeedback(ByVal Effect As Long, ByVal DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'OLESetData event
Public Event OLESetData(ByVal Data As DataObject, ByVal DataFormat As Integer)
Private Sub EditableListView_OLESetData(ByVal Data As MSForms.DataObject, ByVal DataFormat As Integer)
    Dim dobj As DataObject
    Set dobj = New DataObject
    Set dobj = Data
    RaiseEvent OLESetData(dobj, DataFormat)
End Sub

'OLEStartDrag event
Public Event OLEStartDrag(ByVal Data As DataObject, ByVal AllowedEffects As Long)
Private Sub EditableListView_OLEStartDrag(ByVal Data As MSForms.DataObject, ByVal AllowedEffects As Long)
    Dim dobj As DataObject
    Set dobj = New DataObject
    Set dobj = Data
    RaiseEvent OLEStartDrag(dobj, AllowedEffects)
End Sub

'--------------Extended Events----------------



'--------------Compatible Methods----------------
Public Sub AboutBox()
    EditableListView.AboutBox
End Sub

'Finds an item in the list and returns a reference to that item.
Public Function FindItem(Text As String, Optional Where As Long, Optional Index As Long, Optional fPartial As Boolean) As ListItem
    Set FindItem = EditableListView.FindItem(Text, Where, Index, fPartial)
End Function

'Retrieves a reference of the first item visible in the client area.
Public Function GetFirstVisibleItem() As ListItem
    Set GetFirstVisibleItem = EditableListView.GetFirstVisibleItem
End Function

'Returns a reference to the ListItem object or Node object located at the coordinates of x and y. Used with drag and drop operations.
Public Function HitTest(x As Single, y As Single) As ListItem
    Set HitTest = EditableListView.HitTest(x, y)
End Function

'Starts an OLE drag/drop event with the given control as the source.
Public Sub OLEDrag()
    EditableListView.OLEDrag
End Sub

'Forces a complete repaint of a form or control.
Public Sub Refresh()
    EditableListView.Refresh
End Sub

'Begins a label editing operation on a ListItem or Node object.
Public Sub StartLabelEdit()
    EditableListView.StartLabelEdit
End Sub

'--------------Extended Methods----------------
'Gets or sets the type of action the user must take to activate an item.
Public Property Get ActivationType() As Long
    ActivationType = This.ActivationType
End Property
Public Property Let ActivationType(RHS As Long)
    This.ActivationType = RHS
End Property

'Gets or sets the alignment of items in the control.
Public Property Get ItemAlignment() As Long
    ItemAlignment = This.ItemAlignment
End Property
Public Property Let ItemAlignment(RHS As Long)
    This.ItemAlignment = RHS
End Property














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
    Call WindowFromAccessibleObject(ParentFrame, This.hWndFrm)
    
    With This
        .pfn = GetPtr(AddressOf Redirect)
        SetWindowSubclass .hWndFrm, .pfn, .hWndFrm, ObjPtr(Me)

        .hwndLV = EditableListView.hwnd
        SetWindowSubclass .hwndLV, .pfn, .hwndLV, ObjPtr(Me)
    End With

'    With This
'        .pfnFrm = GetPtr(AddressOf RedirectFrm)
''        .pfnFrm = GetPtr(AddressOf ListViewEditModule.WndProcFrm)
'        SetWindowSubclass .hWndFrm, .pfnFrm, .hWndFrm, ObjPtr(Me)
'
'        .hwndLV = EditableListView.hWnd
'        .pfnLV = GetPtr(AddressOf RedirectLV)
''        .pfnLV = GetPtr(AddressOf ListViewEditModule.WndProcLV)
'        SetWindowSubclass .hwndLV, .pfnLV, .hwndLV, ObjPtr(Me)
'    End With

'    With This.ParentFrame
''        .ScrollBars = fmScrollBarsBoth
'    End With
    
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
'    With This
'        .pfn = GetPtr(AddressOf Redirect)
'        SetWindowSubclass .hWndFrm, .pfn, .hWndFrm, ObjPtr(Me)
'
'        .hwndLV = EditableListView.hWnd
'        SetWindowSubclass .hwndLV, .pfn, .hwndLV, ObjPtr(Me)
'    End With
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
    
    Call ShowEditField(EF_TYPE.EF_TEXTBOX)
'    Call ShowEditField(EF_TYPE.EF_COMBOBOX)
End Sub

Private Sub EditableListView_Initialize()
    Dim I As Integer

    For I = 1 To 100
        '1列目
        EditableListView.ListItems.Add Text:=CStr(I)
        '2列目
        EditableListView.ListItems(I).SubItems(1) = "aaa"
        '3列目
        EditableListView.ListItems(I).SubItems(2) = "あああ"

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
    Debug.Print "EditableListView_AfterLabelEdit"
End Sub

Private Sub EditableListView_AfterUpdate()
    Debug.Print "EditableListView_AfterUpdate"
End Sub

Private Sub EditableListView_BeforeLabelEdit(Cancel As Integer)
    Debug.Print "EditableListView_BeforeLabelEdit"
End Sub

Private Sub EditableListView_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "EditableListView_BeforeUpdate"
End Sub

Private Sub EditableListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Debug.Print "EditableListView_ColumnClick"
End Sub

Private Sub EditableListView_DblClick()
    Debug.Print "EditableListView_DblClick"
End Sub

Private Sub EditableListView_Enter()
    Debug.Print "EditableListView_Enter"
'    This.ParentFrame.ScrollBars = fmScrollBarsNone
End Sub

Private Sub EditableListView_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "EditableListView_Exit"
End Sub

Private Sub EditableListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Debug.Print "EditableListView_ItemCheck"
End Sub

Private Sub EditableListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Debug.Print "EditableListView_ItemClick"
End Sub

Private Sub EditableListView_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Debug.Print "EditableListView_KeyDown"
End Sub

Private Sub EditableListView_KeyPress(KeyAscii As Integer)
    Debug.Print "EditableListView_KeyPress"
End Sub

Private Sub EditableListView_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Debug.Print "EditableListView_KeyUp"
End Sub

Private Sub EditableListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    Debug.Print "EditableListView_MouseDown"
End Sub

Private Sub EditableListView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
'    Debug.Print "EditableListView_MouseMove"
   This.lvx = x
   This.lvy = y
End Sub

Private Sub EditableListView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    Debug.Print "EditableListView_MouseUp"
End Sub

Private Sub EditableListView_OLECompleteDrag(Effect As Long)
    Debug.Print "EditableListView_OLECompleteDrag"
End Sub

Private Sub EditableListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "EditableListView_OLEDragDrop"
End Sub

Private Sub EditableListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Debug.Print "EditableListView_OLEDragOver"
End Sub

Private Sub EditableListView_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Debug.Print "EditableListView_OLEGiveFeedback"
End Sub

Private Sub EditableListView_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    Debug.Print "EditableListView_OLESetData"
End Sub

Private Sub EditableListView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Debug.Print "EditableListView_OLEStartDrag"
End Sub

'Private Sub EditableListView_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
''    Attribute EditableListView_Scroll.VB_UserMemId = -606
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
    Debug.Print "TextBoxEditField_KeyDown"
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

    Call SendMessage(This.hwndLV, LVM_SUBITEMHITTEST, 0, HTI)
    
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
    
    ListView_GetSubItemRect = CLng(SendMessage(This.hwndLV, LVM_GETSUBITEMRECT, ByVal This.ItemIndex, lpRect))
End Function

'#define ListView_GetColumnWidth(hwnd, iCol) \
'    (int)SNDMSG((hwnd), LVM_GETCOLUMNWIDTH, (WPARAM)(int)(iCol), 0)
Private Function ListView_GetColumnWidth(ByVal SubItemIndex As Long) As Long
    ListView_GetColumnWidth = CLng(SendMessage(This.hwndLV, LVM_GETCOLUMNWIDTH, ByVal SubItemIndex, 0&))
End Function

'#define ListView_Scroll(hwndLV, dx, dy) \
'    (BOOL)SNDMSG((hwndLV), LVM_SCROLL, (WPARAM)(int)(dx), (LPARAM)(int)(dy))
Private Function ListView_Scroll(ByVal DX As Long, ByVal DY As Long) As Long
    ListView_Scroll = CLng(SendMessage(This.hwndLV, LVM_SCROLL, ByVal DX, ByVal DY))
End Function

'Private Sub ShowEditField(FieldRect As RECT, EditFiledType As EF_TYPE)
Private Sub ShowEditField(EditFiledType As EF_TYPE)
    With ListView_HitTest
        This.ItemIndex = .iItem
        This.SubItemIndex = .iSubItem
    End With
    
    This.ParentFrame.Parent.LabelRowMon = This.ItemIndex
    This.ParentFrame.Parent.LabelColumnMon = This.SubItemIndex

    Dim lvRect As RECT
    If GetClientRect(This.hwndLV, lvRect) <> 0 Then
    Else
        Exit Sub
    End If
    Dim ScrollDx As Long
    If ListView_GetSubItemRect(This.ItemIndex, This.SubItemIndex, LVIR_BOUNDS, This.EditArea) <> 0 Then
'        If This.SubItemIndex = 0 Then
'            This.EditArea.Right = ListView_GetColumnWidth(This.SubItemIndex)
'            Debug.Print This.EditArea.Right
'        End If
        With This.EditArea
            If .Left < lvRect.Left Then
                ScrollDx = .Left - lvRect.Left
                .Right = .Right - ScrollDx
                .Left = lvRect.Left
                Call ListView_Scroll(ScrollDx, 0)
            ElseIf .Right > lvRect.Right Then
                ScrollDx = .Right - lvRect.Right
                .Right = lvRect.Right
                .Left = .Left - ScrollDx
                Call ListView_Scroll(ScrollDx, 0)
            End If
            
            This.ParentFrame.Parent.LabelTopMon = .Top
            This.ParentFrame.Parent.LabelLeftMon = .Left
            This.ParentFrame.Parent.LabelHeightMon = .Bottom - .Top
            This.ParentFrame.Parent.LabelWidthMon = .Right - .Left
        End With
    Else
    
    End If
    
    Dim FieldTop As Single, FieldLeft As Single
    Dim FieldPosPixcel As POINTAPI
    FieldPosPixcel.x = This.EditArea.Left + 1
    FieldPosPixcel.y = This.EditArea.Top + 1
    With PixcelToPoint(FieldPosPixcel)
        FieldTop = .y
        FieldLeft = .x
    End With
    
    Dim FieldBottom As Single, FieldRight As Single
    If This.SubItemIndex = 0 Then
        FieldPosPixcel.x = ListView_GetColumnWidth(This.SubItemIndex) + 1
    Else
        FieldPosPixcel.x = This.EditArea.Right + 1
    End If
'    FieldPosPixcel.X = This.EditArea.Right + 1
    FieldPosPixcel.y = This.EditArea.Bottom + 1
    With PixcelToPoint(FieldPosPixcel)
        FieldBottom = .y
        FieldRight = .x
    End With
    
    Dim FieldHeight As Single, FieldWidth As Single
    FieldHeight = FieldBottom - FieldTop
    FieldWidth = FieldRight - FieldLeft
    
'    FieldPosPixcel.X = ListView_GetColumnWidth(This.SubItemIndex) + 1
'    FieldPosPixcel.Y = This.EditArea.Bottom + 1
''    FieldPosPixcel.y = FieldRect.Bottom + 1
''    This.ParentFrame.Parent.LabelColumnWidthMon = ListView_GetColumnWidth(This.SubItemIndex)
'    With PixcelToPoint(FieldPosPixcel)
'        FieldHeight = .Y - FieldTop
'        FieldWidth = .X
'    End With
    
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
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
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

'' ============================================================
'' listview macros
'
''#define LVM_APPROXIMATEVIEWRECT (LVM_FIRST + 64)
''#define ListView_ApproximateViewRect(hwnd, iWidth, iHeight, iCount) \
''        (DWORD)SNDMSG((hwnd), LVM_APPROXIMATEVIEWRECT, (WPARAM)(iCount), MAKELPARAM(iWidth, iHeight))
'Public Function ListView_ApproximateViewRect(hwnd As Long, iWidth As Long, iHeight As Long, iCount As Long) As Long
'    ListView_ApproximateViewRect = SendMessage(hwnd, LVM_APPROXIMATEVIEWRECT, ByVal iCount, ByVal MAKELPARAM(iWidth, iHeight))
'End Function
'
''#define LVM_ARRANGE             (LVM_FIRST + 22)
''#define ListView_Arrange(hwndLV, code) \
''    (BOOL)SNDMSG((hwndLV), LVM_ARRANGE, (WPARAM)(UINT)(code), 0L)
'Public Function ListView_Arrange(hwndLV As Long, Code As LVA_Flags) As Boolean
'    ListView_Arrange = SendMessage(hwndLV, LVM_ARRANGE, ByVal Code, 0)
'End Function
'
''#define LVM_CREATEDRAGIMAGE     (LVM_FIRST + 33)
''#define ListView_CreateDragImage(hwnd, i, lpptUpLeft) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_CREATEDRAGIMAGE, (WPARAM)(int)(i), (LPARAM)(LPPOINT)(lpptUpLeft))
'Public Function ListView_CreateDragImage(hwnd As Long, i As Long, lpptUpLeft As POINT) As Long
'    ListView_CreateDragImage = SendMessage(hwnd, LVM_CREATEDRAGIMAGE, ByVal i, lpptUpLeft)
'End Function
'
''#define LVM_DELETEITEM          (LVM_FIRST + 8)
''#define ListView_DeleteItem(hwnd, i) \
''    (BOOL)SNDMSG((hwnd), LVM_DELETEITEM, (WPARAM)(int)(i), 0L)
'Public Function ListView_DeleteItem(hwnd As Long, i As Long) As Boolean
'    ListView_DeleteItem = SendMessage(hwnd, LVM_DELETEITEM, ByVal i, 0)
'End Function
'
''#define LVM_EDITLABELA          (LVM_FIRST + 23)
''#define LVM_EDITLABELW          (LVM_FIRST + 118)
''#ifdef UNICODE
''#define LVM_EDITLABEL           LVM_EDITLABELW
''#Else
''#define LVM_EDITLABEL           LVM_EDITLABELA
''#End If
''
''#define ListView_EditLabel(hwndLV, i) \
''    (HWND)SNDMSG((hwndLV), LVM_EDITLABEL, (WPARAM)(int)(i), 0L)
'Public Function ListView_EditLabel(hwndLV As Long, i As Long) As Long
'    ListView_EditLabel = SendMessage(hwndLV, LVM_EDITLABEL, ByVal i, 0)
'End Function
'
''#define LVM_GETBKCOLOR          (LVM_FIRST + 0)
''#define ListView_GetBkColor(hwnd)  \
''    (COLORREF)SNDMSG((hwnd), LVM_GETBKCOLOR, 0, 0L)
'Public Function ListView_GetBkColor(hwnd As Long) As Long
'    ListView_GetBkColor = SendMessage(hwnd, LVM_GETBKCOLOR, 0, 0)
'End Function
'
''#define LVM_SETBKCOLOR          (LVM_FIRST + 1)
''#define ListView_SetBkColor(hwnd, clrBk) \
''    (BOOL)SNDMSG((hwnd), LVM_SETBKCOLOR, 0, (LPARAM)(COLORREF)(clrBk))
'Public Function ListView_SetBkColor(hwnd As Long, clrBk As Long) As Boolean
'    ListView_SetBkColor = SendMessage(hwnd, LVM_SETBKCOLOR, 0, ByVal clrBk)
'End Function
'
''#define LV_MAX_WORKAREAS         16
''#define LVM_SETWORKAREAS         (LVM_FIRST + 65)
''#define ListView_SetWorkAreas(hwnd, nWorkAreas, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_SETWORKAREAS, (WPARAM)(int)(nWorkAreas), (LPARAM)(RECT *)(prc))
'Public Function ListView_SetWorkAreas(hwnd As Long, nWorkAreas As Long, prc() As RECT) As Boolean
'    ListView_SetWorkAreas = SendMessage(hwnd, LVM_SETWORKAREAS, ByVal nWorkAreas, prc(0))
'End Function
'
''#define LVM_GETWORKAREAS        (LVM_FIRST + 70)
''#define ListView_GetWorkAreas(hwnd, nWorkAreas, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETWORKAREAS, (WPARAM)(int)(nWorkAreas), (LPARAM)(RECT *)(prc))
'Public Function ListView_GetWorkAreas(hwnd As Long, nWorkAreas, prc() As RECT) As Boolean
'    ListView_GetWorkAreas = SendMessage(hwnd, LVM_GETWORKAREAS, ByVal nWorkAreas, prc(0))
'End Function
'
''#define LVM_GETNUMBEROFWORKAREAS  (LVM_FIRST + 73)
''#define ListView_GetNumberOfWorkAreas(hwnd, pnWorkAreas) \
''    (BOOL)SNDMSG((hwnd), LVM_GETNUMBEROFWORKAREAS, 0, (LPARAM)(UINT *)(pnWorkAreas))
'Public Function ListView_GetNumberOfWorkAreas(hwnd As Long, pnWorkAreas As Long) As Boolean
'    ListView_GetNumberOfWorkAreas = SendMessage(hwnd, LVM_GETNUMBEROFWORKAREAS, 0, pnWorkAreas)
'End Function
'
''#define LVM_GETSELECTIONMARK    (LVM_FIRST + 66)
''#define ListView_GetSelectionMark(hwnd) \
''    (int)SNDMSG((hwnd), LVM_GETSELECTIONMARK, 0, 0)
'Public Function ListView_GetSelectionMark(hwnd As Long) As Long
'    ListView_GetSelectionMark = SendMessage(hwnd, LVM_GETSELECTIONMARK, 0, 0)
'End Function
'
''#define LVM_SETSELECTIONMARK    (LVM_FIRST + 67)
''#define ListView_SetSelectionMark(hwnd, i) \
''    (int)SNDMSG((hwnd), LVM_SETSELECTIONMARK, 0, (LPARAM)(i))
'Public Function ListView_SetSelectionMark(hwnd As Long, i As Long) As Long
'    ListView_SetSelectionMark = SendMessage(hwnd, LVM_SETSELECTIONMARK, 0, ByVal i)
'End Function
'
''#define LVM_SETHOVERTIME        (LVM_FIRST + 71)
''#define ListView_SetHoverTime(hwndLV, dwHoverTimeMs)\
''        (DWORD)SNDMSG((hwndLV), LVM_SETHOVERTIME, 0, (LPARAM)(dwHoverTimeMs))
'Public Function ListView_SetHoverTime(hwndLV As Long, dwHoverTimeMs As Long) As Long
'    ListView_SetHoverTime = SendMessage(hwndLV, LVM_SETHOVERTIME, 0, ByVal dwHoverTimeMs)
'End Function
'
''#define LVM_GETHOVERTIME        (LVM_FIRST + 72)
''#define ListView_GetHoverTime(hwndLV)\
''        (DWORD)SNDMSG((hwndLV), LVM_GETHOVERTIME, 0, 0)
'Public Function ListView_GetHoverTime(hwndLV As Long) As Long
'    ListView_GetHoverTime = SendMessage(hwndLV, LVM_GETHOVERTIME, 0, 0)
'End Function
'
''#define LVM_GETSTRINGWIDTHA     (LVM_FIRST + 17)
''#define LVM_GETSTRINGWIDTHW     (LVM_FIRST + 87)
''#ifdef UNICODE
''#define  LVM_GETSTRINGWIDTH     LVM_GETSTRINGWIDTHW
''#Else
''#define  LVM_GETSTRINGWIDTH     LVM_GETSTRINGWIDTHA
''#End If
''
''#define ListView_GetStringWidth(hwndLV, psz) \
''    (int)SNDMSG((hwndLV), LVM_GETSTRINGWIDTH, 0, (LPARAM)(LPCTSTR)(psz))
'Public Function ListView_GetStringWidth(hwndLV As Long, psz As String) As Long
'    ListView_GetStringWidth = SendMessage(hwndLV, LVM_GETSTRINGWIDTH, 0, ByVal psz)
'End Function
'
''#define LVM_GETSUBITEMRECT      (LVM_FIRST + 56)
''#define ListView_GetSubItemRect(hwnd, iItem, iSubItem, code, prc) \
''        (BOOL)SNDMSG((hwnd), LVM_GETSUBITEMRECT, (WPARAM)(int)(iItem), \
''                ((prc) ? ((((LPRECT)(prc))->top = (iSubItem)), (((LPRECT)(prc))->left = (code)), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))
'Public Function ListView_GetSubItemRect(hwnd As Long, iItem As Long, iSubItem As Long, Code As LVIR_Flags, prc As RECT) As Boolean
'    prc.Top = iSubItem
'    prc.Left = Code
'    ListView_GetSubItemRect = SendMessage(hwnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
'End Function
'
''#define LVM_GETTEXTBKCOLOR      (LVM_FIRST + 37)
''#define ListView_GetTextBkColor(hwnd)  \
''    (COLORREF)SNDMSG((hwnd), LVM_GETTEXTBKCOLOR, 0, 0L)
'Public Function ListView_GetTextBkColor(hwnd As Long) As Long
'    ListView_GetTextBkColor = SendMessage(hwnd, LVM_GETTEXTBKCOLOR, 0, 0)
'End Function
'
''#define LVM_SETTEXTBKCOLOR      (LVM_FIRST + 38)
''#define ListView_SetTextBkColor(hwnd, clrTextBk) \
''    (BOOL)SNDMSG((hwnd), LVM_SETTEXTBKCOLOR, 0, (LPARAM)(COLORREF)(clrTextBk))
'Public Function ListView_SetTextBkColor(hwnd As Long, ClrTextBk As Long) As Boolean
'    ListView_SetTextBkColor = SendMessage(hwnd, LVM_SETTEXTBKCOLOR, 0, ByVal ClrTextBk)
'End Function
'
''#define LVM_GETTEXTCOLOR        (LVM_FIRST + 35)
''#define ListView_GetTextColor(hwnd)  \
''    (COLORREF)SNDMSG((hwnd), LVM_GETTEXTCOLOR, 0, 0L)
'Public Function ListView_GetTextColor(hwnd As Long) As Long
'    ListView_GetTextColor = SendMessage(hwnd, LVM_GETTEXTCOLOR, 0, 0)
'End Function
'
''#define LVM_SETTEXTCOLOR        (LVM_FIRST + 36)
''#define ListView_SetTextColor(hwnd, clrText) \
''    (BOOL)SNDMSG((hwnd), LVM_SETTEXTCOLOR, 0, (LPARAM)(COLORREF)(clrText))
'Public Function ListView_SetTextColor(hwnd As Long, ClrText As Long) As Boolean
'    ListView_SetTextColor = SendMessage(hwnd, LVM_SETTEXTCOLOR, 0, ByVal ClrText)
'End Function
'
''#define LVM_GETTOPINDEX         (LVM_FIRST + 39)
''#define ListView_GetTopIndex(hwndLV) \
''    (int)SNDMSG((hwndLV), LVM_GETTOPINDEX, 0, 0)
'Public Function ListView_GetTopIndex(hwndLV As Long) As Long
'    ListView_GetTopIndex = SendMessage(hwndLV, LVM_GETTOPINDEX, 0, 0)
'End Function
'
''#define LVM_SUBITEMHITTEST      (LVM_FIRST + 57)
''#define ListView_SubItemHitTest(hwnd, plvhti) \
''        (int)SNDMSG((hwnd), LVM_SUBITEMHITTEST, 0, (LPARAM)(LPLVHITTESTINFO)(plvhti))
'Public Function ListView_SubItemHitTest(hwnd As Long, plvhti As LVHITTESTINFO) As Long
'    ListView_SubItemHitTest = SendMessage(hwnd, LVM_SUBITEMHITTEST, 0, plvhti)
'End Function
'
''#define ListView_SubItemHitTestEx(hwnd, plvhti) \
''        (int)SNDMSG((hwnd), LVM_SUBITEMHITTEST, (WPARAM)-1, (LPARAM)(LPLVHITTESTINFO)(plvhti))
'
''#define LVM_SETTOOLTIPS       (LVM_FIRST + 74)
''#define ListView_SetToolTips(hwndLV, hwndNewHwnd)\
''        (HWND)SNDMSG((hwndLV), LVM_SETTOOLTIPS, (WPARAM)(hwndNewHwnd), 0)
'Public Function ListView_SetToolTips(hwndLV As Long, hwndNewHwnd As Long) As Long
'    ListView_SetToolTips = SendMessage(hwndLV, LVM_SETTOOLTIPS, ByVal hwndNewHwnd, 0)
'End Function
'
''#define LVM_GETTOOLTIPS       (LVM_FIRST + 78)
''#define ListView_GetToolTips(hwndLV)\
''        (HWND)SNDMSG((hwndLV), LVM_GETTOOLTIPS, 0, 0)
'Public Function ListView_GetToolTips(hwndLV As Long) As Long
'    ListView_GetToolTips = SendMessage(hwndLV, LVM_GETTOOLTIPS, 0, 0)
'End Function
'
''#define LVM_GETISEARCHSTRINGA   (LVM_FIRST + 52)
''#define LVM_GETISEARCHSTRINGW   (LVM_FIRST + 117)
''
''#ifdef UNICODE
''#define LVM_GETISEARCHSTRING    LVM_GETISEARCHSTRINGW
''#Else
''#define LVM_GETISEARCHSTRING    LVM_GETISEARCHSTRINGA
''#End If
''
''#define ListView_GetISearchString(hwndLV, lpsz) \
''        (BOOL)SNDMSG((hwndLV), LVM_GETISEARCHSTRING, 0, (LPARAM)(LPTSTR)(lpsz))
'Public Function ListView_GetISearchString(hwndLV As Long, lpsz As String) As Boolean
'    ListView_GetISearchString = SendMessage(hwndLV, LVM_GETISEARCHSTRING, 0, ByVal lpsz)
'End Function
'
''#ifdef UNICODE
''#define LVBKIMAGE               LVBKIMAGEW
''#define LPLVBKIMAGE             LPLVBKIMAGEW
''#define LVM_SETBKIMAGE          LVM_SETBKIMAGEW
''#define LVM_GETBKIMAGE          LVM_GETBKIMAGEW
''#Else
''#define LVBKIMAGE               LVBKIMAGEA
''#define LPLVBKIMAGE             LPLVBKIMAGEA
''#define LVM_SETBKIMAGE          LVM_SETBKIMAGEA
''#define LVM_GETBKIMAGE          LVM_GETBKIMAGEA
''#End If
''
''#define ListView_SetBkImage(hwnd, plvbki) \
''    (BOOL)SNDMSG((hwnd), LVM_SETBKIMAGE, 0, (LPARAM)(plvbki))
'Public Function ListView_SetBkImage(hwnd As Long, plvbki As LVBKIMAGE) As Boolean
'    ListView_SetBkImage = SendMessage(hwnd, LVM_SETBKIMAGE, 0, plvbki)
'End Function
'
''#define ListView_GetBkImage(hwnd, plvbki) \
''    (BOOL)SNDMSG((hwnd), LVM_GETBKIMAGE, 0, (LPARAM)(plvbki))
'Public Function ListView_GetBkImage(hwnd As Long, plvbki As LVBKIMAGE) As Boolean
'    ListView_GetBkImage = SendMessage(hwnd, LVM_GETBKIMAGE, 0, plvbki)
'End Function
'
''#define LVM_SETUNICODEFORMAT     CCM_SETUNICODEFORMAT
''#define ListView_SetUnicodeFormat(hwnd, fUnicode)  \
''    (BOOL)SNDMSG((hwnd), LVM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)
'Public Function ListView_SetUnicodeFormat(hwnd As Long, fUnicode As Boolean) As Boolean
'    ListView_SetUnicodeFormat = SendMessage(hwnd, LVM_SETUNICODEFORMAT, ByVal fUnicode, 0)
'End Function
'
''#define LVM_GETUNICODEFORMAT     CCM_GETUNICODEFORMAT
''#define ListView_GetUnicodeFormat(hwnd)  \
''    (BOOL)SNDMSG((hwnd), LVM_GETUNICODEFORMAT, 0, 0)
'Public Function ListView_GetUnicodeFormat(hwnd As Long) As Boolean
'    ListView_GetUnicodeFormat = SendMessage(hwnd, LVM_GETUNICODEFORMAT, 0, 0)
'End Function
'
''#define LVM_SETEXTENDEDLISTVIEWSTYLE (LVM_FIRST + 54)   // optional wParam == mask
''#define ListView_SetExtendedListViewStyle(hwndLV, dw)\
''        (DWORD)SNDMSG((hwndLV), LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dw)
'
'
''#define ListView_SetExtendedListViewStyleEx(hwndLV, dwMask, dw)\
''        (DWORD)SNDMSG((hwndLV), LVM_SETEXTENDEDLISTVIEWSTYLE, dwMask, dw)
'Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, dwMask As Long, dw As Long) As Long
'    ListView_SetExtendedListViewStyleEx = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, ByVal dwMask, ByVal dw)
'End Function
'
''#define LVM_SETCOLUMNORDERARRAY (LVM_FIRST + 58)
''#define ListView_SetColumnOrderArray(hwnd, iCount, pi) \
''        (BOOL)SNDMSG((hwnd), LVM_SETCOLUMNORDERARRAY, (WPARAM)(iCount), (LPARAM)(LPINT)(pi))
'Public Function ListView_SetColumnOrderArray(hwnd As Long, iCount As Long, lpiArray As Long) As Boolean
'    ListView_SetColumnOrderArray = SendMessage(hwnd, LVM_SETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
'End Function
'
''#define LVM_GETCOLUMNORDERARRAY (LVM_FIRST + 59)
''#define ListView_GetColumnOrderArray(hwnd, iCount, pi) \
''        (BOOL)SNDMSG((hwnd), LVM_GETCOLUMNORDERARRAY, (WPARAM)(iCount), (LPARAM)(LPINT)(pi))
'Public Function ListView_GetColumnOrderArray(hwnd As Long, iCount As Long, lpiArray As Long) As Boolean
'    ListView_GetColumnOrderArray = SendMessage(hwnd, LVM_GETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
'End Function
'
''#define LVM_SETIMAGELIST        (LVM_FIRST + 3)
''#define ListView_SetImageList(hwnd, himl, iImageList) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_SETIMAGELIST, (WPARAM)(iImageList), (LPARAM)(HIMAGELIST)(himl))
'Public Function ListView_SetImageList(hwnd As Long, himl As Long, iImageList As LV_ImageList) As Long
'    ListView_SetImageList = SendMessage(hwnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
'End Function
'
''#define LVM_GETIMAGELIST        (LVM_FIRST + 2)
''#define ListView_GetImageList(hwnd, iImageList) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)(INT)(iImageList), 0L)
'Public Function ListView_GetImageList(hwnd As Long, iImageList As LV_ImageList) As Long
'    ListView_GetImageList = SendMessage(hwnd, LVM_GETIMAGELIST, ByVal iImageList, 0)
'End Function
'
''#define LVM_GETHEADER               (LVM_FIRST + 31)
''#define ListView_GetHeader(hwnd)\
''    (HWND)SNDMSG((hwnd), LVM_GETHEADER, 0, 0L)
'Public Function ListView_GetHeader(hwnd As Long) As Long
'    ListView_GetHeader = SendMessage(hwnd, LVM_GETHEADER, 0, 0)
'End Function
'
''#define LVM_GETITEMA            (LVM_FIRST + 5)
''#define LVM_GETITEMW            (LVM_FIRST + 75)
''#ifdef UNICODE
''#define LVM_GETITEM             LVM_GETITEMW
''#Else
''#define LVM_GETITEM             LVM_GETITEMA
''#End If
''#define ListView_GetItem(hwnd, pitem) \
''    (BOOL)SNDMSG((hwnd), LVM_GETITEM, 0, (LPARAM)(LV_ITEM *)(pitem))
'Public Function ListView_GetItem(hwnd As Long, pItem As LVITEM) As Boolean
'    ListView_GetItem = SendMessage(hwnd, LVM_GETITEM, 0, pItem)
'End Function
'
''#define LVM_SETITEMA            (LVM_FIRST + 6)
''#define LVM_SETITEMW            (LVM_FIRST + 76)
''#ifdef UNICODE
''#define LVM_SETITEM             LVM_SETITEMW
''#Else
''#define LVM_SETITEM             LVM_SETITEMA
''#End If
''#define ListView_SetItem(hwnd, pitem) \
''    (BOOL)SNDMSG((hwnd), LVM_SETITEM, 0, (LPARAM)(const LV_ITEM *)(pitem))
'Public Function ListView_SetItem(hwnd As Long, pItem As LVITEM) As Boolean
'    ListView_SetItem = SendMessage(hwnd, LVM_SETITEM, 0, pItem)
'End Function
''
''#define LVM_SETCALLBACKMASK     (LVM_FIRST + 11)
''#define ListView_SetCallbackMask(hwnd, mask) \
''    (BOOL)SNDMSG((hwnd), LVM_SETCALLBACKMASK, (WPARAM)(UINT)(mask), 0)
'Public Function ListView_SetCallbackMask(hwnd As Long, mask As Long) As Boolean
'    ListView_SetCallbackMask = SendMessage(hwnd, LVM_SETCALLBACKMASK, ByVal mask, 0)
'End Function
'
''#define LVM_GETCALLBACKMASK     (LVM_FIRST + 10)
''#define ListView_GetCallbackMask(hwnd) \
''    (BOOL)SNDMSG((hwnd), LVM_GETCALLBACKMASK, 0, 0)
'Public Function ListView_GetCallbackMask(hwnd As Long) As Long   ' LVStyles
'    ListView_GetCallbackMask = SendMessage(hwnd, LVM_GETCALLBACKMASK, 0, 0)
'End Function
'
''#define LVM_GETCOLUMNA          (LVM_FIRST + 25)
''#define LVM_GETCOLUMNW          (LVM_FIRST + 95)
''#ifdef UNICODE
''#define  LVM_GETCOLUMN          LVM_GETCOLUMNW
''#Else
''#define  LVM_GETCOLUMN          LVM_GETCOLUMNA
''#End If
''
''#define ListView_GetColumn(hwnd, iCol, pcol) \
''    (BOOL)SNDMSG((hwnd), LVM_GETCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(LV_COLUMN *)(pcol))
'Public Function ListView_GetColumn(hwnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
'    ListView_GetColumn = SendMessage(hwnd, LVM_GETCOLUMN, ByVal iCol, pcol)
'End Function
'
''#define LVM_SETCOLUMNA          (LVM_FIRST + 26)
''#define LVM_SETCOLUMNW          (LVM_FIRST + 96)
''#ifdef UNICODE
''#define  LVM_SETCOLUMN          LVM_SETCOLUMNW
''#Else
''#define  LVM_SETCOLUMN          LVM_SETCOLUMNA
''#End If
''
''#define ListView_SetColumn(hwnd, iCol, pcol) \
''    (BOOL)SNDMSG((hwnd), LVM_SETCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(const LV_COLUMN *)(pcol))
'Public Function ListView_SetColumn(hwnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
'    ListView_SetColumn = SendMessage(hwnd, LVM_SETCOLUMN, ByVal iCol, pcol)
'End Function
'
''#define LVM_GETCOUNTPERPAGE     (LVM_FIRST + 40)
''#define ListView_GetCountPerPage(hwndLV) \
''    (int)SNDMSG((hwndLV), LVM_GETCOUNTPERPAGE, 0, 0)
'Public Function ListView_GetCountPerPage(hwndLV As Long) As Long
'    ListView_GetCountPerPage = SendMessage(hwndLV, LVM_GETCOUNTPERPAGE, 0, 0)
'End Function
'
''#define LVM_GETORIGIN           (LVM_FIRST + 41)
''#define ListView_GetOrigin(hwndLV, ppt) \
''    (BOOL)SNDMSG((hwndLV), LVM_GETORIGIN, (WPARAM)0, (LPARAM)(POINT *)(ppt))
'Public Function ListView_GetOrigin(hwndLV As Long, ppt As POINT) As Boolean
'    ListView_GetOrigin = SendMessage(hwndLV, LVM_GETORIGIN, 0, ppt)
'End Function
'
''#define LVM_GETEDITCONTROL      (LVM_FIRST + 24)
''#define ListView_GetEditControl(hwndLV) \
''    (HWND)SNDMSG((hwndLV), LVM_GETEDITCONTROL, 0, 0L)
'Public Function ListView_GetEditControl(hwndLV As Long) As Long
'    ListView_GetEditControl = SendMessage(hwndLV, LVM_GETEDITCONTROL, 0, 0)
'End Function
'
''#define LVM_GETEXTENDEDLISTVIEWSTYLE (LVM_FIRST + 55)
''#define ListView_GetExtendedListViewStyle(hwndLV)\
''        (DWORD)SNDMSG((hwndLV), LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
'Public Function ListView_GetExtendedListViewStyle(hwndLV As Long) As Long
'    ListView_GetExtendedListViewStyle = SendMessage(hwndLV, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
'End Function
'
''#define LVM_SETHOTITEM  (LVM_FIRST + 60)
''#define ListView_SetHotItem(hwnd, i) \
''        (int)SNDMSG((hwnd), LVM_SETHOTITEM, (WPARAM)(i), 0)
'Public Function ListView_SetHotItem(hwnd As Long, i As Long) As Long
'    ListView_SetHotItem = SendMessage(hwnd, LVM_SETHOTITEM, ByVal i, 0)
'End Function
'
''#define LVM_GETHOTITEM  (LVM_FIRST + 61)
''#define ListView_GetHotItem(hwnd) \
''        (int)SNDMSG((hwnd), LVM_GETHOTITEM, 0, 0)
'Public Function ListView_GetHotItem(hwnd As Long) As Long
'    ListView_GetHotItem = SendMessage(hwnd, LVM_GETHOTITEM, 0, 0)
'End Function
'
''#define LVM_SETHOTCURSOR  (LVM_FIRST + 62)
''#define ListView_SetHotCursor(hwnd, hcur) \
''        (HCURSOR)SNDMSG((hwnd), LVM_SETHOTCURSOR, 0, (LPARAM)(hcur))
'Public Function ListView_SetHotCursor(hwnd As Long, hcur As Long) As Long
'    ListView_SetHotCursor = SendMessage(hwnd, LVM_SETHOTCURSOR, 0, ByVal hcur)
'End Function
'
''#define LVM_GETHOTCURSOR  (LVM_FIRST + 63)
''#define ListView_GetHotCursor(hwnd) \
''        (HCURSOR)SNDMSG((hwnd), LVM_GETHOTCURSOR, 0, 0)
'Public Function ListView_GetHotCursor(hwnd As Long) As Long
'    ListView_GetHotCursor = SendMessage(hwnd, LVM_GETHOTCURSOR, 0, 0)
'End Function
'
''#define LVM_SETITEMTEXTA        (LVM_FIRST + 46)
''#define LVM_SETITEMTEXTW        (LVM_FIRST + 116)
''
''#ifdef UNICODE
''#define  LVM_SETITEMTEXT        LVM_SETITEMTEXTW
''#Else
''#define  LVM_SETITEMTEXT        LVM_SETITEMTEXTA
''#End If
''
''#define ListView_SetItemText(hwndLV, i, iSubItem_, pszText_) \
''{ LV_ITEM _macro_lvi;\
''  _macro_lvi.iSubItem = (iSubItem_);\
''  _macro_lvi.pszText = (pszText_);\
''  SNDMSG((hwndLV), LVM_SETITEMTEXT, (WPARAM)(i), (LPARAM)(LV_ITEM *)&_macro_lvi);\
''}
'Public Sub ListView_SetItemText(hwndLV As Long, i As Long, iSubItem As Long, pszText As String)
'    Dim lvi As LVITEM
'    lvi.iSubItem = iSubItem
'    lvi.pszText = pszText
'    lvi.cchTextMax = Len(pszText) + 1
'    SendMessage hwndLV, LVM_SETITEMTEXT, ByVal i, lvi
'End Sub
'
''#define LVM_SETICONSPACING      (LVM_FIRST + 53)
''// -1 for cx and cy means we'll use the default (system settings)
''// 0 for cx or cy means use the current setting (allows you to change just one param)
''#define ListView_SetIconSpacing(hwndLV, cx, cy) \
''        (DWORD)SNDMSG((hwndLV), LVM_SETICONSPACING, 0, MAKELONG(cx,cy))
'Public Function ListView_SetIconSpacing(hwndLV As Long, CX As Long, CY As Long) As Long
'    ListView_SetIconSpacing = SendMessage(hwndLV, LVM_SETICONSPACING, 0, ByVal MAKELONG(CX, CY))
'End Function
'
''// these flags only apply to LVS_OWNERDATA listviews in report or list mode
''#define LVSICF_NOINVALIDATEALL  0x00000001
''#define LVSICF_NOSCROLL         0x00000002
''
''#define LVM_SETITEMCOUNT        (LVM_FIRST + 47)
''#define ListView_SetItemCount(hwndLV, cItems) \
''  SNDMSG((hwndLV), LVM_SETITEMCOUNT, (WPARAM)(cItems), 0)
'Public Sub ListView_SetItemCount(hwndLV As Long, cItems As Long)
'    SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, 0
'End Sub
'
'#If (WIN32_IE >= &H300) Then
'
''#define ListView_SetItemCountEx(hwndLV, cItems, dwFlags) \
''  SNDMSG((hwndLV), LVM_SETITEMCOUNT, (WPARAM)(cItems), (LPARAM)(dwFlags))
'Public Sub ListView_SetItemCountEx(hwndLV As Long, cItems As Long, dwFlags As Long)
'    SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, ByVal dwFlags
'End Sub
''
'#End If
'
'' ListView_GetNextItem
''#define LVM_GETNEXTITEM         (LVM_FIRST + 12)
''#define ListView_GetNextItem(hwnd, i, flags) \
''    (int)SNDMSG((hwnd), LVM_GETNEXTITEM, (WPARAM)(int)(i), MAKELPARAM((flags), 0))
'Public Function ListView_GetNextItem(hwnd As Long, i As Long, Flags As LVNI_Flags) As Long
'    ListView_GetNextItem = SendMessage(hwnd, LVM_GETNEXTITEM, ByVal i, ByVal Flags)    ' ByVal MAKELPARAM(flags, 0))
'End Function
'
'' Returns the index of the item that is selected and has the focus rectangle (user-defined)
'
'Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
'    ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
'End Function
'
'
''#define LV_FINDINFOA    LVFINDINFOA
''#define LV_FINDINFOW    LVFINDINFOW
''
''#define LV_FINDINFO  LVFINDINFO
''
''typedef struct tagLVFINDINFOA
''{
''    UINT flags;
''    LPCSTR psz;
''    LPARAM lParam;
''    POINT pt;
''    UINT vkDirection;
''} LVFINDINFOA, *LPFINDINFOA;
''
''typedef struct tagLVFINDINFOW
''{
''    UINT flags;
''    LPCWSTR psz;
''    LPARAM lParam;
''    POINT pt;
''    UINT vkDirection;
''} LVFINDINFOW, *LPFINDINFOW;
''
''#ifdef UNICODE
''#define  LVFINDINFO            LVFINDINFOW
''#Else
''#define  LVFINDINFO            LVFINDINFOA
''#End If
''
''#define LVM_FINDITEMA           (LVM_FIRST + 13)
''#define LVM_FINDITEMW           (LVM_FIRST + 83)
''#ifdef UNICODE
''#define  LVM_FINDITEM           LVM_FINDITEMW
''#Else
''#define  LVM_FINDITEM           LVM_FINDITEMA
''#End If
''
''#define ListView_FindItem(hwnd, iStart, plvfi) \
''    (int)SNDMSG((hwnd), LVM_FINDITEM, (WPARAM)(int)(iStart), (LPARAM)(const LV_FINDINFO *)(plvfi))
'Public Function ListView_FindItem(hwnd As Long, iStart, plvfi As LVFINDINFO) As Long
'    ListView_FindItem = SendMessage(hwnd, LVM_FINDITEM, ByVal iStart, plvfi)
'End Function
'
''#define LVM_GETITEMRECT         (LVM_FIRST + 14)
''#define ListView_GetItemRect(hwnd, i, prc, code) \
''     (BOOL)SNDMSG((hwnd), LVM_GETITEMRECT, (WPARAM)(int)(i), \
''           ((prc) ? (((RECT *)(prc))->left = (code),(LPARAM)(RECT *)(prc)) : (LPARAM)(RECT *)NULL))
'Public Function ListView_GetItemRect(hwnd As Long, i As Long, prc As RECT, Code As LVIR_Flags) As Boolean
'    prc.Left = Code
'    ListView_GetItemRect = SendMessage(hwnd, LVM_GETITEMRECT, ByVal i, prc)
'End Function
'
''#define ListView_GetCheckState(hwndLV, i) \
''   ((((UINT)(SNDMSG((hwndLV), LVM_GETITEMSTATE, (WPARAM)(i), LVIS_STATEIMAGEMASK))) >> 12) -1)
'Public Function ListView_GetCheckState(hwndLV As Long, iIndex As Long) As Long   ' updated
'    Dim dwState As Long
'    dwState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal iIndex, ByVal LVIS_STATEIMAGEMASK)
'    ListView_GetCheckState = (dwState \ 2 ^ 12) - 1
'  '((((UINT)(SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, LVIS_STATEIMAGEMASK))) >> 12) -1)
'End Function
'
''#define ListView_SetCheckState(hwndLV, i, fCheck) \
''  ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK((fCheck)?2:1), LVIS_STATEIMAGEMASK)
'Public Function ListView_SetCheckState(hwndLV As Long, i As Long, fCheck As Long) As Long
''#define ListView_SetCheckState(hwndLV, i, fCheck) \
''  ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK((fCheck)?2:1), LVIS_STATEIMAGEMASK)
'    ListView_SetCheckState = ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK(IIf(fCheck, 2, 1)), LVIS_STATEIMAGEMASK)
'End Function
'
''#define LVM_GETITEMCOUNT        (LVM_FIRST + 4)
''#define ListView_GetItemCount(hwnd) \
''    (int)SNDMSG((hwnd), LVM_GETITEMCOUNT, 0, 0L)
'Public Function ListView_GetItemCount(hwnd As Long) As Long
'    ListView_GetItemCount = SendMessage(hwnd, LVM_GETITEMCOUNT, 0, 0)
'End Function
'
''#define LVM_GETITEMPOSITION     (LVM_FIRST + 16)
''#define ListView_GetItemPosition(hwndLV, i, ppt) \
''    (BOOL)SNDMSG((hwndLV), LVM_GETITEMPOSITION, (WPARAM)(int)(i), (LPARAM)(POINT *)(ppt))
'Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINT) As Boolean
'    ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
'End Function
'
''#define LVM_SETITEMPOSITION     (LVM_FIRST + 15)
''#define ListView_SetItemPosition(hwndLV, i, x, y) \
''    (BOOL)SNDMSG((hwndLV), LVM_SETITEMPOSITION, (WPARAM)(int)(i), MAKELPARAM((x), (y)))
'Public Function ListView_SetItemPosition(hwndLV As Long, i As Long, X As Long, Y As Long) As Boolean
'    ListView_SetItemPosition = SendMessage(hwndLV, LVM_SETITEMPOSITION, ByVal i, ByVal MAKELPARAM(X, Y))
'End Function
'
''#define LVM_SETITEMPOSITION32   (LVM_FIRST + 49)
''#define ListView_SetItemPosition32(hwndLV, i, x0, y0) \
''{   POINT ptNewPos; \
''    ptNewPos.x = (x0); ptNewPos.y = (y0); \
''    SNDMSG((hwndLV), LVM_SETITEMPOSITION32, (WPARAM)(int)(i), (LPARAM)&ptNewPos); \
''}
'Public Sub ListView_SetItemPosition32(hwndLV As Long, i As Long, X As Long, Y As Long)
'    Dim ptNewPos As POINT
'    ptNewPos.X = X
'    ptNewPos.Y = Y
'    SendMessage hwndLV, LVM_SETITEMPOSITION32, ByVal i, ptNewPos
'End Sub
'
'Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
'    ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)
'End Function
'
''#define LVM_UPDATE              (LVM_FIRST + 42)
''#define ListView_Update(hwndLV, i) \
''    (BOOL)SNDMSG((hwndLV), LVM_UPDATE, (WPARAM)(i), 0L)
'Public Function ListView_Update(hwndLV As Long, i As Long) As Boolean
'    ListView_Update = SendMessage(hwndLV, LVM_UPDATE, ByVal i, 0)
'End Function
'
''#define LVM_GETITEMSPACING      (LVM_FIRST + 51)
''#define ListView_GetItemSpacing(hwndLV, fSmall) \
''        (DWORD)SNDMSG((hwndLV), LVM_GETITEMSPACING, fSmall, 0L)
'Public Function ListView_GetItemSpacing(hwndLV As Long, fSmall As Boolean) As Long
'    ListView_GetItemSpacing = SendMessage(hwndLV, LVM_GETITEMSPACING, ByVal fSmall, 0)
'End Function
'
''#define LVM_GETITEMSTATE        (LVM_FIRST + 44)
''#define ListView_GetItemState(hwndLV, i, mask) \
''   (UINT)SNDMSG((hwndLV), LVM_GETITEMSTATE, (WPARAM)(i), (LPARAM)(mask))
'Public Function ListView_GetItemState(hwndLV As Long, i As Long, mask As LVITEM_state) As Long   ' LVITEM_state
'    ListView_GetItemState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, ByVal mask)
'End Function
'
''#define LVM_GETITEMTEXTA        (LVM_FIRST + 45)
''#define LVM_GETITEMTEXTW        (LVM_FIRST + 115)
''
''#ifdef UNICODE
''#define  LVM_GETITEMTEXT        LVM_GETITEMTEXTW
''#Else
''#define  LVM_GETITEMTEXT        LVM_GETITEMTEXTA
''#End If
''
''#define ListView_GetItemText(hwndLV, i, iSubItem_, pszText_, cchTextMax_) \
''{ LV_ITEM _macro_lvi;\
''  _macro_lvi.iSubItem = (iSubItem_);\
''  _macro_lvi.cchTextMax = (cchTextMax_);\
''  _macro_lvi.pszText = (pszText_);\
''  SNDMSG((hwndLV), LVM_GETITEMTEXT, (WPARAM)(i), (LPARAM)(LV_ITEM *)&_macro_lvi);\
''}
'Public Sub ListView_GetItemText(hwndLV As Long, i As Long, iSubItem As Long, _
'                                                     pszText As Long, cchTextMax As Long)
'    Dim lvi As LVITEM
'    lvi.iSubItem = iSubItem
'    lvi.cchTextMax = cchTextMax
'    lvi.pszText = pszText
'    SendMessage hwndLV, LVM_GETITEMTEXT, ByVal i, lvi
'    pszText = lvi.pszText   ' fills pszText w/ pointer
'End Sub
'
'
''#define LVM_HITTEST             (LVM_FIRST + 18)
''#define ListView_HitTest(hwndLV, pinfo) \
''    (int)SNDMSG((hwndLV), LVM_HITTEST, 0, (LPARAM)(LV_HITTESTINFO *)(pinfo))
'Public Function ListView_HitTest(hwndLV As Long, pInfo As LVHITTESTINFO) As Long
'    ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pInfo)
'End Function
'
''#define LVM_INSERTITEMA         (LVM_FIRST + 7)
''#define LVM_INSERTITEMW         (LVM_FIRST + 77)
''#ifdef UNICODE
''#define LVM_INSERTITEM          LVM_INSERTITEMW
''#Else
''#define LVM_INSERTITEM          LVM_INSERTITEMA
''#End If
''#define ListView_InsertItem(hwnd, pitem)   \
''    (int)SNDMSG((hwnd), LVM_INSERTITEM, 0, (LPARAM)(const LV_ITEM *)(pitem))
'Public Function ListView_InsertItem(hwnd As Long, pItem As LVITEM) As Long
'    ListView_InsertItem = SendMessage(hwnd, LVM_INSERTITEM, 0, pItem)
'End Function
'
''#define LVM_DELETECOLUMN        (LVM_FIRST + 28)
''#define ListView_DeleteColumn(hwnd, iCol) \
''    (BOOL)SNDMSG((hwnd), LVM_DELETECOLUMN, (WPARAM)(int)(iCol), 0)
'Public Function ListView_DeleteColumn(hwnd As Long, iCol As Long) As Boolean
'    ListView_DeleteColumn = SendMessage(hwnd, LVM_DELETECOLUMN, ByVal iCol, 0)
'End Function
'
''#define LVM_ENSUREVISIBLE       (LVM_FIRST + 19)
''#define ListView_EnsureVisible(hwndLV, i, fPartialOK) \
''    (BOOL)SNDMSG((hwndLV), LVM_ENSUREVISIBLE, (WPARAM)(int)(i), MAKELPARAM((fPartialOK), 0))
'Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As CBoolean) As Boolean
'    ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal fPartialOK)   ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
'End Function
'
''#define LVM_INSERTCOLUMNA       (LVM_FIRST + 27)
''#define LVM_INSERTCOLUMNW       (LVM_FIRST + 97)
''#ifdef UNICODE
''#   define  LVM_INSERTCOLUMN    LVM_INSERTCOLUMNW
''#Else
''#   define  LVM_INSERTCOLUMN    LVM_INSERTCOLUMNA
''#End If
''
''#define ListView_InsertColumn(hwnd, iCol, pcol) \
''    (int)SNDMSG((hwnd), LVM_INSERTCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(const LV_COLUMN *)(pcol))
'Public Function ListView_InsertColumn(hwnd As Long, iCol As Long, pcol As LVCOLUMN) As Long
'    ListView_InsertColumn = SendMessage(hwnd, LVM_INSERTCOLUMN, ByVal iCol, pcol)
'End Function
'
''#define LVM_SCROLL              (LVM_FIRST + 20)
''#define ListView_Scroll(hwndLV, dx, dy) \
''    (BOOL)SNDMSG((hwndLV), LVM_SCROLL, (WPARAM)(int)(dx), (LPARAM)(int)(dy))
'Public Function ListView_Scroll(hwndLV As Long, DX As Long, DY As Long) As Boolean
'    ListView_Scroll = SendMessage(hwndLV, LVM_SCROLL, ByVal DX, ByVal DY)
'End Function
'
''#define LVM_DELETEALLITEMS      (LVM_FIRST + 9)
''#define ListView_DeleteAllItems(hwnd) \
''    (BOOL)SNDMSG((hwnd), LVM_DELETEALLITEMS, 0, 0L)
'Public Function ListView_DeleteAllItems(hwnd As Long) As Boolean
'    ListView_DeleteAllItems = SendMessage(hwnd, LVM_DELETEALLITEMS, 0, 0)
'End Function
'
''#define LVM_GETCOLUMNWIDTH      (LVM_FIRST + 29)
''#define ListView_GetColumnWidth(hwnd, iCol) \
''    (int)SNDMSG((hwnd), LVM_GETCOLUMNWIDTH, (WPARAM)(int)(iCol), 0)
'Public Function ListView_GetColumnWidth(hwnd As Long, iCol As Long) As Long
'    ListView_GetColumnWidth = SendMessage(hwnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0)
'End Function
'
''#define LVSCW_AUTOSIZE              -1
''#define LVSCW_AUTOSIZE_USEHEADER    -2
''#define LVM_SETCOLUMNWIDTH          (LVM_FIRST + 30)
''
''#define ListView_SetColumnWidth(hwnd, iCol, cx) \
''    (BOOL)SNDMSG((hwnd), LVM_SETCOLUMNWIDTH, (WPARAM)(int)(iCol), MAKELPARAM((cx), 0))
'Public Function ListView_SetColumnWidth(hwnd As Long, iCol As Long, CX As Long) As Boolean
'    ListView_SetColumnWidth = SendMessage(hwnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(CX, 0))
'End Function
'
''#define LVM_REDRAWITEMS         (LVM_FIRST + 21)
''#define ListView_RedrawItems(hwndLV, iFirst, iLast) \
''    (BOOL)SNDMSG((hwndLV), LVM_REDRAWITEMS, (WPARAM)(int)(iFirst), (LPARAM)(int)(iLast))
'Public Function ListView_RedrawItems(hwndLV As Long, iFirst, iLast) As Boolean
'    ListView_RedrawItems = SendMessage(hwndLV, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast)
'End Function
'
''#define LVM_GETSELECTEDCOUNT    (LVM_FIRST + 50)
''#define ListView_GetSelectedCount(hwndLV) \
''    (UINT)SNDMSG((hwndLV), LVM_GETSELECTEDCOUNT, 0, 0L)
'Public Function ListView_GetSelectedCount(hwndLV As Long) As Long
'    ListView_GetSelectedCount = SendMessage(hwndLV, LVM_GETSELECTEDCOUNT, 0, 0)
'End Function
'
''#define LVM_SETVIEW         (LVM_FIRST + 142)
''#define ListView_SetView(hwnd, iView) \
''    (DWORD)SNDMSG((hwnd), LVM_SETVIEW, (WPARAM)(DWORD)(iView), 0)
'
''#define LVM_GETVIEW         (LVM_FIRST + 143)
''#define ListView_GetView(hwnd) \
''    (DWORD)SNDMSG((hwnd), LVM_GETVIEW, 0, 0)
'Public Function ListView_GetView(hwnd As Long) As Long
'    ListView_GetView = SendMessage(hwnd, LVM_GETVIEW, 0, ByVal 0&)
'End Function
'
''#define LVM_GETVIEWRECT         (LVM_FIRST + 34)
''#define ListView_GetViewRect(hwnd, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETVIEWRECT, 0, (LPARAM)(RECT *)(prc))
'Public Function ListView_GetViewRect(hwnd As Long, prc As RECT) As Boolean
'    ListView_GetViewRect = SendMessage(hwnd, LVM_GETVIEWRECT, 0, prc)
'End Function
'' ListView_SetItemState
'
''#define LVM_SETITEMSTATE        (LVM_FIRST + 43)
''#define ListView_SetItemState(hwndLV, i, data, mask) \
''{ LV_ITEM _macro_lvi;\
''  _macro_lvi.stateMask = (mask);\
''  _macro_lvi.state = (data);\
''  SNDMSG((hwndLV), LVM_SETITEMSTATE, (WPARAM)(i), (LPARAM)(LV_ITEM *)&_macro_lvi);\
''}
'Public Function ListView_SetItemState(hwndLV As Long, i As Long, State As Long, mask As Long) As Boolean
'    Dim lvi As LVITEM
'    lvi.State = State
'    lvi.stateMask = mask
'    ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
'End Function
'
'' Selects all listview items. The item with the focus rectangle maintains it (user-defined).
'
'Public Function ListView_SelectAll(hwndLV As Long) As Boolean
'    ListView_SelectAll = ListView_SetItemState(hwndLV, -1, LVIS_SELECTED, LVIS_SELECTED)
'End Function
'
'Public Function ListView_SelectNone(hwndLV As Long) As Boolean
'    Dim lv As LVITEM
'
'    With lv
'        .mask = LVIF_STATE
'        .State = False
'        .stateMask = LVIS_SELECTED
'    End With
'
'    ListView_SelectNone = SendMessage(hwndLV, LVM_SETITEMSTATE, -1, lv)
'End Function
'
'' Selects the specified item and gives it the focus rectangle.
'' does not de-select any currently selected items (user-defined).
'
'Public Function ListView_SetFocusedItem(hwndLV As Long, i As Long) As Boolean
'    ListView_SetFocusedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)
'End Function
'
''#define LVM_SORTITEMS           (LVM_FIRST + 48)
''#define ListView_SortItems(hwndLV, _pfnCompare, _lPrm) \
''  (BOOL)SNDMSG((hwndLV), LVM_SORTITEMS, (WPARAM)(LPARAM)(_lPrm), \
''  (LPARAM)(PFNLVCOMPARE)(_pfnCompare))
'Public Function ListView_SortItems(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
'    ListView_SortItems = SendMessage(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
'End Function
'
''#define LVM_SORTITEMSEX          (LVM_FIRST + 81)
''#define ListView_SortItemsEx(hwndLV, _pfnCompare, _lPrm) \
''  (BOOL)SNDMSG((hwndLV), LVM_SORTITEMSEX, (WPARAM)(LPARAM)(_lPrm), (LPARAM)(PFNLVCOMPARE)(_pfnCompare))
'Public Function ListView_SortItemsEx(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
'    ListView_SortItemsEx = SendMessage(hwndLV, LVM_SORTITEMSEX, ByVal lParamSort, ByVal pfnCompare)
'End Function
'
'Public Function ListView_SetExtendedStyle(hwnd As Long, lST As LVStylesEx) As Long
'    Dim lStyle As Long
'
'    lStyle = SendMessage(hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
'    lStyle = lStyle Or lST
'    Call SendMessage(hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)
'End Function
'
'Public Function ListView_GetStyle(hwnd As Long) As LVStyles
'    ListView_GetStyle = GetWindowLong(hwnd, GWL_STYLE)
'End Function
'
'Public Function ListView_SetStyle(hwnd As Long, dwStyle As LVStyles) As Long
'    ListView_SetStyle = SetWindowLong(hwnd, GWL_STYLE, dwStyle)
'End Function
'
''THE MACROS BELOW ARE ONLY FOR VISTA AND HIGHER
''#define LVM_CANCELEDITLABEL     (LVM_FIRST + 179)
''#define ListView_CancelEditLabel(hwnd) \
''    (VOID)SNDMSG((hwnd), LVM_CANCELEDITLABEL, (WPARAM)0, (LPARAM)0)
'Public Function ListView_CancelEditLabel(hwnd As Long) As Long
'    ListView_CancelEditLabel = SendMessage(hwnd, LVM_CANCELEDITLABEL, 0, ByVal 0&)
'End Function
'
''#define LVM_ENABLEGROUPVIEW         (LVM_FIRST + 157)
''#define ListView_EnableGroupView(hwnd, fEnable) \
''    SNDMSG((hwnd), LVM_ENABLEGROUPVIEW, (WPARAM)(fEnable), 0)
'Public Function ListView_EnableGroupView(hwnd As Long, fEnable As Long) As Long
'    ListView_EnableGroupView = SendMessage(hwnd, LVM_ENABLEGROUPVIEW, fEnable, ByVal 0&)
'End Function
'
''#define LVM_GETEMPTYTEXT (LVM_FIRST + 204)
''#define ListView_GetEmptyText(hwnd, pszText, cchText) \
''    (BOOL)SNDMSG((hwnd), LVM_GETEMPTYTEXT, (WPARAM)(cchText), (LPARAM)(pszText))
'Public Function ListView_GetEmptyText(hwnd As Long, cchText As Long, pszText As String) As Long
'    ListView_GetEmptyText = SendMessage(hwnd, LVM_GETEMPTYTEXT, cchText, ByVal pszText)
'End Function
'
''#define LVM_GETFOCUSEDGROUP             (LVM_FIRST + 93)
''#define ListView_GetFocusedGroup(hwnd) \
''    SNDMSG((hwnd), LVM_GETFOCUSEDGROUP, 0, 0)
'Public Function ListView_GetFocusedGroup(hwnd As Long) As Long
''#define ListView_GetFocusedGroup(hwnd) \
''    SNDMSG((hwnd), LVM_GETFOCUSEDGROUP, 0, 0)
'    ListView_GetFocusedGroup = SendMessage(hwnd, LVM_GETFOCUSEDGROUP, 0, ByVal 0&)
'End Function
'
''// footer flags
''#define LVFF_ITEMCOUNT          0x00000001
''
''typedef struct tagLVFOOTERINFO
''{
''    UINT mask;          // LVFF_*
''    LPWSTR pszText;
''    int cchTextMax;
''    UINT cItems;
''} LVFOOTERINFO, *LPLVFOOTERINFO;
''
''#define LVM_GETFOOTERINFO (LVM_FIRST + 206)
''#define ListView_GetFooterInfo(hwnd, plvfi) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERINFO, (WPARAM)(0), (LPARAM)(plvfi))
'Public Function ListView_GetFooterInfo(hwnd As Long, plvfi As Long) As Long
''#define ListView_GetFooterInfo(hwnd, plvfi) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERINFO, (WPARAM)(0), (LPARAM)(plvfi))
'    ListView_GetFooterInfo = SendMessage(hwnd, LVM_GETFOOTERINFO, 0, ByVal plvfi)
'End Function
'
''// footer item flags
''#define LVFIF_TEXT               0x00000001
''#define LVFIF_STATE              0x00000002
''
''// footer item state
''#define LVFIS_FOCUSED            0x0001
''
''typedef struct tagLVFOOTERITEM
''{
''    UINT mask;          // LVFIF_*
''    int iItem;
''    LPWSTR pszText;
''    int cchTextMax;
''    UINT state;         // LVFIS_*
''    UINT stateMask;     // LVFIS_*
''} LVFOOTERITEM, *LPLVFOOTERITEM;
''
''#define LVM_GETFOOTERITEM (LVM_FIRST + 208)
''#define ListView_GetFooterItem(hwnd, iItem, pfi) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEM, (WPARAM)(iItem), (LPARAM)(pfi))
'Public Function ListView_GetFooterItem(hwnd As Long, iItem As Long, pfi As LVFOOTERITEM) As Long
''#define ListView_GetFooterItem(hwnd, iItem, pfi) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEM, (WPARAM)(iItem), (LPARAM)(pfi))
'    ListView_GetFooterItem = SendMessage(hwnd, LVM_GETFOOTERITEM, iItem, pfi)
'End Function
'
''#define LVM_GETFOOTERITEMRECT (LVM_FIRST + 207)
''#define ListView_GetFooterItemRect(hwnd, iItem, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEMRECT, (WPARAM)(iItem), (LPARAM)(prc))
'Public Function ListView_GetFooterItemRect(hwnd As Long, iItem As Long, prc As RECT) As Long
''#define ListView_GetFooterItemRect(hwnd, iItem, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEMRECT, (WPARAM)(iItem), (LPARAM)(prc))
'    ListView_GetFooterItemRect = SendMessage(hwnd, LVM_GETFOOTERITEMRECT, iItem, prc)
'End Function
'
''#define LVM_GETFOOTERRECT (LVM_FIRST + 205)
''#define ListView_GetFooterRect(hwnd, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERRECT, (WPARAM)(0), (LPARAM)(prc))
'Public Function ListView_GetFooterRect(hwnd As Long, prc As RECT) As Long
''#define ListView_GetFooterRect(hwnd, prc) \
''    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERRECT, (WPARAM)(0), (LPARAM)(prc))
'    ListView_GetFooterRect = SendMessage(hwnd, LVM_GETFOOTERRECT, 0, prc)
'End Function
'
''#define ListView_GetGroupHeaderImageList(hwnd) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, 0L)
'Public Function ListView_GetGroupHeaderImageList(hwnd As Long) As Long
''#define ListView_GetGroupHeaderImageList(hwnd) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, 0L)
'    ListView_GetGroupHeaderImageList = SendMessage(hwnd, LVM_GETIMAGELIST, LVSIL_GROUPHEADER, ByVal 0&)
'End Function
'
''#define ListView_SetGroupHeaderImageList(hwnd, himl) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_SETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, (LPARAM)(HIMAGELIST)(himl))
'Public Function ListView_SetGroupHeaderImageList(hwnd As Long, himl As Long) As Long
''#define ListView_GetGroupHeaderImageList(hwnd) \
''    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, 0L)
'    ListView_SetGroupHeaderImageList = SendMessage(hwnd, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal himl)
'End Function
'
''#define LVM_GETGROUPINFO         (LVM_FIRST + 149)
''#define ListView_GetGroupInfo(hwnd, iGroupId, pgrp) \
''    SNDMSG((hwnd), LVM_GETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
'Public Function ListView_GetGroupInfo(hwnd As Long, iGroupId As Long, pgrp As lvGroup) As Long
''#define ListView_GetGroupInfo(hwnd, iGroupId, pgrp) \
''    SNDMSG((hwnd), LVM_GETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
'    ListView_GetGroupInfo = SendMessage(hwnd, LVM_GETGROUPINFO, iGroupId, pgrp)
'End Function
'
''#define LVM_SETGROUPINFO         (LVM_FIRST + 147)
''#define ListView_SetGroupInfo(hwnd, iGroupId, pgrp) \
''    SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
'Public Function ListView_SetGroupInfo(hwnd As Long, iGroupId As Long, pgrp As lvGroup) As Long
''#define ListView_SetGroupInfo(hwnd, iGroupId, pgrp) \
''    SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
'    ListView_SetGroupInfo = SendMessage(hwnd, LVM_SETGROUPINFO, iGroupId, pgrp)
'End Function
'
''#define LVM_GETGROUPINFOBYINDEX         (LVM_FIRST + 153)
''#define ListView_GetGroupInfoByIndex(hwnd, iIndex, pgrp) \
''    SNDMSG((hwnd), LVM_GETGROUPINFOBYINDEX, (WPARAM)(iIndex), (LPARAM)(pgrp))
'Public Function ListView_GetGroupInfoByIndex(hwnd As Long, iIndex As Long, pgrp As lvGroup) As Long
''#define ListView_GetGroupInfoByIndex(hwnd, iIndex, pgrp) \
''    SNDMSG((hwnd), LVM_GETGROUPINFOBYINDEX, (WPARAM)(iIndex), (LPARAM)(pgrp))
'    ListView_GetGroupInfoByIndex = SendMessage(hwnd, LVM_GETGROUPINFOBYINDEX, iIndex, pgrp)
'End Function
'
''#define LVM_SETGROUPMETRICS         (LVM_FIRST + 155)
''#define ListView_SetGroupMetrics(hwnd, pGroupMetrics) \
''    SNDMSG((hwnd), LVM_SETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
'Public Function ListView_SetGroupMetrics(hwnd As Long, pGroupMetrics As LVGROUPMETRICS) As Long
''#define ListView_SetGroupMetrics(hwnd, pGroupMetrics) \
''    SNDMSG((hwnd), LVM_SETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
'    ListView_SetGroupMetrics = SendMessage(hwnd, LVM_SETGROUPMETRICS, 0, pGroupMetrics)
'End Function
'
''#define LVM_GETGROUPMETRICS         (LVM_FIRST + 156)
''#define ListView_GetGroupMetrics(hwnd, pGroupMetrics) \
''    SNDMSG((hwnd), LVM_GETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
'Public Function ListView_GetGroupMetrics(hwnd As Long, pGroupMetrics As LVGROUPMETRICS) As Long
''#define ListView_GetGroupMetrics(hwnd, pGroupMetrics) \
''    SNDMSG((hwnd), LVM_GETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
'    ListView_GetGroupMetrics = SendMessage(hwnd, LVM_GETGROUPMETRICS, 0, pGroupMetrics)
'End Function
'
''#define LVM_GETGROUPRECT               (LVM_FIRST + 98)
''#define ListView_GetGroupRect(hwnd, iGroupId, type, prc) \
''    SNDMSG((hwnd), LVM_GETGROUPRECT, (WPARAM)(iGroupId), \
''        ((prc) ? (((RECT*)(prc))->top = (type)), (LPARAM)(RECT*)(prc) : (LPARAM)(RECT*)NULL))
'Public Function ListView_GetGroupRect(hwnd As Long, iGroup As Long, Item As LVGROUPRECT, rc As RECT) As Long
'    rc.Top = Item
'    ListView_GetGroupRect = SendMessage(hwnd, LVM_GETGROUPRECT, iGroup, rc)
'End Function
'
''#define LVM_GETGROUPCOUNT         (LVM_FIRST + 152)
''#define ListView_GetGroupCount(hwnd) \
''    SNDMSG((hwnd), LVM_GETGROUPCOUNT, (WPARAM)0, (LPARAM)0)
'Public Function ListView_GetGroupCount(hwnd As Long, iGroup As Long) As Long
'    Dim LVG As lvGroup
'    LVG.mask = LVGF_ITEMS
'    LVG.cbSize = LenB(LVG)
'    Call SendMessage(hwnd, LVM_GETGROUPINFO, iGroup, LVG)
'    ListView_GetGroupCount = LVG.cItems
'End Function
'
''#define LVM_GETGROUPSTATE               (LVM_FIRST + 92)
''#define ListView_GetGroupState(hwnd, dwGroupId, dwMask) \
''  (UINT) SNDMSG((hwnd), LVM_GETGROUPSTATE, (WPARAM)(dwGroupId), (LPARAM)(dwMask))
'Public Function ListView_GetGroupState(hwnd As Long, dwGroupId As Long, dwMask As Long) As Long
'    ListView_GetGroupState = SendMessage(hwnd, LVM_GETGROUPSTATE, dwGroupId, ByVal dwMask)
'End Function
'
''#define ListView_SetGroupState(hwnd, dwGroupId, dwMask, dwState) \
''{ LVGROUP _macro_lvg;\
''  _macro_lvg.cbSize = sizeof(_macro_lvg);\
''  _macro_lvg.mask = LVGF_STATE;\
''  _macro_lvg.stateMask = dwMask;\
''  _macro_lvg.state = dwState;\
''  SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)(dwGroupId), (LPARAM)(LVGROUP *)&_macro_lvg);\
'Public Function ListView_SetGroupState(hwnd As Long, dwGroupId As Long, dwMask As Long, dwState As Long) As Long
'    Dim LVG As lvGroup
'    LVG.cbSize = LenB(LVG)
'    LVG.mask = LVGF_STATE
'    LVG.stateMask = dwMask
'    LVG.State = dwState
'    ListView_SetGroupState = SendMessage(hwnd, LVM_SETGROUPINFO, dwGroupId, LVG)
'End Function
'
''#define LVM_GETINSERTMARK                   (LVM_FIRST + 167)
''#define ListView_GetInsertMark(hwnd, lvim) \
''    (BOOL)SNDMSG((hwnd), LVM_GETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
'Public Function ListView_GetInsertMark(hwnd As Long, LVIM As LVINSERTMARK) As Long
''#define ListView_GetInsertMark(hwnd, lvim) \
''    (BOOL)SNDMSG((hwnd), LVM_GETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
'    ListView_GetInsertMark = SendMessage(hwnd, LVM_GETINSERTMARK, 0, LVIM)
'End Function
'
''typedef struct
''{
''    UINT cbSize;
''    DWORD dwFlags;
''    int iItem;
''    DWORD dwReserved;
''} LVINSERTMARK, * LPLVINSERTMARK;
''
''#define LVIM_AFTER      0x00000001 // TRUE = insert After iItem, otherwise before
''
''#define LVM_SETINSERTMARK                   (LVM_FIRST + 166)
''#define ListView_SetInsertMark(hwnd, lvim) \
''    (BOOL)SNDMSG((hwnd), LVM_SETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
'Public Function ListView_SetInsertMark(hwnd As Long, LVIM As LVINSERTMARK) As Long
''#define ListView_SetInsertMark(hwnd, lvim) \
''    (BOOL)SNDMSG((hwnd), LVM_SETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
'    ListView_SetInsertMark = SendMessage(hwnd, LVM_SETINSERTMARK, 0, LVIM)
'End Function
'
''#define LVM_GETINSERTMARKCOLOR                 (LVM_FIRST + 171)
''#define ListView_GetInsertMarkColor(hwnd) \
''    (COLORREF)SNDMSG((hwnd), LVM_GETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)0)
'Public Function ListView_GetInsertMarkColor(hwnd As Long) As Long
''#define ListView_GetInsertMarkColor(hwnd) \
''    (COLORREF)SNDMSG((hwnd), LVM_GETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)0)
'    ListView_GetInsertMarkColor = SendMessage(hwnd, LVM_GETINSERTMARKCOLOR, 0, ByVal 0&)
'End Function
'
''#define LVM_SETINSERTMARKCOLOR                 (LVM_FIRST + 170)
''#define ListView_SetInsertMarkColor(hwnd, color) \
''    (COLORREF)SNDMSG((hwnd), LVM_SETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
'Public Function ListView_SetInsertMarkColor(hwnd As Long, Color As Long) As Long
''#define ListView_SetInsertMarkColor(hwnd, color) \
''    (COLORREF)SNDMSG((hwnd), LVM_SETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
'    ListView_SetInsertMarkColor = SendMessage(hwnd, LVM_SETINSERTMARKCOLOR, 0, ByVal Color)
'End Function
'
''#define LVM_GETINSERTMARKRECT               (LVM_FIRST + 169)
''#define ListView_GetInsertMarkRect(hwnd, rc) \
''    (int)SNDMSG((hwnd), LVM_GETINSERTMARKRECT, (WPARAM)0, (LPARAM)(LPRECT)(rc))
'Public Function ListView_GetInsertMarkRect(hwnd As Long, rc As RECT) As Long
''#define ListView_GetInsertMarkRect(hwnd, rc) \
''    (int)SNDMSG((hwnd), LVM_GETINSERTMARKRECT, (WPARAM)0, (LPARAM)(LPRECT)(rc))
'    ListView_GetInsertMarkRect = SendMessage(hwnd, LVM_GETINSERTMARKRECT, 0, rc)
'End Function
'
''#define LVM_INSERTMARKHITTEST               (LVM_FIRST + 168)
''#define ListView_InsertMarkHitTest(hwnd, point, lvim) \
''    (int)SNDMSG((hwnd), LVM_INSERTMARKHITTEST, (WPARAM)(LPPOINT)(point), (LPARAM)(LPLVINSERTMARK)(lvim))
'Public Function ListView_InsertMarkHitTest(hwnd As Long, POINT As POINT, LVIM As LVINSERTMARK) As Long
''#define ListView_InsertMarkHitTest(hwnd, point, lvim) \
''    (int)SNDMSG((hwnd), LVM_INSERTMARKHITTEST, (WPARAM)(LPPOINT)(point), (LPARAM)(LPLVINSERTMARK)(lvim))
'    ListView_InsertMarkHitTest = SendMessage(hwnd, LVM_INSERTMARKHITTEST, VarPtr(POINT), LVIM)
'End Function
'
''// supports a single item in multiple groups.
''typedef struct tagLVITEMINDEX
''{
''    int iItem;          // listview item index
''    int iGroup;         // group index (must be -1 if group view is not enabled)
''} LVITEMINDEX, *PLVITEMINDEX;
''
''#define LVM_GETITEMINDEXRECT    (LVM_FIRST + 209)
''#define ListView_GetItemIndexRect(hwnd, plvii, iSubItem, code, prc) \
''        (BOOL)SNDMSG((hwnd), LVM_GETITEMINDEXRECT, (WPARAM)(LVITEMINDEX*)(plvii), \
''                ((prc) ? ((((LPRECT)(prc))->top = (iSubItem)), (((LPRECT)(prc))->left = (code)), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))
'Public Function ListView_GetItemIndexRect(hwnd As Long, lvii As LVITEMINDEX, iSubItem As Long, Code As Long, prc As RECT) As Long
''#define ListView_GetItemIndexRect(hwnd, plvii, iSubItem, code, prc) \
''        (BOOL)SNDMSG((hwnd), LVM_GETITEMINDEXRECT, (WPARAM)(LVITEMINDEX*)(plvii), \
''                ((prc) ? ((((LPRECT)(prc))->top = (iSubItem)), (((LPRECT)(prc))->left = (code)), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))
'    prc.Top = iSubItem
'    prc.Left = Code
'    ListView_GetItemIndexRect = SendMessage(hwnd, LVM_GETITEMINDEXRECT, VarPtr(lvii), prc)
'End Function
'
''#define LVM_GETNEXTITEMINDEX    (LVM_FIRST + 211)
''#define ListView_GetNextItemIndex(hwnd, plvii, flags) \
''    (BOOL)SNDMSG((hwnd), LVM_GETNEXTITEMINDEX, (WPARAM)(LVITEMINDEX*)(plvii), MAKELPARAM((flags), 0))
'Public Function ListView_GetNextItemIndex(hwnd As Long, plvii As LVITEMINDEX, ByVal Flags As LVNI_Flags) As Long
' '#define ListView_GetNextItemIndex(hwnd, plvii, flags) \
' '    (BOOL)SNDMSG((hwnd), LVM_GETNEXTITEMINDEX, (WPARAM)(LVITEMINDEX*)(plvii), MAKELPARAM((flags), 0))
'    ListView_GetNextItemIndex = SendMessage(hwnd, LVM_GETNEXTITEMINDEX, VarPtr(plvii), ByVal Flags)
'End Function
'
''#define LVM_GETOUTLINECOLOR     (LVM_FIRST + 176)
''#define ListView_GetOutlineColor(hwnd) \
''    (COLORREF)SNDMSG((hwnd), LVM_GETOUTLINECOLOR, 0, 0)
'Public Function ListView_GetOutlineColor(hwnd As Long) As Long
''#define ListView_GetOutlineColor(hwnd) \
''    (COLORREF)SNDMSG((hwnd), LVM_GETOUTLINECOLOR, 0, 0)
'    ListView_GetOutlineColor = SendMessage(hwnd, LVM_GETOUTLINECOLOR, 0, ByVal 0&)
'End Function
'
''#define LVM_SETOUTLINECOLOR     (LVM_FIRST + 177)
''#define ListView_SetOutlineColor(hwnd, color) \
''    (COLORREF)SNDMSG((hwnd), LVM_SETOUTLINECOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
'Public Function ListView_SetOutlineColor(hwnd As Long, Color As Long) As Long
''#define ListView_SetOutlineColor(hwnd, color) \
''    (COLORREF)SNDMSG((hwnd), LVM_SETOUTLINECOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
'    ListView_SetOutlineColor = SendMessage(hwnd, LVM_SETOUTLINECOLOR, 0, ByVal Color)
'End Function
'
''#define LVM_GETSELECTEDCOLUMN   (LVM_FIRST + 174)
''#define ListView_GetSelectedColumn(hwnd) \
''    (UINT)SNDMSG((hwnd), LVM_GETSELECTEDCOLUMN, 0, 0)
'Public Function ListView_GetSelectedColumn(hwnd As Long) As Long
''#define ListView_GetSelectedColumn(hwnd) \
''    (UINT)SNDMSG((hwnd), LVM_GETSELECTEDCOLUMN, 0, 0)
'    ListView_GetSelectedColumn = SendMessage(hwnd, LVM_GETSELECTEDCOLUMN, 0, ByVal 0&)
'End Function
'
''#define LVM_GETTILEINFO                     (LVM_FIRST + 165)
''#define ListView_GetTileInfo(hwnd, pti) \
''    SNDMSG((hwnd), LVM_GETTILEINFO, 0, (LPARAM)(pti))
'Public Function ListView_GetTileInfo(hwnd As Long, pTI As LVTILEINFO) As Long
''#define ListView_GetTileInfo(hwnd, pti) \
''    SNDMSG((hwnd), LVM_GETTILEINFO, 0, (LPARAM)(pti))
'    ListView_GetTileInfo = SendMessage(hwnd, LVM_GETTILEINFO, 0, pTI)
'End Function
'
''#define LVM_SETTILEINFO                     (LVM_FIRST + 164)
''#define ListView_SetTileInfo(hwnd, pti) \
''    SNDMSG((hwnd), LVM_SETTILEINFO, 0, (LPARAM)(pti))
'Public Function ListView_SetTileInfo(hwnd As Long, pTI As LVTILEINFO) As Long
''#define ListView_SetTileInfo(hwnd, pti) \
''    SNDMSG((hwnd), LVM_SETTILEINFO, 0, (LPARAM)(pti))
'    ListView_SetTileInfo = SendMessage(hwnd, LVM_SETTILEINFO, 0, pTI)
'End Function
'
''#define LVM_GETTILEVIEWINFO                 (LVM_FIRST + 163)
''#define ListView_GetTileViewInfo(hwnd, ptvi) \
''    SNDMSG((hwnd), LVM_GETTILEVIEWINFO, 0, (LPARAM)(ptvi))
'Public Function ListView_GetTileViewInfo(hwnd As Long, ptvi As LVTILEVIEWINFO) As Long
''#define ListView_GetTileViewInfo(hwnd, ptvi) \
''    SNDMSG((hwnd), LVM_GETTILEVIEWINFO, 0, (LPARAM)(ptvi))
'    ListView_GetTileViewInfo = SendMessage(hwnd, LVM_GETTILEVIEWINFO, 0, ptvi)
'End Function
'
''typedef struct tagLVTILEVIEWINFO
''{
''    UINT    cbSize;
''    DWORD   dwMask;     //LVTVIM_*
''    DWORD   dwFlags;    //LVTVIF_*
''    SIZE    sizeTile;
''    int     cLines;
''    RECT    rcLabelMargin;
''} LVTILEVIEWINFO, *PLVTILEVIEWINFO;
''
''typedef struct tagLVTILEINFO
''{
''    UINT    cbSize;
''    int     iItem;
''    UINT    cColumns;
''    PUINT   puColumns;
''#if (NTDDI_VERSION >= NTDDI_VISTA)
''    int*    piColFmt;
''#End If
''} LVTILEINFO, *PLVTILEINFO;
''
''#define LVTILEINFO_V5_SIZE CCSIZEOF_STRUCT(LVTILEINFO, puColumns)
''
''#define LVM_SETTILEVIEWINFO                 (LVM_FIRST + 162)
''#define ListView_SetTileViewInfo(hwnd, ptvi) \
''    SNDMSG((hwnd), LVM_SETTILEVIEWINFO, 0, (LPARAM)(ptvi))
'Public Function ListView_SetTileViewInfo(hwnd As Long, ptvi As LVTILEVIEWINFO) As Long
''#define ListView_SetTileViewInfo(hwnd, ptvi) \
''    SNDMSG((hwnd), LVM_SETTILEVIEWINFO, 0, (LPARAM)(ptvi))
'    ListView_SetTileViewInfo = SendMessage(hwnd, LVM_SETTILEVIEWINFO, 0, ptvi)
'End Function
'
''#define LVM_HASGROUP                    (LVM_FIRST + 161)
''#define ListView_HasGroup(hwnd, dwGroupId) \
''    SNDMSG((hwnd), LVM_HASGROUP, dwGroupId, 0)
'Public Function ListView_HasGroup(hwnd As Long, dwGroupId As Long) As Long
''#define ListView_HasGroup(hwnd, dwGroupId) \
''    SNDMSG((hwnd), LVM_HASGROUP, dwGroupId, 0)
'    ListView_HasGroup = SendMessage(hwnd, LVM_HASGROUP, dwGroupId, ByVal 0&)
'End Function
'
''#define ListView_HitTestEx(hwndLV, pinfo) \
''    (int)SNDMSG((hwndLV), LVM_HITTEST, (WPARAM)-1, (LPARAM)(LV_HITTESTINFO *)(pinfo))
'Public Function ListView_HitTestEx(hwndLV As Long, pInfo As LVHITTESTINFO) As Long
''HitTestEx is used if you need the iGroup and iSubItem members filled
'    ListView_HitTestEx = SendMessage(hwndLV, LVM_HITTEST, -1, pInfo)
'End Function
'
''#define LVM_INSERTGROUP         (LVM_FIRST + 145)
''#define ListView_InsertGroup(hwnd, index, pgrp) \
''    SNDMSG((hwnd), LVM_INSERTGROUP, (WPARAM)(index), (LPARAM)(pgrp))
'Public Function ListView_InsertGroup(hwnd As Long, Index As Long, pgrp As lvGroup) As Long
''#define ListView_InsertGroup(hwnd, index, pgrp) \
''    SNDMSG((hwnd), LVM_INSERTGROUP, (WPARAM)(index), (LPARAM)(pgrp))
'    ListView_InsertGroup = SendMessage(hwnd, LVM_INSERTGROUP, Index, pgrp)
'End Function
'
''typedef struct tagLVINSERTGROUPSORTED
''{
''    PFNLVGROUPCOMPARE pfnGroupCompare;
''    void *pvData;
''    LVGROUP lvGroup;
''}LVINSERTGROUPSORTED, *PLVINSERTGROUPSORTED;
''
''#define LVM_INSERTGROUPSORTED           (LVM_FIRST + 159)
''#define ListView_InsertGroupSorted(hwnd, structInsert) \
''    SNDMSG((hwnd), LVM_INSERTGROUPSORTED, (WPARAM)(structInsert), 0)
'Public Function ListView_InsertGroupSorted(hwnd As Long, structInsert As LVINSERTGROUPSORTED) As Long
''#define ListView_InsertGroupSorted(hwnd, structInsert) \
''    SNDMSG((hwnd), LVM_INSERTGROUPSORTED, (WPARAM)(structInsert), 0)
'    ListView_InsertGroupSorted = SendMessage(hwnd, LVM_INSERTGROUPSORTED, VarPtr(structInsert), ByVal 0&)
'End Function
'
''#define LVM_ISGROUPVIEWENABLED  (LVM_FIRST + 175)
''#define ListView_IsGroupViewEnabled(hwnd) \
''    (BOOL)SNDMSG((hwnd), LVM_ISGROUPVIEWENABLED, 0, 0)
'Public Function ListView_IsGroupViewEnabled(hwnd As Long) As Long
''#define ListView_IsGroupViewEnabled(hwnd) \
''    (BOOL)SNDMSG((hwnd), LVM_ISGROUPVIEWENABLED, 0, 0)
'    ListView_IsGroupViewEnabled = SendMessage(hwnd, LVM_ISGROUPVIEWENABLED, 0, ByVal 0&)
'End Function
'
''#define LVM_ISITEMVISIBLE    (LVM_FIRST + 182)
''#define ListView_IsItemVisible(hwnd, index) \
''    (UINT)SNDMSG((hwnd), LVM_ISITEMVISIBLE, (WPARAM)(index), (LPARAM)0)
'Public Function ListView_IsItemVisible(hwnd As Long, Index As Long) As Long
''#define ListView_IsItemVisible(hwnd, index) \
''    (UINT)SNDMSG((hwnd), LVM_ISITEMVISIBLE, (WPARAM)(index), (LPARAM)0)
'    ListView_IsItemVisible = SendMessage(hwnd, LVM_ISITEMVISIBLE, Index, ByVal 0&)
'End Function
'
''#define LVM_MAPIDTOINDEX     (LVM_FIRST + 181)
''#define ListView_MapIDToIndex(hwnd, id) \
''    (UINT)SNDMSG((hwnd), LVM_MAPIDTOINDEX, (WPARAM)(id), (LPARAM)0)
'Public Function ListView_MapIDToIndex(hwnd As Long, id As Long) As Long
''#define ListView_MapIDToIndex(hwnd, id) \
''    (UINT)SNDMSG((hwnd), LVM_MAPIDTOINDEX, (WPARAM)(id), (LPARAM)0)
'    ListView_MapIDToIndex = SendMessage(hwnd, LVM_MAPIDTOINDEX, id, ByVal 0&)
'End Function
'
''// These next to methods make it easy to identify an item that can be repositioned
''// within listview. For example: Many developers use the lParam to store an identifier that is
''// unique. Unfortunatly, in order to find this item, they have to iterate through all of the items
''// in the listview. Listview will maintain a unique identifier.  The upper bound is the size of a DWORD.
''#define LVM_MAPINDEXTOID     (LVM_FIRST + 180)
''#define ListView_MapIndexToID(hwnd, index) \
''    (UINT)SNDMSG((hwnd), LVM_MAPINDEXTOID, (WPARAM)(index), (LPARAM)0)
'Public Function ListView_MapIndexToID(hwnd As Long, Index As Long) As Long
''#define ListView_MapIndexToID(hwnd, index) \
''    (UINT)SNDMSG((hwnd), LVM_MAPINDEXTOID, (WPARAM)(index), (LPARAM)0)
'    ListView_MapIndexToID = SendMessage(hwnd, LVM_MAPINDEXTOID, Index, ByVal 0&)
'End Function
'
''#define LVM_MOVEGROUP         (LVM_FIRST + 151)
''#define ListView_MoveGroup(hwnd, iGroupId, toIndex) \
''    SNDMSG((hwnd), LVM_MOVEGROUP, (WPARAM)(iGroupId), (LPARAM)(toIndex))
'Public Function ListView_MoveGroup(hwnd As Long, iGroupId As Long, toIndex As Long) As Long
''NOT IMPLEMENTED
''#define ListView_MoveGroup(hwnd, iGroupId, toIndex) \
''    SNDMSG((hwnd), LVM_MOVEGROUP, (WPARAM)(iGroupId), (LPARAM)(toIndex))
'    ListView_MoveGroup = SendMessage(hwnd, LVM_MOVEGROUP, iGroupId, ByVal toIndex)
'End Function
'
''#define LVM_MOVEITEMTOGROUP            (LVM_FIRST + 154)
''#define ListView_MoveItemToGroup(hwnd, idItemFrom, idGroupTo) \
''    SNDMSG((hwnd), LVM_MOVEITEMTOGROUP, (WPARAM)(idItemFrom), (LPARAM)(idGroupTo))
'Public Function ListView_MoveItemToGroup(hwnd As Long, idItemFrom As Long, idGroupTo As Long) As Long
''NOT IMPLEMENTED
''#define ListView_MoveItemToGroup(hwnd, idItemFrom, idGroupTo) \
''    SNDMSG((hwnd), LVM_MOVEITEMTOGROUP, (WPARAM)(idItemFrom), (LPARAM)(idGroupTo))
'    ListView_MoveItemToGroup = SendMessage(hwnd, LVM_MOVEITEMTOGROUP, idItemFrom, ByVal idGroupTo)
'End Function
'
''#define LVM_REMOVEALLGROUPS             (LVM_FIRST + 160)
''#define ListView_RemoveAllGroups(hwnd) \
''    SNDMSG((hwnd), LVM_REMOVEALLGROUPS, 0, 0)
'Public Function ListView_RemoveAllGroups(hwnd As Long) As Long
''#define ListView_RemoveAllGroups(hwnd) \
''    SNDMSG((hwnd), LVM_REMOVEALLGROUPS, 0, 0)
'    ListView_RemoveAllGroups = SendMessage(hwnd, LVM_REMOVEALLGROUPS, 0, ByVal 0&)
'End Function
'
''#define LVM_REMOVEGROUP         (LVM_FIRST + 150)
''#define ListView_RemoveGroup(hwnd, iGroupId) \
''    SNDMSG((hwnd), LVM_REMOVEGROUP, (WPARAM)(iGroupId), 0)
'Public Function ListView_RemoveGroup(hwnd As Long, iGroupId As Long) As Long
''#define ListView_RemoveGroup(hwnd, iGroupId) \
''    SNDMSG((hwnd), LVM_REMOVEGROUP, (WPARAM)(iGroupId), 0)
'    ListView_RemoveGroup = SendMessage(hwnd, LVM_REMOVEGROUP, iGroupId, ByVal 0&)
'End Function
'
''typedef struct tagLVSETINFOTIP
''{
''    UINT cbSize;
''    DWORD dwFlags;
''    LPWSTR pszText;
''    int iItem;
''    int iSubItem;
''} LVSETINFOTIP, *PLVSETINFOTIP;
''
''#define  LVM_SETINFOTIP         (LVM_FIRST + 173)
''
''#define ListView_SetInfoTip(hwndLV, plvInfoTip)\
''        (BOOL)SNDMSG((hwndLV), LVM_SETINFOTIP, (WPARAM)0, (LPARAM)(plvInfoTip))
'Public Function ListView_SetInfoTip(hwnd As Long, plvInfoTip As LVSETINFOTIP) As Long
''#define ListView_SetInfoTip(hwndLV, plvInfoTip)\
''        (BOOL)SNDMSG((hwndLV), LVM_SETINFOTIP, (WPARAM)0, (LPARAM)(plvInfoTip))
'    ListView_SetInfoTip = SendMessage(hwnd, LVM_SETINFOTIP, 0, plvInfoTip)
'End Function
'
''#define LVM_SETITEMINDEXSTATE   (LVM_FIRST + 210)
''#define ListView_SetItemIndexState(hwndLV, plvii, data, mask) \
''{ LV_ITEM _macro_lvi;\
''  _macro_lvi.stateMask = (mask);\
''  _macro_lvi.state = (data);\
''  SNDMSG((hwndLV), LVM_SETITEMINDEXSTATE, (WPARAM)(LVITEMINDEX*)(plvii), (LPARAM)(LV_ITEM *)&_macro_lvi);\
''}
'Public Function ListView_SetItemIndexState(hwndLV As Long, plvii As LVITEMINDEX, Data As Long, mask As Long) As Long
''#define ListView_SetItemIndexState(hwndLV, plvii, data, mask) \
''{ LV_ITEM _macro_lvi;\
''  _macro_lvi.stateMask = (mask);\
''  _macro_lvi.state = (data);\
''  SNDMSG((hwndLV), LVM_SETITEMINDEXSTATE, (WPARAM)(LVITEMINDEX*)(plvii), (LPARAM)(LV_ITEM *)&_macro_lvi);\}
'
'    Dim lvi As LVITEM
'    lvi.stateMask = mask
'    lvi.State = Data
'    ListView_SetItemIndexState = SendMessage(hwndLV, LVM_SETITEMINDEXSTATE, VarPtr(plvii), lvi)
'End Function
'
''#define LVM_SETSELECTEDCOLUMN         (LVM_FIRST + 140)
''#define ListView_SetSelectedColumn(hwnd, iCol) \
''    SNDMSG((hwnd), LVM_SETSELECTEDCOLUMN, (WPARAM)(iCol), 0)
'Public Function ListView_SetSelectedColumn(hwnd As Long, iCol As Long) As Long
''#define ListView_SetSelectedColumn(hwnd, iCol) \
''    SNDMSG((hwnd), LVM_SETSELECTEDCOLUMN, (WPARAM)(iCol), 0)
'    ListView_SetSelectedColumn = SendMessage(hwnd, LVM_SETSELECTEDCOLUMN, iCol, ByVal 0&)
'End Function
'
''#define LVM_SORTGROUPS         (LVM_FIRST + 158)
''#define ListView_SortGroups(hwnd, _pfnGroupCompate, _plv) \
''    SNDMSG((hwnd), LVM_SORTGROUPS, (WPARAM)(_pfnGroupCompate), (LPARAM)(_plv))
'

Public Function WndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'    Dim iCode As LongPtr
'    Dim lvi As LVITEM
'    Dim nlv As NMLISTVIEW
'    Dim nia As NMITEMACTIVATE  ' ←追加

'    If hWnd = TT.hParent Then
    If hwnd = This.hWndFrm Then
        Call ShowMessage(hwnd)
        Call ShowMessage2(This.hWndFrm)
        Call ShowMessage3(uMsg)
        Select Case uMsg

            Case WM_SETFOCUS
                SetFocus This.hwndLV
                Call ShowMessage(lParam)
                Exit Function

            Case WM_NOTIFY
    ''            MoveMemory iCode, ByVal lParam + 8, 4
    '
    ''            Select Case iCode
                Call ShowMessage(lParam)
                Select Case lParam

'                    Case LVN_ITEMCHANGED 'ItemSelected
'                        MoveMemory nlv, ByVal lParam, LenB(nlv)
'                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
'                        Exit Function
'
'                    Case NM_CLICK 'ItemClick
'                        MoveMemory nia, ByVal lParam, LenB(nia)
'                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'                        Exit Function
'
'                    Case NM_DBLCLK
'                        MoveMemory nia, ByVal lParam, LenB(nia)
'                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'                        Exit Function

                    'Case いろいろ...
                    '      :
                    '      :
                    Case Else

                End Select
'
        End Select

    ElseIf hwnd = This.hwndLV Then
        Select Case uMsg

            Case WM_VSCROLL
                Debug.Print "ListView:WM_VSCROLL"
            Case WM_HSCROLL
                Debug.Print "ListView:WM_HSCROLL"
            Case WM_MOUSEWHEEL
                Debug.Print "ListView:WM_MOUSEWHEEL:" & wParam
            
            Case WM_NOTIFY
'            Case WM_MOUSEACTIVATE
''                If GetFocus() <> hWnd Then
''                    acc.accSelect 1&
''                    WndProc = MA_NOACTIVATE
''                    Exit Function
''                End If
''
            Case WM_NCDESTROY
                Call RemoveSubClass
                Exit Function
'           '----------------------------
'            Case WM_KEYDOWN
''                Select Case wParam
''                    Case VK_RETURN
''                        Exit Function
''                    Case VK_UP, VK_DOWN
''                        WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
''                        Exit Function
''                    Case Else
''                End Select
''                Exit Function
'       '-----------------------------
            Case Else
        End Select
    End If
    WndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)

End Function

'Friend Function WndProcLV(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
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
'                Debug.Print "ListView:WM_VSCROLL"
'            Case WM_HSCROLL
'                Debug.Print "ListView:WM_HSCROLL"
'            Case Else
'        End Select
''    End If
'    WndProcLV = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'
'End Function
'
'Friend Function WndProcFrm(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
''    If hWnd = This.hWndFrm Then
''            Call ShowMessage(hWnd)
''        Call ShowMessage2(This.hWndFrm)
'        Call ShowMessage3(uMsg)
'        Select Case uMsg
'
'            Case WM_SETFOCUS
'                SetFocus This.hwndLV
'                Call ShowMessage(lParam)
'                Exit Function
'
'            Case WM_VSCROLL
'                Debug.Print "Frame:WM_VSCROLL"
'            Case WM_HSCROLL
'                Debug.Print "Frame:WM_HSCROLL"
'
'            Case WM_NOTIFY
'    ''            MoveMemory iCode, ByVal lParam + 8, 4
'    '
'    ''            Select Case iCode
'                Call ShowMessage(lParam)
'                Select Case lParam
'
''                    Case LVN_ITEMCHANGED 'ItemSelected
''                        MoveMemory nlv, ByVal lParam, LenB(nlv)
''                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
''                        Exit Function
''
''                    Case NM_CLICK 'ItemClick
''                        MoveMemory nia, ByVal lParam, LenB(nia)
''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
''                        Exit Function
''
''                    Case NM_DBLCLK
''                        MoveMemory nia, ByVal lParam, LenB(nia)
''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
''                        Exit Function
'
'                    'Case いろいろ...
'                    '      :
'                    '      :
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
'
Private Function GetPtr(ByVal ptr As LongPtr) As LongPtr
    GetPtr = ptr
End Function

Private Sub RemoveSubClass()
    With This
        If .hwndLV Then
            RemoveWindowSubclass .hwndLV, .pfn, .hwndLV
            .hwndLV = 0
        End If

        If .hWndFrm Then
            RemoveWindowSubclass .hWndFrm, .pfn, .hWndFrm
            .hWndFrm = 0
        End If
    End With
End Sub

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

'#define MAKEWORD(a, b)      ((WORD)(((BYTE)(((DWORD_PTR)(a)) & 0xff)) | ((WORD)((BYTE)(((DWORD_PTR)(b)) & 0xff))) << 8))

'#define MAKELONG(a, b)      ((LONG)(((WORD)(((DWORD_PTR)(a)) & 0xffff)) | ((DWORD)((WORD)(((DWORD_PTR)(b)) & 0xffff))) << 16))
Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

'#define MAKELPARAM(l, h)      ((LPARAM)(DWORD)MAKELONG(l, h))
Private Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
    MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

'#define MAKEWPARAM(l, h)      ((WPARAM)(DWORD)MAKELONG(l, h))

'#define MAKELRESULT(l, h)     ((LRESULT)(DWORD)MAKELONG(l, h))

'#define LOWORD(l)           ((WORD)(((DWORD_PTR)(l)) & 0xffff))
Public Function LoWord(ByVal dwValue As Long) As Integer
' Returns the low 16-bit integer from a 32-bit long integer
    CopyMemory LoWord, dwValue, 2&
End Function

'#define HIWORD(l)           ((WORD)((((DWORD_PTR)(l)) >> 16) & 0xffff))

'#define LOBYTE(w)           ((BYTE)(((DWORD_PTR)(w)) & 0xff))

'#define HIBYTE(w)           ((BYTE)((((DWORD_PTR)(w)) >> 8) & 0xff))

Private Function ProcMouseWheel(wParam As LongPtr) As Integer
'#define GET_WHEEL_DELTA_WPARAM(wParam)  ((short)HIWORD(wParam))

    ProcMouseWheel = CInt(wParam)
End Function

Private Function ProcNotify(hwnd As LongPtr, lParam As LongPtr) As Long
    Dim sText As String, sSubText As String
    Dim tNMH As NMHDR
    CopyMemory tNMH, ByVal lParam, Len(tNMH)

    Select Case tNMH.Code

        Case LVN_GETDISPINFOW
'               Debug.Print "GetDispInfo"
            Dim LVDI As NMLVDISPINFO
            CopyMemory ByVal VarPtr(LVDI), ByVal lParam, LenB(LVDI)
            With LVDI.Item
                
                If (.mask And LVIF_TEXT) Then
                    Select Case .iSubItem
                        Case 0
'                                sText = ItemTxt(.iItem)
'                                .cchTextMax = Len(sText)
                            .pszText = StrPtr(VLItems(.iItem).sText)
                        Case Else
                            'sSubText = "Subitem " & .iSubItem
'                                .cchTextMax = Len(sSubText)
                            .pszText = StrPtr(VLItems(.iItem).sSubItems(.iSubItem - 1))
'                                Debug.Print "subitemtext=" & StrPtr(sSubText)
                    End Select
                End If
                
                If (.mask And LVIF_IMAGE) Then
                    Select Case .iSubItem
                        Case 0
                            .IImage = VLItems(.iItem).IImage
                    End Select
                End If
                CopyMemory ByVal lParam, ByVal VarPtr(LVDI), LenB(LVDI)
            End With
        End Select
    Exit Function
End Function
