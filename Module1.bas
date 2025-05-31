Attribute VB_Name = "Module1"
Option Explicit

'API宣言部（標準モジュール）modAPI
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_EX_CLIENTEDGE = &H200&

Public Const WM_SETFOCUS = 7&
Public Const WM_NOTIFY = &H4E&
Public Const WM_MOUSEACTIVATE = &H21&
Public Const WM_NCDESTROY = &H82&
Public Const WM_SETFONT = &H30

Public Const MA_NOACTIVATE = 3&

Public Const LVS_REPORT = 1&
Public Const LVS_SHOWSELALWAYS = 8&
Public Const LVS_EX_FULLROWSELECT = &H20&
Public Const LVS_EX_GRIDLINES = 1&
Public Const LVS_EX_CHECKBOXES = 4&

Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = &H1036
Public Const LVM_SETITEM = &H1000& + 76
Public Const LVM_INSERTITEM = &H1000& + 77
Public Const LVM_INSERTCOLUMN = &H1000& + 97
Public Const LVM_GETITEMTEXT = &H1000& + 115

Public Const LVN_ITEMCHANGED = -100& - 1

Public Const LVCF_WIDTH = 2&
Public Const LVCF_TEXT = 4&
Public Const LVCF_SUBITEM = 8&
Public Const LVIF_TEXT = 1&

Public Const LVM_FIRST = &H1000
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38) ' テキストの背景色を設定
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)      ' 背景色の設定
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)   ' テキストの文字色を設定

Public Const NM_CLICK = -2&
Public Const NM_DBLCLK = -3&

Type InitCommonControlsExType
     dwSize As Long
     dwICC As Long
End Type

Declare PtrSafe Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExType) As Boolean
'Declare Function InitCommonControlsEx& Lib "comctl32" _
    (ByVal lpInitCtrls&)
    
Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
'Declare Function SetWindowSubclass& Lib "comctl32" _
    (ByVal hwnd&, _
     ByVal pfnSubclass&, _
     ByVal uIdSubclass&, _
     ByVal dwRefData&)
Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'Declare Function DefSubclassProc& Lib "comctl32" _
    (ByVal hwnd&, _
     ByVal uMsg&, _
     ByVal wParam&, _
     ByVal lParam&)
'Declare Function RemoveWindowSubclass& Lib "comctl32" _
    (ByVal hwnd&, _
     ByVal pfnSubclass&, _
     ByVal uIdSubclass&)
Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
'Declare Function CreateWindowExW& Lib "user32" _
    (ByVal dwExStyle&, _
     ByVal lpClassName&, _
     ByVal lpWindowName&, _
     ByVal dwStyle&, _
     ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, _
     ByVal HwndParent&, _
     ByVal HMENU&, _
     ByVal hInstance&, _
     ByVal lpParam&)
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'Declare Function SendMessageW& Lib "user32" _
    (ByVal hwnd&, _
     ByVal uMsg&, _
     ByVal wParam&, _
     ByVal lParam&)

Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
'Declare Function GetFocus& Lib "user32" ()
'Declare Function SetFocus& Lib "user32" (ByVal hwnd&)

Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, _
     pSrc As Any, _
     ByVal cbLen&)
Declare PtrSafe Function SysAllocString Lib "OleAut32.dll" (ByVal psz As LongPtr) As LongPtr
'Declare Function SysAllocString& Lib "Oleaut32" (ByVal ptr&)

Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As LongPtr
'Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
    (ByVal H As Long, _
     ByVal W As Long, _
     ByVal E As Long, _
     ByVal o As Long, _
     ByVal W As Long, _
     ByVal i As Long, _
     ByVal u As Long, _
     ByVal S As Long, _
     ByVal C As Long, _
     ByVal OP As Long, _
     ByVal CP As Long, _
     ByVal Q As Long, _
     ByVal PAF As Long, _
     ByVal F As String) As Long

Type LVCOLUMN
    mask        As Long
    fmt         As Long
    cx          As Long
    pszText     As LongPtr
    cchTextMax  As Long
    iSubItem    As Long
    iImage      As Long
    buf(3)      As Long
End Type

Type LVITEM
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    state       As Long
    stateMask   As Long
    pszText     As LongPtr
    cchTextMax  As Long
    buf(7)      As Long
End Type

Type NMITEMACTIVATE
    hdr(2)      As Long
    iItem       As Long
    iSubItem    As Long
    buf(6)      As Long
End Type

Type NMLISTVIEW
    hrd(2)      As Long
    iItem       As Long
    iSubItem    As Long
    buf(4)      As Long
End Type
 
Type TT
    hParent As LongPtr
    hChild As LongPtr
    pfn As LongPtr
End Type

Public TT As TT, acc As IAccessible

Public Function Redirect(ByVal hWnd As LongPtr, ByVal uMsg&, ByVal wParam&, _
                    ByVal lParam&, ByVal id&, ByVal lv As ListView) As LongPtr
    Redirect = lv.WndProc(hWnd, uMsg, wParam, lParam)
    
End Function




