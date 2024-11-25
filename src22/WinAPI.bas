Attribute VB_Name = "WinAPI"
Option Explicit
Option Private Module

'Windows API宣言
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Declare PtrSafe Function CreateMenu Lib "user32" () As LongPtr
Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
Declare PtrSafe Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As Any) As Long
Declare PtrSafe Function SetMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal hMenu As LongPtr) As Long
Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr

Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As LongPtr
Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr

#If Win64 Then
Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'Windows API定数
Public Const GWL_STYLE         As Long = (-16)
Public Const SC_CLOSE          As Long = &HF060        '閉じるボタン
Public Const MF_BYCOMMAND      As Long = &H0&          '定数の設定
Public Const WS_MAXIMIZEBOX    As LongPtr = &H10000
Public Const WS_MINIMIZEBOX    As LongPtr = &H20000
Public Const WS_THICKFRAME     As LongPtr = &H40000
Public Const GWL_WNDPROC       As Long = (-4)

Public Const MF_MENUBREAK = &H40&      '水平
Public Const MF_MENUBARBREAK = &H20&   '区切り
Public Const MF_POPUP = &H10&
Public Const MF_OWNERDRAW = &O100&     'オーナードローで表示
Public Const MF_SEPARATOR = 800&       'セパレータ
Public Const MF_ENABLED = &H0
Public Const MF_STRING = &H0&
Public Const MF_HELP = &H4000&
Public Const MFS_DEFAULT = &H1000&

' Window Messages
Public Const WM_COMMAND = &H111

Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long

Declare PtrSafe Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Declare PtrSafe Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As LongPtr, ByVal pszPath As String, ByVal dx As Long) As Long

'API定義
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long

Declare PtrSafe Function GetInputState Lib "user32" () As Long

Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

'List-view messages
Public Const LVM_FIRST = &H1000&
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    
Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpszName As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As LongPtr

'Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal acc As IAccessible, ByRef hwnd As LongPtr) As Long
Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long

