Attribute VB_Name = "VListViewModule"
Option Explicit

Public ptr As Long

Public Type VListItem
    sText As String
    sSubItems() As String
    iImage As Long
    iSubItemImages() As Long 'LVS_EX_SUBITEMIMAGES must be enabled, then must dim same as sSubItems
    iGrp As Long
    iPos As Long
End Type

Public Type VListGroup
    items() As Long
    gid As Long 'groupid, doesn't have to be the same as the index
                'but in the case of virtual groups should be, since
                'alot of stuff goes by index
End Type

Public VLItems() As VListItem
Public VLGroups() As VListGroup
Public lGroupCount As Long

'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

'Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long

'Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#If Win64 Then
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'Public Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Public Declare PtrSafe Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal pszSubAppName As LongPtr, ByVal pszSubIdList As LongPtr) As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

'Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As Long
Public Declare PtrSafe Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As LongPtr

'Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr


'Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr

'Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long

'Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare PtrSafe Sub InitCommonControls Lib "COMCTL32" ()

Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type


'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

'Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
'Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long

'Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpDest As Long, ByVal lpSource As Long) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr

'Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Const CCM_FIRST = &H2000
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Public Const WM_DESTROY = &H2
Public Const WM_NOTIFYFORMAT = &H55
Public Const NFR_UNICODE = 2
Public Const WM_NOTIFY = &H4E

Public Enum CBoolean
  Cfalse = 0
  CTrue = 1
End Enum

Public Const MAX_PATH = 260

' Message structure

' Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
'Public Enum GWL_nIndex
'  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
'  GWL_ID = (-12)
'  GWL_STYLE = (-16)
'  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
'End Enum

' // Begin SHGetFileInfo
Public Type SHFILEINFO   ' shfi
    hIcon As LongPtr
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
'Public Type SHFILEINFO
'    hIcon As LongPtr                         '  out: icon
'    iIcon As Long                            '  out: icon index
'    dwAttributes As Long                     '  out: SFGAO_ flags
'    szDisplayName(0 To MAX_PATH - 1) As Byte '  out: display name (or path)
'    szTypeName(0 To 79) As Byte              '  out: type name
'End Type

Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000&                '  put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_OPENICON = &H2                       '  get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
'Public Enum SHGFI_flags
'  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
'  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
'  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
'  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
'  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
'  ' Indicates that the function should not attempt to access the file specified by pszPath.
'  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
'  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
'  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
'  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
'  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
'  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
'  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
'  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
'  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
'                                                        ' containing the icon, rtns BOOL
'  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
'  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
'  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
'  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
'  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
'End Enum

' Window Styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

' Common Window Styles
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Public Const WS_CHILDWINDOW = (WS_CHILD)

' Extended Window Styles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
  
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LTRREADING = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_RIGHTSCROLLBAR = &H0
  
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000
  
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
'Public Enum WinStyles
'  WS_OVERLAPPED = &H0
'  WS_TABSTOP = &H10000
'  WS_MAXIMIZEBOX = &H10000
'  WS_MINIMIZEBOX = &H20000
'  WS_GROUP = &H20000
'  WS_THICKFRAME = &H40000
'  WS_SYSMENU = &H80000
'  WS_HSCROLL = &H100000
'  WS_VSCROLL = &H200000
'  WS_DLGFRAME = &H400000
'  WS_BORDER = &H800000
'  WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
'  WS_MAXIMIZE = &H1000000
'  WS_CLIPCHILDREN = &H2000000
'  WS_CLIPSIBLINGS = &H4000000
'  WS_DISABLED = &H8000000
'  WS_VISIBLE = &H10000000
'  WS_MINIMIZE = &H20000000
'  WS_CHILD = &H40000000
'  WS_POPUP = &H80000000
'
'  WS_TILED = WS_OVERLAPPED
'  WS_ICONIC = WS_MINIMIZE
'  WS_SIZEBOX = WS_THICKFRAME
'
'  ' Common Window Styles
'  WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
'  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
'  WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
'  WS_CHILDWINDOW = WS_CHILD
'End Enum   ' WinStyles
'
'Public Enum WinStylesEx
'  WS_EX_DLGMODALFRAME = &H1
'  WS_EX_NOPARENTNOTIFY = &H4
'  WS_EX_TOPMOST = &H8
'  WS_EX_ACCEPTFILES = &H10
'  WS_EX_TRANSPARENT = &H20
'
'  WS_EX_MDICHILD = &H40
'  WS_EX_TOOLWINDOW = &H80
'  WS_EX_WINDOWEDGE = &H100
'  WS_EX_CLIENTEDGE = &H200
'  WS_EX_CONTEXTHELP = &H400
'
'  WS_EX_RIGHT = &H1000
'  WS_EX_LEFT = &H0
'  WS_EX_RTLREADING = &H2000
'  WS_EX_LTRREADING = &H0
'  WS_EX_LEFTSCROLLBAR = &H4000
'  WS_EX_RIGHTSCROLLBAR = &H0
'
'  WS_EX_CONTROLPARENT = &H10000
'  WS_EX_STATICEDGE = &H20000
'  WS_EX_APPWINDOW = &H40000
'
'  WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
'  WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
'End Enum   ' WinStylesEx

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type NMHDR
    hwndFrom As LongPtr
    idfrom As LongPtr
    code As Long
End Type
'Public Type NMHDR
'  hwndFrom As Long   ' Window handle of control sending message
'  idfrom As Long        ' Identifier of control sending message
'  code  As Long          ' Specifies the notification code
'End Type

'Private OldWndProc As Long
Private OldWndProc As LongPtr
Private schWnd As Long
Public ItemTxt(99) As String
Public SubItemTxt(99) As String

'Public Function Subclass2(hWnd As Long, lpfn As Long, Optional uId As Long = 0&, Optional dwRefData As Long = 0&) As Boolean
Public Function Subclass2(hWnd As LongPtr, lpfn As LongPtr, Optional uId As LongPtr = 0, Optional dwRefData As LongPtr = 0) As Boolean
    If uId = 0 Then uId = hWnd
        Subclass2 = SetWindowSubclass(hWnd, lpfn, uId, dwRefData):      Debug.Assert Subclass2
End Function

'Public Function UnSubclass2(hWnd As Long, ByVal lpfn As Long, pid As Long) As Boolean
Public Function UnSubclass2(hWnd As LongPtr, ByVal lpfn As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpfn, pid)
End Function

'Public Function FGVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Public Function FGVWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_NOTIFYFORMAT
            'Debug.Print "Got NFMT on ftv main"
            FGVWndProc = NFR_UNICODE
            Exit Function

        Case WM_NOTIFY
            Dim dwRtn As LongPtr
            Static nmh As NMHDR
    
      
            If (wParam = IDD_LISTVIEW) Then
                dwRtn = DoGVNotify(hWnd, lParam)
            End If
            If dwRtn Then
              FGVWndProc = dwRtn
              Exit Function
            End If
    


        Case WM_DESTROY
            Call UnSubclass2(hWnd, PtrFGVWndProc(), uIdSubclass)
    End Select
    FGVWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Public Function PtrFGVWndProc() As LongPtr
    PtrFGVWndProc = FARPROC(AddressOf FGVWndProc)
End Function

Public Function LVGWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    Dim sText As String
    Select Case uMsg
        Case WM_DESTROY
            Call UnSubclass2(hWnd, PtrLVGWndProc, uIdSubclass)
    End Select
    LVGWndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Public Function PtrLVGWndProc() As LongPtr
    PtrLVGWndProc = FARPROC(AddressOf LVGWndProc)
End Function

Public Function DoGVNotify(hWnd As LongPtr, lParam As LongPtr) As LongPtr
    Dim sText As String, sSubText As String
    Dim tNMH As NMHDR
    CopyMemory tNMH, ByVal lParam, Len(tNMH)

    Select Case tNMH.code
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
                            .iImage = VLItems(.iItem).iImage
                    End Select
                End If
                CopyMemory ByVal lParam, ByVal VarPtr(LVDI), LenB(LVDI)
            End With
        End Select
    Exit Function
End Function

Public Function FARPROC(pfn As LongPtr) As LongPtr
    FARPROC = pfn
End Function

Private Sub Str2WCHAR(sz As String, iOut() As Integer)
    Dim i As Long
    ReDim iOut(255)
    'If Len(sz) > MAX_PATH Then sz = Left$(sz, MAX_PATH)
    For i = 1 To Len(sz)
        iOut(i - 1) = AscW(Mid(sz, i, 1))
    Next i
End Sub


