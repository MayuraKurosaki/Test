Attribute VB_Name = "CommonModule"
Option Explicit

'--------------Constants----------------
'Window Messages
Public Const WM_SETFOCUS                    As Long = &H7
Public Const WM_SETREDRAW                   As Long = &HB
Public Const WM_NOTIFY                      As Long = &H4E
Public Const WM_MOUSEWHEEL                  As Long = &H20A
Public Const WM_MOUSEACTIVATE               As Long = &H21
Public Const WM_SETFONT                     As Long = &H30
Public Const WM_NCDESTROY                   As Long = &H82
Public Const WM_HSCROLL                     As Long = &H114
Public Const WM_VSCROLL                     As Long = &H115

' Private Window Messages Start Here:
Public Const WM_USER                        As Long = &H400

'Common control shared messages
Public Const CCM_FIRST                      As Long = &H2000
Public Const CCM_SETBKCOLOR                 As Long = (CCM_FIRST + 1)       'lParam is bkColor
Public Const CCM_SETCOLORSCHEME             As Long = (CCM_FIRST + 2)       'lParam is color scheme
Public Const CCM_GETCOLORSCHEME             As Long = (CCM_FIRST + 3)       'fills in COLORSCHEME pointed to by lParam
Public Const CCM_GETDROPTARGET              As Long = (CCM_FIRST + 4)
Public Const CCM_SETUNICODEFORMAT           As Long = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT           As Long = (CCM_FIRST + 6)
Public Const CCM_SETVERSION                 As Long = (CCM_FIRST + 7)
Public Const CCM_GETVERSION                 As Long = (CCM_FIRST + 8)
Public Const CCM_SETNOTIFYWINDOW            As Long = (CCM_FIRST + 9)       'wParam == hwndParent.
Public Const CCM_SETWINDOWTHEME             As Long = (CCM_FIRST + &HB)
Public Const CCM_DPISCALE                   As Long = (CCM_FIRST + &HC)     'wParam == Awareness

'Window Styles
Public Const WS_OVERLAPPED                  As Long = &H0
Public Const WS_POPUP                       As Long = &H80000000
Public Const WS_CHILD                       As Long = &H40000000
Public Const WS_MINIMIZE                    As Long = &H20000000
Public Const WS_VISIBLE                     As Long = &H10000000
Public Const WS_DISABLED                    As Long = &H8000000
Public Const WS_CLIPSIBLINGS                As Long = &H4000000
Public Const WS_CLIPCHILDREN                As Long = &H2000000
Public Const WS_MAXIMIZE                    As Long = &H1000000
Public Const WS_CAPTION                     As Long = &HC00000       'WS_BORDER | WS_DLGFRAME
Public Const WS_BORDER                      As Long = &H800000
Public Const WS_DLGFRAME                    As Long = &H400000
Public Const WS_VSCROLL                     As Long = &H200000
Public Const WS_HSCROLL                     As Long = &H100000
Public Const WS_SYSMENU                     As Long = &H80000
Public Const WS_THICKFRAME                  As Long = &H40000
Public Const WS_GROUP                       As Long = &H20000
Public Const WS_TABSTOP                     As Long = &H10000

Public Const WS_MINIMIZEBOX                 As Long = &H20000
Public Const WS_MAXIMIZEBOX                 As Long = &H10000

Public Const WS_TILED                       As Long = WS_OVERLAPPED
Public Const WS_ICONIC                      As Long = WS_MINIMIZE
Public Const WS_SIZEBOX                     As Long = WS_THICKFRAME

'Common Window Styles
Public Const WS_OVERLAPPEDWINDOW            As Long = WS_OVERLAPPED Or _
                                                      WS_CAPTION Or _
                                                      WS_SYSMENU Or _
                                                      WS_THICKFRAME Or _
                                                      WS_MINIMIZEBOX Or _
                                                      WS_MAXIMIZEBOX

Public Const WS_POPUPWINDOW                 As Long = WS_POPUP Or _
                                                      WS_BORDER Or _
                                                      WS_SYSMENU

Public Const WS_CHILDWINDOW                 As Long = WS_CHILD

Public Const WS_TILEDWINDOW                 As Long = WS_OVERLAPPEDWINDOW

'Extended Window Styles
Public Const WS_EX_DLGMODALFRAME            As Long = &H1
Public Const WS_EX_NOPARENTNOTIFY           As Long = &H4
Public Const WS_EX_TOPMOST                  As Long = &H8
Public Const WS_EX_ACCEPTFILES              As Long = &H10
Public Const WS_EX_TRANSPARENT              As Long = &H20
Public Const WS_EX_MDICHILD                 As Long = &H40
Public Const WS_EX_TOOLWINDOW               As Long = &H80
Public Const WS_EX_WINDOWEDGE               As Long = &H100
Public Const WS_EX_CLIENTEDGE               As Long = &H200
Public Const WS_EX_CONTEXTHELP              As Long = &H400

Public Const WS_EX_RIGHT                    As Long = &H1000
Public Const WS_EX_LEFT                     As Long = &H0
Public Const WS_EX_RTLREADING               As Long = &H2000
Public Const WS_EX_LTRREADING               As Long = &H0
Public Const WS_EX_LEFTSCROLLBAR            As Long = &H4000
Public Const WS_EX_RIGHTSCROLLBAR           As Long = &H0

Public Const WS_EX_CONTROLPARENT            As Long = &H10000
Public Const WS_EX_STATICEDGE               As Long = &H20000
Public Const WS_EX_APPWINDOW                As Long = &H40000

Public Const WS_EX_OVERLAPPEDWINDOW         As Long = WS_EX_WINDOWEDGE Or _
                                                      WS_EX_CLIENTEDGE
Public Const WS_EX_PALETTEWINDOW            As Long = WS_EX_WINDOWEDGE Or _
                                                      WS_EX_TOOLWINDOW Or _
                                                      WS_EX_TOPMOST

Public Const WS_EX_LAYERED                  As Long = &H80000
Public Const WS_EX_NOINHERITLAYOUT          As Long = &H100000      'Disable inheritence of mirroring by children
Public Const WS_EX_NOREDIRECTIONBITMAP      As Long = &H200000
Public Const WS_EX_LAYOUTRTL                As Long = &H400000      'Right to left mirroring
Public Const WS_EX_COMPOSITED               As Long = &H2000000
Public Const WS_EX_NOACTIVATE               As Long = &H8000000

'WM_MOUSEACTIVATE Return Codes
Public Const MA_ACTIVATE                    As Long = &H1
Public Const MA_ACTIVATEANDEAT              As Long = &H2
Public Const MA_NOACTIVATE                  As Long = &H3
Public Const MA_NOACTIVATEANDEAT            As Long = &H4


Public Const WHEEL_DELTA                    As Long = 120

Public Const NM_CLICK                       As Long = -2
Public Const NM_DBLCLK                      As Long = -3

' Global Memory Flags
Public Const GMEM_FIXED                     As Long = &H0
Public Const GMEM_MOVEABLE                  As Long = &H2
Public Const GMEM_ZEROINIT                  As Long = &H40
Public Const GHND                           As Long = GMEM_MOVEABLE Or _
                                                      GMEM_ZEROINIT
Public Const GPTR                           As Long = GMEM_FIXED Or _
                                                      GMEM_ZEROINIT

' Device Parameters for GetDeviceCaps()
Public Const DRIVERVERSION                  As Long = 0     '  Device driver version
Public Const TECHNOLOGY                     As Long = 2     '  Device classification
Public Const HORZSIZE                       As Long = 4     '  Horizontal size in millimeters
Public Const VERTSIZE                       As Long = 6     '  Vertical size in millimeters
Public Const HORZRES                        As Long = 8     '  Horizontal width in pixels
Public Const VERTRES                        As Long = 10    '  Vertical width in pixels
Public Const BITSPIXEL                      As Long = 12    '  Number of bits per pixel
Public Const PLANES                         As Long = 14    '  Number of planes
Public Const NUMBRUSHES                     As Long = 16    '  Number of brushes the device has
Public Const NUMPENS                        As Long = 18    '  Number of pens the device has
Public Const NUMMARKERS                     As Long = 20    '  Number of markers the device has
Public Const NUMFONTS                       As Long = 22    '  Number of fonts the device has
Public Const NUMCOLORS                      As Long = 24    '  Number of colors the device supports
Public Const PDEVICESIZE                    As Long = 26    '  Size required for device descriptor
Public Const CURVECAPS                      As Long = 28    '  Curve capabilities
Public Const LINECAPS                       As Long = 30    '  Line capabilities
Public Const POLYGONALCAPS                  As Long = 32    '  Polygonal capabilities
Public Const TEXTCAPS                       As Long = 34    '  Text capabilities
Public Const CLIPCAPS                       As Long = 36    '  Clipping capabilities
Public Const RASTERCAPS                     As Long = 38    '  Bitblt capabilities
Public Const ASPECTX                        As Long = 40    '  Length of the X leg
Public Const ASPECTY                        As Long = 42    '  Length of the Y leg
Public Const ASPECTXY                       As Long = 44    '  Length of the hypotenuse

Public Const LOGPIXELSX                     As Long = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY                     As Long = 90    '  Logical pixels/inch in Y

Public Const SIZEPALETTE                    As Long = 104   '  Number of entries in physical palette
Public Const NUMRESERVED                    As Long = 106   '  Number of reserved entries in palette
Public Const COLORRES                       As Long = 108   '  Actual color resolution


'' Font Families
''
'Public Const FF_DONTCARE = 0    '  Don't care or don't know.
'Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.
'
'' Times Roman, Century Schoolbook, etc.
'Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.
'
'' Helvetica, Swiss, etc.
'Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.
'
'' Pica, Elite, Courier, etc.
'Public Const FF_SCRIPT = 64     '  Cursive, etc.
'Public Const FF_DECORATIVE = 80 '  Old English, etc.

''/* Font Weights */
'Public Const FW_DONTCARE As Long = 0
'Public Const FW_THIN  As Long = 100
'Public Const FW_EXTRALIGHT  As Long = 200
'Public Const FW_LIGHT  As Long = 300
'Public Const FW_NORMAL As Long = 400
'Public Const FW_MEDIUM  As Long = 500
'Public Const FW_SEMIBOLD As Long = 600
'Public Const FW_BOLD  As Long = 700
'Public Const FW_EXTRABOLD  As Long = 800
'Public Const FW_HEAVY  As Long = 900
'
'Public Const FW_ULTRALIGHT  As Long = FW_EXTRALIGHT
'Public Const FW_REGULAR  As Long = FW_NORMAL
'Public Const FW_DEMIBOLD  As Long = FW_SEMIBOLD
'Public Const FW_ULTRABOLD  As Long = FW_EXTRABOLD
'Public Const FW_BLACK  As Long = FW_HEAVY
'
'Public Const OUT_DEFAULT_PRECIS  As Long = 0
'Public Const OUT_STRING_PRECIS  As Long = 1
'Public Const OUT_CHARACTER_PRECIS  As Long = 2
'Public Const OUT_STROKE_PRECIS  As Long = 3
'Public Const OUT_TT_PRECIS  As Long = 4
'Public Const OUT_DEVICE_PRECIS  As Long = 5
'Public Const OUT_RASTER_PRECIS  As Long = 6
'Public Const OUT_TT_ONLY_PRECIS  As Long = 7
'Public Const OUT_OUTLINE_PRECIS  As Long = 8
'Public Const OUT_SCREEN_OUTLINE_PRECIS  As Long = 9
'Public Const OUT_PS_ONLY_PRECIS  As Long = 10
'
'Public Const CLIP_DEFAULT_PRECIS  As Long = 0
'Public Const CLIP_CHARACTER_PRECIS  As Long = 1
'Public Const CLIP_STROKE_PRECIS  As Long = 2
'Public Const CLIP_MASK  As Long = &HF
'Public Const CLIP_LH_ANGLES  As Long = &H10
'Public Const CLIP_TT_ALWAYS  As Long = &H20
'Public Const CLIP_DFA_DISABLE  As Long = &H40
'Public Const CLIP_EMBEDDED  As Long = &H80
'
'Public Const DEFAULT_QUALITY  As Long = 0
'Public Const DRAFT_QUALITY  As Long = 1
'Public Const PROOF_QUALITY  As Long = 2
'Public Const NONANTIALIASED_QUALITY  As Long = 3
'Public Const ANTIALIASED_QUALITY  As Long = 4
'Public Const CLEARTYPE_QUALITY  As Long = 5
'Public Const CLEARTYPE_NATURAL_QUALITY  As Long = 6
'
'Public Const DEFAULT_PITCH  As Long = 0
'Public Const FIXED_PITCH  As Long = 1
'Public Const VARIABLE_PITCH  As Long = 2
'Public Const MONO_FONT  As Long = 8
'
'Public Const ANSI_CHARSET  As Long = 0
'Public Const DEFAULT_CHARSET  As Long = 1
'Public Const SYMBOL_CHARSET  As Long = 2
'Public Const SHIFTJIS_CHARSET  As Long = 128
'Public Const HANGEUL_CHARSET  As Long = 129
'Public Const HANGUL_CHARSET  As Long = 129
'Public Const GB2312_CHARSET  As Long = 134
'Public Const CHINESEBIG5_CHARSET  As Long = 136
'Public Const OEM_CHARSET  As Long = 255
'Public Const JOHAB_CHARSET  As Long = 130
'Public Const HEBREW_CHARSET  As Long = 177
'Public Const ARABIC_CHARSET  As Long = 178
'Public Const GREEK_CHARSET  As Long = 161
'Public Const TURKISH_CHARSET  As Long = 162
'Public Const VIETNAMESE_CHARSET  As Long = 163
'Public Const THAI_CHARSET  As Long = 222
'Public Const EASTEUROPE_CHARSET  As Long = 238
'Public Const RUSSIAN_CHARSET  As Long = 204
'
'Public Const MAC_CHARSET  As Long = 77
'Public Const BALTIC_CHARSET  As Long = 186

'' Logical Font
'Public Const LF_FACESIZE = 32
'Public Const LF_FULLFACESIZE = 64

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

'Public Enum ListPictureAlignmentConstants
'    lvwTopLeft = 0
'    lvwTopRight = 1
'    lvwBottomLeft = 2
'    lvwBottomRight = 3
'    lvwCenter = 4
'    lvwTile = 5
'End Enum
'
'Public Enum ListSortOrderConstants
'    lvwAscending = 1
'    lvwDescending = 2
'End Enum
'
'Public Enum ListTextBackgroundConstants
'    lvwTransparent = 0
'    lvwOpaque = 1
'End Enum
'
'Public Enum ListViewConstants
'    lvwIcon = 0
'    lvwSmallIcon = 1
'    lvwList = 2
'    lvwReport = 3
'End Enum
'
'Public Enum LISTVIEWITEMRECT
'    LVIR_BOUNDS = 0
'    LVIR_ICON = 1
'    LVIR_LABEL = 2
'    LVIR_SELECTBOUNDS = 3
'End Enum
'
'Public Enum LVHT_FLAGS
'     LVHT_NOWHERE = &H1   ' in LV client area, but not over item
'     LVHT_ONITEMICON = &H2
'     LVHT_ONITEMLABEL = &H4
'     LVHT_ONITEMSTATEICON = &H8
'     LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
'
'    '  ' outside the LV's client area
'     LVHT_ABOVE = &H8
'     LVHT_BELOW = &H10
'     LVHT_TORIGHT = &H20
'     LVHT_TOLEFT = &H40
'#If (WIN32_IE >= &H600) Then
'    LVHT_EX_GROUP_HEADER = &H10000000
'    LVHT_EX_GROUP_FOOTER = &H20000000
'    LVHT_EX_GROUP_COLLAPSE = &H40000000
'    LVHT_EX_GROUP_BACKGROUND = &H80000000
'    LVHT_EX_GROUP_STATEICON = &H1000000
'    LVHT_EX_GROUP_SUBSETLINK = &H2000000
'    LVHT_EX_GROUP = (LVHT_EX_GROUP_BACKGROUND Or LVHT_EX_GROUP_COLLAPSE Or LVHT_EX_GROUP_FOOTER Or LVHT_EX_GROUP_HEADER Or LVHT_EX_GROUP_STATEICON Or LVHT_EX_GROUP_SUBSETLINK)
'    LVHT_EX_ONCONTENTS = &H4000000          'On item AND not on the background
'    LVHT_EX_FOOTER = &H8000000
'#End If
'End Enum



'--------------Type Definitions----------------
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type InitCommonControlsExType
    dwSize As Long
    dwICC As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type POINTF
    x As Single
    y As Single
End Type

Public Type NMHDR
    hWndFrom As LongPtr   ' Window handle of control sending message
    IDFrom As LongPtr        ' Identifier of control sending message
    Code  As Long         ' Specifies the notification code
End Type
'typedef struct tagNMHDR
'{
'    HWND      hwndFrom;
'    UINT_PTR  idFrom;
'    UINT      code;         // NM_ code
'}   NMHDR;



'--------------API Declarations----------------
Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" _
         (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, _
         ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, _
         ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long

'Public Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

'Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Public Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Public Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare PtrSafe Function InitCommonControlsEx Lib "COMCTL32" (ByRef LPINITCOMMONCONTROLSEX As InitCommonControlsExType) As Long
Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Public Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr

Public Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Public Declare PtrSafe Function SysAllocString Lib "OleAut32.dll" (ByVal psz As LongPtr) As LongPtr

Public Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
Public Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As LongPtr) As LongPtr
Public Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONTA) As LongPtr

Public Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Public Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, lprcUpdate As RECT, ByVal hrgnUpdate As LongPtr, ByVal Flags As Long) As Long
Public Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hwnd As LongPtr) As Long

'Public Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr

Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalHandle Lib "kernel32" (wMem As Any) As LongPtr

' 文字列の幅(ピクセル)を取得する関数
Public Function GetStringWidthPixel(ByVal Text As String, ByVal fontName As String, ByVal fontSize As Long) As Long
    Dim hDC As LongPtr, hFont As LongPtr, hOldFont As LongPtr
    Dim sz As SIZE
    
    hDC = GetDC(0) ' デスクトップのDCを取得
    
    ' フォントの作成（※簡易的な設定）
'    hFont = CreateFont(-fontSize * 1.33, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, fontName)
    hFont = CreateFont(fontSize, 0, 0, 0, 0, 0, 0, 0, SHIFTJIS_CHARSET, 0, 0, 0, 0, fontName)
    hOldFont = SelectObject(hDC, hFont)
    
    ' 幅を取得
    Call GetTextExtentPoint32(hDC, Text, Len(Text), sz)
    
    ' 後処理
    Call SelectObject(hDC, hOldFont)
    Call DeleteObject(hFont)
    Call ReleaseDC(0, hDC)
    
    GetStringWidthPixel = sz.cx
End Function

'#define MAKEWORD(a, b)      ((WORD)(((BYTE)(((DWORD_PTR)(a)) & 0xff)) | ((WORD)((BYTE)(((DWORD_PTR)(b)) & 0xff))) << 8))

'#define MAKELONG(a, b)      ((LONG)(((WORD)(((DWORD_PTR)(a)) & 0xffff)) | ((DWORD)((WORD)(((DWORD_PTR)(b)) & 0xffff))) << 16))
Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
End Function

'#define MAKELPARAM(l, h)      ((LPARAM)(DWORD)MAKELONG(l, h))
Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
    MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

'#define MAKEWPARAM(l, h)      ((WPARAM)(DWORD)MAKELONG(l, h))

'#define MAKELRESULT(l, h)     ((LRESULT)(DWORD)MAKELONG(l, h))

'#define LOWORD(l)           ((WORD)(((DWORD_PTR)(l)) & 0xffff))
Public Function LOWORD(ByVal dwValue As Long) As Integer
' Returns the low 16-bit integer from a 32-bit long integer
    CopyMemory LOWORD, dwValue, 2&
End Function

'#define HIWORD(l)           ((WORD)((((DWORD_PTR)(l)) >> 16) & 0xffff))

'#define LOBYTE(w)           ((BYTE)(((DWORD_PTR)(w)) & 0xff))

'#define HIBYTE(w)           ((BYTE)((((DWORD_PTR)(w)) >> 8) & 0xff))
