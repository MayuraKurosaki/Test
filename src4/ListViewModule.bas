Attribute VB_Name = "ListViewModule"
'Option Explicit
'
'API宣言部（標準モジュール）modAPI
Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_EX_CLIENTEDGE As Long = &H200

Public Const WM_SETFOCUS As Long = 7
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_MOUSEACTIVATE As Long = &H21
Public Const WM_NCDESTROY As Long = &H82
Public Const WM_SETFONT As Long = &H30

Public Const MA_NOACTIVATE As Long = 3

Public Const LVS_REPORT As Long = 1
Public Const LVS_SHOWSELALWAYS As Long = 8
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVS_EX_GRIDLINES As Long = 1
Public Const LVS_EX_CHECKBOXES As Long = 4

'Window Messages
Public Const LVM_FIRST                      As Long = &H1000
Public Const LVM_GETBKCOLOR                 As Long = (LVM_FIRST + 0)
Public Const LVM_SETBKCOLOR                 As Long = (LVM_FIRST + 1)
Public Const LVM_GETIMAGELIST               As Long = (LVM_FIRST + 2)
Public Const LVM_SETIMAGELIST               As Long = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT               As Long = (LVM_FIRST + 4)
Public Const LVM_GETITEM                    As Long = (LVM_FIRST + 5)
Public Const LVM_SETITEM                    As Long = (LVM_FIRST + 6)
Public Const LVM_INSERTITEM                 As Long = (LVM_FIRST + 7)
Public Const LVM_DELETEITEM                 As Long = (LVM_FIRST + 8)
Public Const LVM_DELETEALLITEMS             As Long = (LVM_FIRST + 9)
Public Const LVM_GETCALLBACKMASK            As Long = (LVM_FIRST + 10)
Public Const LVM_SETCALLBACKMASK            As Long = (LVM_FIRST + 11)
Public Const LVM_GETNEXTITEM                As Long = (LVM_FIRST + 12)
Public Const LVM_FINDITEM                   As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMRECT                As Long = (LVM_FIRST + 14)
Public Const LVM_SETITEMPOSITION            As Long = (LVM_FIRST + 15)
Public Const LVM_GETITEMPOSITION            As Long = (LVM_FIRST + 16)
Public Const LVM_GETSTRINGWIDTH             As Long = (LVM_FIRST + 17)
Public Const LVM_HITTEST                    As Long = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE              As Long = (LVM_FIRST + 19)
Public Const LVM_SCROLL                     As Long = (LVM_FIRST + 20)
Public Const LVM_REDRAWITEMS                As Long = (LVM_FIRST + 21)
Public Const LVM_ARRANGE                    As Long = (LVM_FIRST + 22)
Public Const LVM_EDITLABEL                  As Long = (LVM_FIRST + 23)
Public Const LVM_GETEDITCONTROL             As Long = (LVM_FIRST + 24)
Public Const LVM_GETCOLUMN                  As Long = (LVM_FIRST + 25)
Public Const LVM_SETCOLUMN                  As Long = (LVM_FIRST + 26)
Public Const LVM_INSERTCOLUMN               As Long = (LVM_FIRST + 27)
Public Const LVM_DELETECOLUMN               As Long = (LVM_FIRST + 28)
Public Const LVM_GETCOLUMNWIDTH             As Long = (LVM_FIRST + 29)
Public Const LVM_SETCOLUMNWIDTH             As Long = (LVM_FIRST + 30)
Public Const LVM_GETHEADER                  As Long = (LVM_FIRST + 31)
Public Const LVM_CREATEDRAGIMAGE            As Long = (LVM_FIRST + 33)
Public Const LVM_GETVIEWRECT                As Long = (LVM_FIRST + 34)
Public Const LVM_GETTEXTCOLOR               As Long = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR               As Long = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR             As Long = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR             As Long = (LVM_FIRST + 38)
Public Const LVM_GETTOPINDEX                As Long = (LVM_FIRST + 39)
Public Const LVM_GETCOUNTPERPAGE            As Long = (LVM_FIRST + 40)
Public Const LVM_GETORIGIN                  As Long = (LVM_FIRST + 41)
Public Const LVM_UPDATE                     As Long = (LVM_FIRST + 42)
Public Const LVM_SETITEMSTATE               As Long = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE               As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT                As Long = (LVM_FIRST + 45)
Public Const LVM_SETITEMTEXT                As Long = (LVM_FIRST + 46)
Public Const LVM_SETITEMCOUNT               As Long = (LVM_FIRST + 47)
Public Const LVM_SORTITEMS                  As Long = (LVM_FIRST + 48)
Public Const LVM_SETITEMPOSITION32          As Long = (LVM_FIRST + 49)
Public Const LVM_GETSELECTEDCOUNT           As Long = (LVM_FIRST + 50)
Public Const LVM_GETITEMSPACING             As Long = (LVM_FIRST + 51)
Public Const LVM_GETISEARCHSTRING           As Long = (LVM_FIRST + 52)
Public Const LVM_SETICONSPACING             As Long = (LVM_FIRST + 53)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE   As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE   As Long = (LVM_FIRST + 55)
Public Const LVM_GETSUBITEMRECT             As Long = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST             As Long = (LVM_FIRST + 57)
Public Const LVM_SETCOLUMNORDERARRAY        As Long = (LVM_FIRST + 58)
Public Const LVM_GETCOLUMNORDERARRAY        As Long = (LVM_FIRST + 59)
Public Const LVM_SETHOTITEM                 As Long = (LVM_FIRST + 60)
Public Const LVM_GETHOTITEM                 As Long = (LVM_FIRST + 61)
Public Const LVM_SETHOTCURSOR               As Long = (LVM_FIRST + 62)
Public Const LVM_GETHOTCURSOR               As Long = (LVM_FIRST + 63)
Public Const LVM_APPROXIMATEVIEWRECT        As Long = (LVM_FIRST + 64)
Public Const LVM_SETWORKAREAS               As Long = (LVM_FIRST + 65)
Public Const LVM_GETSELECTIONMARK           As Long = (LVM_FIRST + 66)
Public Const LVM_SETSELECTIONMARK           As Long = (LVM_FIRST + 67)
Public Const LVM_SETBKIMAGE                 As Long = (LVM_FIRST + 68)
Public Const LVM_GETBKIMAGE                 As Long = (LVM_FIRST + 69)
Public Const LVM_GETWORKAREAS               As Long = (LVM_FIRST + 70)
Public Const LVM_SETHOVERTIME               As Long = (LVM_FIRST + 71)
Public Const LVM_GETHOVERTIME               As Long = (LVM_FIRST + 72)
Public Const LVM_GETNUMBEROFWORKAREAS       As Long = (LVM_FIRST + 73)
Public Const LVM_SETTOOLTIPS                As Long = (LVM_FIRST + 74)
Public Const LVM_GETITEMW                   As Long = (LVM_FIRST + 75)
Public Const LVM_SETITEMW                   As Long = (LVM_FIRST + 76)  'Unicode
Public Const LVM_INSERTITEMW                As Long = (LVM_FIRST + 77) 'Unicode
Public Const LVM_GETTOOLTIPS                As Long = (LVM_FIRST + 78)
Public Const LVM_GETHOTLIGHTCOLOR           As Long = (LVM_FIRST + 79) 'UNDOCUMENTED
Public Const LVM_SETHOTLIGHTCOLOR           As Long = (LVM_FIRST + 80) 'UNDOCUMENTED
Public Const LVM_SORTITEMSEX                As Long = (LVM_FIRST + 81)
Public Const LVM_SETRANGEOBJECT             As Long = (LVM_FIRST + 82) 'UNDOCUMENTED
Public Const LVM_FINDITEMW                  As Long = (LVM_FIRST + 83) 'Unicode
Public Const LVM_RESETEMPTYTEXT             As Long = (LVM_FIRST + 84) 'UNDOCUMENTED
Public Const LVM_SETFROZENITEM              As Long = (LVM_FIRST + 85) 'UNDOCUMENTED
Public Const LVM_GETFROZENITEM              As Long = (LVM_FIRST + 86) 'UNDOCUMENTED
Public Const LVM_GETSTRINGWIDTHW            As Long = (LVM_FIRST + 87)
Public Const LVM_SETFROZENSLOT              As Long = (LVM_FIRST + 88) 'UNDOCUMENTED
Public Const LVM_GETFROZENSLOT              As Long = (LVM_FIRST + 89) 'UNDOCUMENTED
Public Const LVM_SETVIEWMARGIN              As Long = (LVM_FIRST + 90) 'UNDOCUMENTED
Public Const LVM_GETVIEWMARGIN              As Long = (LVM_FIRST + 91) 'UNDOCUMENTED
Public Const LVM_GETGROUPSTATE              As Long = (LVM_FIRST + 92)
Public Const LVM_GETFOCUSEDGROUP            As Long = (LVM_FIRST + 93)
Public Const LVM_EDITGROUPLABEL             As Long = (LVM_FIRST + 94) 'UNDOCUMENTED
Public Const LVM_GETCOLUMNW                 As Long = (LVM_FIRST + 95) 'Unicode
Public Const LVM_SETCOLUMNW                 As Long = (LVM_FIRST + 96) 'Unicode
Public Const LVM_INSERTCOLUMNW              As Long = (LVM_FIRST + 97) 'Unicode
Public Const LVM_GETGROUPRECT               As Long = (LVM_FIRST + 98)

Public Const LVM_GETITEMTEXTW               As Long = (LVM_FIRST + 115)     'Unicode
Public Const LVM_SETITEMTEXTW               As Long = (LVM_FIRST + 116)           'Unicode
Public Const LVM_GETISEARCHSTRINGW          As Long = (LVM_FIRST + 117)
Public Const LVM_EDITLABELW                 As Long = (LVM_FIRST + 118)

Public Const LVM_SETBKIMAGEW                As Long = (LVM_FIRST + 138)
Public Const LVM_GETBKIMAGEW                As Long = (LVM_FIRST + 139)
Public Const LVM_SETSELECTEDCOLUMN          As Long = (LVM_FIRST + 140)
Public Const LVM_SETTILEWIDTH               As Long = (LVM_FIRST + 141)
Public Const LVM_SETVIEW                    As Long = (LVM_FIRST + 142)
Public Const LVM_GETVIEW                    As Long = (LVM_FIRST + 143)

Public Const LVM_INSERTGROUP                As Long = (LVM_FIRST + 145)

Public Const LVM_SETGROUPINFO               As Long = (LVM_FIRST + 147)

Public Const LVM_GETGROUPINFO               As Long = (LVM_FIRST + 149)
Public Const LVM_REMOVEGROUP                As Long = (LVM_FIRST + 150)
Public Const LVM_MOVEGROUP                  As Long = (LVM_FIRST + 151)
Public Const LVM_GETGROUPCOUNT              As Long = (LVM_FIRST + 152)
Public Const LVM_GETGROUPINFOBYINDEX        As Long = (LVM_FIRST + 153)
Public Const LVM_MOVEITEMTOGROUP            As Long = (LVM_FIRST + 154)
Public Const LVM_SETGROUPMETRICS            As Long = (LVM_FIRST + 155)
Public Const LVM_GETGROUPMETRICS            As Long = (LVM_FIRST + 156)
Public Const LVM_ENABLEGROUPVIEW            As Long = (LVM_FIRST + 157)
Public Const LVM_SORTGROUPS                 As Long = (LVM_FIRST + 158)
Public Const LVM_INSERTGROUPSORTED          As Long = (LVM_FIRST + 159)
Public Const LVM_REMOVEALLGROUPS            As Long = (LVM_FIRST + 160)
Public Const LVM_HASGROUP                   As Long = (LVM_FIRST + 161)
Public Const LVM_SETTILEVIEWINFO            As Long = (LVM_FIRST + 162)
Public Const LVM_GETTILEVIEWINFO            As Long = (LVM_FIRST + 163)
Public Const LVM_SETTILEINFO                As Long = (LVM_FIRST + 164)
Public Const LVM_GETTILEINFO                As Long = (LVM_FIRST + 165)
Public Const LVM_SETINSERTMARK              As Long = (LVM_FIRST + 166)
Public Const LVM_GETINSERTMARK              As Long = (LVM_FIRST + 167)
Public Const LVM_INSERTMARKHITTEST          As Long = (LVM_FIRST + 168)
Public Const LVM_GETINSERTMARKRECT          As Long = (LVM_FIRST + 169)
Public Const LVM_SETINSERTMARKCOLOR         As Long = (LVM_FIRST + 170)
Public Const LVM_GETINSERTMARKCOLOR         As Long = (LVM_FIRST + 171)

Public Const LVM_SETINFOTIP                 As Long = (LVM_FIRST + 173)
Public Const LVM_GETSELECTEDCOLUMN          As Long = (LVM_FIRST + 174)
Public Const LVM_ISGROUPVIEWENABLED         As Long = (LVM_FIRST + 175)
Public Const LVM_GETOUTLINECOLOR            As Long = (LVM_FIRST + 176)
Public Const LVM_SETOUTLINECOLOR            As Long = (LVM_FIRST + 177)
Public Const LVM_SETKEYBOARDSELECTED        As Long = (LVM_FIRST + 178)  'UNDOCUMENTED
Public Const LVM_CANCELEDITLABEL            As Long = (LVM_FIRST + 179)
Public Const LVM_MAPINDEXTOID               As Long = (LVM_FIRST + 180)
Public Const LVM_MAPIDTOINDEX               As Long = (LVM_FIRST + 181)
Public Const LVM_ISITEMVISIBLE              As Long = (LVM_FIRST + 182)
Public Const LVM_EDITSUBITEM                As Long = (LVM_FIRST + 183)          'UNDOCUMENTED
Public Const LVM_ENSURESUBITEMVISIBLE       As Long = (LVM_FIRST + 184) 'UNDOCUMENTED
Public Const LVM_GETCLIENTRECT              As Long = (LVM_FIRST + 185)        'UNDOCUMENTED
Public Const LVM_GETFOCUSEDCOLUMN           As Long = (LVM_FIRST + 186)     'UNDOCUMENTED
Public Const LVM_SETOWNERDATACALLBACK       As Long = (LVM_FIRST + 187) 'UNDOCUMENTED
Public Const LVM_RECOMPUTEITEMS             As Long = (LVM_FIRST + 188)      'UNDOCUMENTED
Public Const LVM_QUERYINTERFACE             As Long = (LVM_FIRST + 189)      'UNDOCUMENTED: NOT OFFICIAL NAME
Public Const LVM_SETGROUPSUBSETCOUNT        As Long = (LVM_FIRST + 190) 'UNDOCUMENTED
Public Const LVM_GETGROUPSUBSETCOUNT        As Long = (LVM_FIRST + 191) 'UNDOCUMENTED
Public Const LVM_ORDERTOINDEX               As Long = (LVM_FIRST + 192)        'UNDOCUMENTED
Public Const LVM_GETACCVERSION              As Long = (LVM_FIRST + 193)       'UNDOCUMENTED
Public Const LVM_MAPACCIDTOACCINDEX         As Long = (LVM_FIRST + 194)  'UNDOCUMENTED
Public Const LVM_MAPACCINDEXTOACCID         As Long = (LVM_FIRST + 195)  'UNDOCUMENTED
Public Const LVM_GETOBJECTCOUNT             As Long = (LVM_FIRST + 196)      'UNDOCUMENTED
Public Const LVM_GETOBJECTRECT              As Long = (LVM_FIRST + 197)       'UNDOCUMENTED
Public Const LVM_ACCHITTEST                 As Long = (LVM_FIRST + 198)          'UNDOCUMENTED
Public Const LVM_GETFOCUSEDOBJECT           As Long = (LVM_FIRST + 199)    'UNDOCUMENTED
Public Const LVM_GETOBJECTROLE              As Long = (LVM_FIRST + 200)       'UNDOCUMENTED
Public Const LVM_GETOBJECTSTATE             As Long = (LVM_FIRST + 201)      'UNDOCUMENTED
Public Const LVM_ACCNAVIGATE                As Long = (LVM_FIRST + 202)         'UNDOCUMENTED
Public Const LVM_INVOKEDEFAULTACTION        As Long = (LVM_FIRST + 203) 'UNDOCUMENTED
Public Const LVM_GETEMPTYTEXT               As Long = (LVM_FIRST + 204)
Public Const LVM_GETFOOTERRECT              As Long = (LVM_FIRST + 205)
Public Const LVM_GETFOOTERINFO              As Long = (LVM_FIRST + 206)
Public Const LVM_GETFOOTERITEMRECT          As Long = (LVM_FIRST + 207)
Public Const LVM_GETFOOTERITEM              As Long = (LVM_FIRST + 208)
Public Const LVM_GETITEMINDEXRECT           As Long = (LVM_FIRST + 209)
Public Const LVM_SETITEMINDEXSTATE          As Long = (LVM_FIRST + 210)
Public Const LVM_GETNEXTITEMINDEX           As Long = (LVM_FIRST + 211)
Public Const LVM_SETPRESERVEALPHA           As Long = (LVM_FIRST + 212)    'UNDOCUMENTED

Public Const CCM_FIRST                      As Long = &H2000       '// Common control shared messages
Public Const CCM_SETBKCOLOR                 As Long = (CCM_FIRST + 1)      '// lParam is bkColor
Public Const CCM_SETCOLORSCHEME             As Long = (CCM_FIRST + 2)      '// lParam is color scheme
Public Const CCM_GETCOLORSCHEME             As Long = (CCM_FIRST + 3)      '// fills in COLORSCHEME pointed to by lParam
Public Const CCM_GETDROPTARGET              As Long = (CCM_FIRST + 4)
Public Const CCM_SETUNICODEFORMAT           As Long = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT           As Long = (CCM_FIRST + 6)
Public Const CCM_SETVERSION                 As Long = (CCM_FIRST + 7)
Public Const CCM_GETVERSION                 As Long = (CCM_FIRST + 8)
Public Const CCM_SETNOTIFYWINDOW            As Long = (CCM_FIRST + 9)      '// wParam == hwndParent.
Public Const CCM_SETWINDOWTHEME             As Long = (CCM_FIRST + &HB)
Public Const CCM_DPISCALE                   As Long = (CCM_FIRST + &HC)      '// wParam == Awareness

Public Const LVM_SETUNICODEFORMAT           As Long = CCM_SETUNICODEFORMAT
Public Const LVM_GETUNICODEFORMAT           As Long = CCM_GETUNICODEFORMAT

Public Const LVN_ITEMCHANGED  As Long = -100 - 1

Public Const LVCF_WIDTH  As Long = 2
Public Const LVCF_TEXT  As Long = 4
Public Const LVCF_SUBITEM  As Long = 8
Public Const LVIF_TEXT  As Long = 1

Public Const NM_CLICK  As Long = -2
Public Const NM_DBLCLK  As Long = -3

'/* Font Weights */
Public Const FW_DONTCARE As Long = 0
Public Const FW_THIN  As Long = 100
Public Const FW_EXTRALIGHT  As Long = 200
Public Const FW_LIGHT  As Long = 300
Public Const FW_NORMAL As Long = 400
Public Const FW_MEDIUM  As Long = 500
Public Const FW_SEMIBOLD As Long = 600
Public Const FW_BOLD  As Long = 700
Public Const FW_EXTRABOLD  As Long = 800
Public Const FW_HEAVY  As Long = 900

Public Const FW_ULTRALIGHT  As Long = FW_EXTRALIGHT
Public Const FW_REGULAR  As Long = FW_NORMAL
Public Const FW_DEMIBOLD  As Long = FW_SEMIBOLD
Public Const FW_ULTRABOLD  As Long = FW_EXTRABOLD
Public Const FW_BLACK  As Long = FW_HEAVY

Public Const OUT_DEFAULT_PRECIS  As Long = 0
Public Const OUT_STRING_PRECIS  As Long = 1
Public Const OUT_CHARACTER_PRECIS  As Long = 2
Public Const OUT_STROKE_PRECIS  As Long = 3
Public Const OUT_TT_PRECIS  As Long = 4
Public Const OUT_DEVICE_PRECIS  As Long = 5
Public Const OUT_RASTER_PRECIS  As Long = 6
Public Const OUT_TT_ONLY_PRECIS  As Long = 7
Public Const OUT_OUTLINE_PRECIS  As Long = 8
Public Const OUT_SCREEN_OUTLINE_PRECIS  As Long = 9
Public Const OUT_PS_ONLY_PRECIS  As Long = 10

Public Const CLIP_DEFAULT_PRECIS  As Long = 0
Public Const CLIP_CHARACTER_PRECIS  As Long = 1
Public Const CLIP_STROKE_PRECIS  As Long = 2
Public Const CLIP_MASK  As Long = &HF
Public Const CLIP_LH_ANGLES  As Long = &H10
Public Const CLIP_TT_ALWAYS  As Long = &H20
Public Const CLIP_DFA_DISABLE  As Long = &H40
Public Const CLIP_EMBEDDED  As Long = &H80

Public Const DEFAULT_QUALITY  As Long = 0
Public Const DRAFT_QUALITY  As Long = 1
Public Const PROOF_QUALITY  As Long = 2
Public Const NONANTIALIASED_QUALITY  As Long = 3
Public Const ANTIALIASED_QUALITY  As Long = 4
Public Const CLEARTYPE_QUALITY  As Long = 5
Public Const CLEARTYPE_NATURAL_QUALITY  As Long = 6

Public Const DEFAULT_PITCH  As Long = 0
Public Const FIXED_PITCH  As Long = 1
Public Const VARIABLE_PITCH  As Long = 2
Public Const MONO_FONT  As Long = 8

Public Const ANSI_CHARSET  As Long = 0
Public Const DEFAULT_CHARSET  As Long = 1
Public Const SYMBOL_CHARSET  As Long = 2
Public Const SHIFTJIS_CHARSET  As Long = 128
Public Const HANGEUL_CHARSET  As Long = 129
Public Const HANGUL_CHARSET  As Long = 129
Public Const GB2312_CHARSET  As Long = 134
Public Const CHINESEBIG5_CHARSET  As Long = 136
Public Const OEM_CHARSET  As Long = 255
Public Const JOHAB_CHARSET  As Long = 130
Public Const HEBREW_CHARSET  As Long = 177
Public Const ARABIC_CHARSET  As Long = 178
Public Const GREEK_CHARSET  As Long = 161
Public Const TURKISH_CHARSET  As Long = 162
Public Const VIETNAMESE_CHARSET  As Long = 163
Public Const THAI_CHARSET  As Long = 222
Public Const EASTEUROPE_CHARSET  As Long = 238
Public Const RUSSIAN_CHARSET  As Long = 204

Public Const MAC_CHARSET  As Long = 77
Public Const BALTIC_CHARSET  As Long = 186

'Type InitCommonControlsExStruct
Type InitCommonControlsExType
    lngSize As Long
    lngICC As Long
End Type

''Declare Function InitCommonControlsEx& Lib "comctl32" _
''    (ByVal lpInitCtrls&)
'
''Declare Function SetWindowSubclass& Lib "comctl32" _
''    (ByVal hwnd&, _
''     ByVal pfnSubclass&, _
''     ByVal uIdSubclass&, _
''     ByVal dwRefData&)
''Declare Function DefSubclassProc& Lib "comctl32" _
''    (ByVal hwnd&, _
''     ByVal uMsg&, _
''     ByVal wParam&, _
''     ByVal lParam&)
''Declare Function RemoveWindowSubclass& Lib "comctl32" _
''    (ByVal hwnd&, _
''     ByVal pfnSubclass&, _
''     ByVal uIdSubclass&)
'
Declare PtrSafe Function InitCommonControlsEx Lib "COMCTL32" (ByRef LPINITCOMMONCONTROLSEX As InitCommonControlsExType) As Long
Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'
''Declare Function CreateWindowExW& Lib "user32" _
''    (ByVal dwExStyle&, _
''     ByVal lpClassName&, _
''     ByVal lpWindowName&, _
''     ByVal dwStyle&, _
''     ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, _
''     ByVal HwndParent&, _
''     ByVal HMENU&, _
''     ByVal hInstance&, _
''     ByVal lpParam&)
''Declare Function SendMessageW& Lib "user32" _
''    (ByVal hwnd&, _
''     ByVal uMsg&, _
''     ByVal wParam&, _
''     ByVal lParam&)
''Declare Function GetFocus& Lib "user32" ()
''Declare Function SetFocus& Lib "user32" (ByVal hwnd&)
''Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" _
''    (pDest As Any, _
''     pSrc As Any, _
''     ByVal cbLen&)
'
Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hwnd As LongPtr) As Long

Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'
''Declare Function SysAllocString& Lib "Oleaut32" (ByVal ptr&)
Declare PtrSafe Function SysAllocString Lib "OleAut32.dll" (ByVal psz As LongPtr) As LongPtr
'
''Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
''Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
''    (ByVal H As Long, _
''     ByVal W As Long, _
''     ByVal E As Long, _
''     ByVal o As Long, _
''     ByVal W As Long, _
''     ByVal i As Long, _
''     ByVal u As Long, _
''     ByVal S As Long, _
''     ByVal C As Long, _
''     ByVal OP As Long, _
''     ByVal CP As Long, _
''     ByVal Q As Long, _
''     ByVal PAF As Long, _
''     ByVal F As String) As Long
'
Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As LongPtr
'WINGDIAPI HFONT   WINAPI CreateFontW( _In_ int cHeight, _In_ int cWidth, _In_ int cEscapement, _In_ int cOrientation, _In_ int cWeight, _In_ DWORD bItalic,
'                             _In_ DWORD bUnderline, _In_ DWORD bStrikeOut, _In_ DWORD iCharSet, _In_ DWORD iOutPrecision, _In_ DWORD iClipPrecision,
'                             _In_ DWORD iQuality, _In_ DWORD iPitchAndFamily, _In_opt_ LPCWSTR pszFaceName);
Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As LongPtr) As LongPtr

Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, lprcUpdate As RECT, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long



Type LVCOLUMNA
    mask        As Long
    fmt         As Long
    cx          As Long
    pszText     As String
    cchTextMax  As Long
    iSubItem    As Long
    iImage      As Long
    iOrder      As Long
    cxMin       As Long     '// min snap point
    cxDefault   As Long     '// default snap point
    cxIdeal     As Long     '// read only. ideal may not eqaul current width if auto sized (LVS_EX_AUTOSIZECOLUMNS) to a lesser width.
End Type

Type LVCOLUMNW
    mask        As Long
    fmt         As Long
    cx          As Long
    pszText     As LongPtr
    cchTextMax  As Long
    iSubItem    As Long
    iImage      As Long
    iOrder      As Long
    cxMin       As Long     '// min snap point
    cxDefault   As Long     '// default snap point
    cxIdeal     As Long     '// read only. ideal may not eqaul current width if auto sized (LVS_EX_AUTOSIZECOLUMNS) to a lesser width.
End Type

Type LVITEMA
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    state       As Long
    stateMask   As Long
    pszText     As String
    cchTextMax  As Long
    iImage      As Long
    lParam      As LongPtr
    iIndent     As Long
    iGroupId    As Long
    cColumns    As Long     '// tile view columns
    puColumns   As LongPtr
    piColFmt    As LongPtr
    iGroup      As Long     '// readonly. only valid for owner data.
End Type

Type LVITEMW
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    state       As Long
    stateMask   As Long
    pszText     As LongPtr
    cchTextMax  As Long
    iImage      As Long
    lParam      As LongPtr
    iIndent     As Long
    iGroupId    As Long
    cColumns    As Long     '// tile view columns
    puColumns   As LongPtr
    piColFmt    As LongPtr
    iGroup      As Long     '// readonly. only valid for owner data.
End Type

'typedef struct tagLVFINDINFOA
'{
'    UINT flags;
'    LPCSTR psz;
'    LPARAM lParam;
'    POINT pt;
'    UINT vkDirection;
'} LVFINDINFOA, *LPFINDINFOA;
'
'typedef struct tagLVFINDINFOW
'{
'    UINT flags;
'    LPCWSTR psz;
'    LPARAM lParam;
'    POINT pt;
'    UINT vkDirection;
'} LVFINDINFOW, *LPFINDINFOW;

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

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type TT
    hParent As LongPtr
    hChild As LongPtr
    pfn As LongPtr
End Type
'
Public TT As TT, acc As IAccessible

Public Function Redirect(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id As Long, ByVal lv As ListView) As LongPtr
    Redirect = lv.WndProc(hwnd, uMsg, wParam, lParam)
    
End Function




