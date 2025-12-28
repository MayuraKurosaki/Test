Attribute VB_Name = "ListViewModule"
Option Explicit
'
'API
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

' ============================================
' Notifications

Public Enum LVNotifications
  LVN_FIRST = -100&   ' &HFFFFFF9C   ' (0U-100U)
  LVN_LAST = -199&   ' &HFFFFFF39   ' (0U-199U)
                                                                          ' lParam points to:
  LVN_ITEMCHANGING = (LVN_FIRST - 0)            ' NMLISTVIEW, ?, rtn T/F
  LVN_ITEMCHANGED = (LVN_FIRST - 1)             ' NMLISTVIEW, ?
  LVN_INSERTITEM = (LVN_FIRST - 2)                  ' NMLISTVIEW, iItem
  LVN_DELETEITEM = (LVN_FIRST - 3)                 ' NMLISTVIEW, iItem
  LVN_DELETEALLITEMS = (LVN_FIRST - 4)         ' NMLISTVIEW, iItem = -1, rtn T/F

  LVN_COLUMNCLICK = (LVN_FIRST - 8)              ' NMLISTVIEW, iItem = -1, iSubItem = column
  LVN_BEGINDRAG = (LVN_FIRST - 9)                  ' NMLISTVIEW, iItem
  LVN_BEGINRDRAG = (LVN_FIRST - 11)              ' NMLISTVIEW, iItem

  LVN_ODCACHEHINT = (LVN_FIRST - 13)           ' NMLVCACHEHINT
  LVN_ITEMACTIVATE = (LVN_FIRST - 14)           ' v4.70 = NMHDR, v4.71 = NMITEMACTIVATE
  LVN_ODSTATECHANGED = (LVN_FIRST - 15)  ' NMLVODSTATECHANGE, rtn T/F
  LVN_HOTTRACK = (LVN_FIRST - 21)                 ' NMLISTVIEW, see docs, rtn T/F
  LVN_BEGINLABELEDITA = (LVN_FIRST - 5)        ' NMLVDISPINFO, iItem, rtn T/F
  LVN_ENDLABELEDITA = (LVN_FIRST - 6)           ' NMLVDISPINFO, see docs
 
  LVN_GETDISPINFOA = (LVN_FIRST - 50)            ' NMLVDISPINFO, see docs
  LVN_SETDISPINFOA = (LVN_FIRST - 51)            ' NMLVDISPINFO, see docs
  LVN_ODFINDITEMA = (LVN_FIRST - 52)             ' NMLVFINDITEM
 
  LVN_KEYDOWN = (LVN_FIRST - 55)                 ' NMLVKEYDOWN
  LVN_MARQUEEBEGIN = (LVN_FIRST - 56)       ' NMLISTVIEW, rtn T/F
  LVN_GETINFOTIPA = (LVN_FIRST - 57)             ' NMLVGETINFOTIP
  LVN_INCREMENTALSEARCHA = (LVN_FIRST - 62)
  LVN_INCREMENTALSEARCHW = (LVN_FIRST - 63)
'#If (WIN32_IE >= &H600) Then
  LVN_COLUMNDROPDOWN = (LVN_FIRST - 64)
  LVN_COLUMNOVERFLOWCLICK = (LVN_FIRST - 66)
'#End If
  LVN_BEGINSCROLL = (LVN_FIRST - 80)
  LVN_ENDSCROLL = (LVN_FIRST - 81)
  LVN_LINKCLICK = (LVN_FIRST - 84)
  LVN_GETEMPTYMARKUP = (LVN_FIRST - 87)
  LVN_GROUPCHANGED = (LVN_FIRST - 88)   ' Undocumented
  LVN_BEGINLABELEDITW = (LVN_FIRST - 75)
  LVN_ENDLABELEDITW = (LVN_FIRST - 76)
  LVN_GETDISPINFOW = (LVN_FIRST - 77)
  LVN_SETDISPINFOW = (LVN_FIRST - 78)
  LVN_ODFINDITEMW = (LVN_FIRST - 79)             ' NMLVFINDITEM
  LVN_GETINFOTIPW = (LVN_FIRST - 58)              ' NMLVGETINFOTIP

#If Unicode Then
  LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITW
  LVN_ENDLABELEDIT = LVN_ENDLABELEDITW
  LVN_GETDISPINFO = LVN_GETDISPINFOW
  LVN_SETDISPINFO = LVN_SETDISPINFOW
  LVN_ODFINDITEM = LVN_ODFINDITEMW         ' NMLVFINDITEM
  LVN_GETINFOTIP = LVN_GETINFOTIPW              ' NMLVGETINFOTIP
  LVN_INCREMENTALSEARCH = LVN_INCREMENTALSEARCHW
#Else
  LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITA
  LVN_ENDLABELEDIT = LVN_ENDLABELEDITA
  LVN_GETDISPINFO = LVN_GETDISPINFOA
  LVN_SETDISPINFO = LVN_SETDISPINFOA
  LVN_ODFINDITEM = LVN_ODFINDITEMA         ' NMLVFINDITEM
  LVN_GETINFOTIP = LVN_GETINFOTIPA              ' NMLVGETINFOTIP
  LVN_INCREMENTALSEARCH = LVN_INCREMENTALSEARCHA
#End If   ' UNICODE

End Enum   ' LVNotifications

'Public Const LVN_ITEMCHANGED  As Long = -100 - 1
' LVITEM mask
Public Enum LVITEM_mask
  LVIF_TEXT = &H1
  LVIF_IMAGE = &H2
  LVIF_PARAM = &H4
  LVIF_STATE = &H8
  LVIF_INDENT = &H10
  LVIF_GROUPID = &H100
  LVIF_COLUMNS = &H200
  LVIF_NORECOMPUTE = &H800
  LVIF_DI_SETITEM = &H1000   ' NMLVDISPINFO notification
  '6.0
  LVIF_COLFMT = &H10000
End Enum

Public Const LVCF_WIDTH  As Long = 2
Public Const LVCF_TEXT  As Long = 4
Public Const LVCF_SUBITEM  As Long = 8
'Public Const LVIF_TEXT  As Long = 1

Public Const LVSCW_AUTOSIZE   As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER   As Long = -2

Public Const NM_CLICK  As Long = -2
Public Const NM_DBLCLK  As Long = -3

' Font Families
'
Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.

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
Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
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
Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hWnd As LongPtr) As Long

Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
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
Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As LongPtr
'WINGDIAPI HFONT   WINAPI CreateFontW( _In_ int cHeight, _In_ int cWidth, _In_ int cEscapement, _In_ int cOrientation, _In_ int cWeight, _In_ DWORD bItalic,
'                             _In_ DWORD bUnderline, _In_ DWORD bStrikeOut, _In_ DWORD iCharSet, _In_ DWORD iOutPrecision, _In_ DWORD iClipPrecision,
'                             _In_ DWORD iQuality, _In_ DWORD iPitchAndFamily, _In_opt_ LPCWSTR pszFaceName);
Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As LongPtr) As LongPtr

Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, lprcUpdate As RECT, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type POINTF
    x As Single
    Y As Single
End Type

Public Type NMHDR
    hWndFrom As LongPtr   ' Window handle of control sending message
    IDFrom As LongPtr        ' Identifier of control sending message
    Code  As Long         ' Specifies the notification code
End Type

Public Type LVCOLUMNA
    mask        As Long
    fmt         As Long
    CX          As Long
    pszText     As String
    cchTextMax  As Long
    iSubItem    As Long
    iImage      As Long
    iOrder      As Long
    cxMin       As Long     '// min snap point
    cxDefault   As Long     '// default snap point
    cxIdeal     As Long     '// read only. ideal may not eqaul current width if auto sized (LVS_EX_AUTOSIZECOLUMNS) to a lesser width.
End Type

Public Type LVCOLUMNW
    mask        As Long
    fmt         As Long
    CX          As Long
    pszText     As LongPtr
    cchTextMax  As Long
    iSubItem    As Long
    iImage      As Long
    iOrder      As Long
    cxMin       As Long     '// min snap point
    cxDefault   As Long     '// default snap point
    cxIdeal     As Long     '// read only. ideal may not eqaul current width if auto sized (LVS_EX_AUTOSIZECOLUMNS) to a lesser width.
End Type

Public Type LVITEMA
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    State       As Long
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

Public Type LVITEMW
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    State       As Long
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

Public Type NMLVDISPINFOA   ' was LV_DISPINFO
    hdr As NMHDR
    Item As LVITEMA
End Type

Public Type NMLVDISPINFOW   ' was LV_DISPINFO
    hdr As NMHDR
    Item As LVITEMW
End Type

Public Type NMLISTVIEW   ' was NM_LISTVIEW
    hdr As NMHDR
    iItem As Long
    iSubItem As Long
    uNewState As Long
    uOldState As Long
    uChanged As Long
    ptAction As POINTAPI
    lParam As LongPtr
End Type
'typedef struct tagNMLISTVIEW
'{
'    NMHDR   hdr;
'    int     iItem;
'    int     iSubItem;
'    UINT    uNewState;
'    UINT    uOldState;
'    UINT    uChanged;
'    POINT   ptAction;
'    LPARAM  lParam;
'} NMLISTVIEW, *LPNMLISTVIEW;

Public Type NMITEMACTIVATE
    hdr As NMHDR
    iItem As Long
    iSubItem As Long
    uNewState As Long
    uOldState As Long
    uChanged As Long
    ptAction As POINTAPI
    lParam As LongPtr
    uKeyFlags As Long
End Type
'typedef struct tagNMITEMACTIVATE
'{
'    NMHDR   hdr;
'    int     iItem;
'    int     iSubItem;
'    UINT    uNewState;
'    UINT    uOldState;
'    UINT    uChanged;
'    POINT   ptAction;
'    LPARAM  lParam;
'    UINT    uKeyFlags;
'} NMITEMACTIVATE, *LPNMITEMACTIVATE;

Public Type VListItem
    sText As String
    sSubItems() As String
    iImage As Long
    iSubItemImages() As Long 'LVS_EX_SUBITEMIMAGES must be enabled, then must dim same as sSubItems
    iGrp As Long
    iPos As Long
End Type

'Public Type VListGroup
'    items() As Long
'    gid As Long 'groupid, doesn't have to be the same as the index
'                'but in the case of virtual groups should be, since
'                'alot of stuff goes by index
'End Type

Public VLItems() As VListItem
'Public VLGroups() As VListGroup
'Public lGroupCount As Long

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

'Type NMITEMACTIVATE
'    hdr(2)      As Long
'    iItem       As Long
'    iSubItem    As Long
'    buf(6)      As Long
'End Type

'Type NMLISTVIEW
'    hrd(2)      As Long
'    iItem       As Long
'    iSubItem    As Long
'    buf(4)      As Long
'End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Type TT
'    hParent As LongPtr
'    hChild As LongPtr
'    pfn As LongPtr
'End Type
''
'Public TT As TT ', acc As IAccessible
'Public pfn As LongPtr
'Public hParent As LongPtr
'Public hChild As LongPtr

'Public Function Redirect(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
'                    ByVal lParam As LongPtr, ByVal id As Long, ByVal lv As ListView) As LongPtr
'    Redirect = lv.WndProc(hWnd, uMsg, wParam, lParam)
'End Function

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
    MoveMemory LOWORD, dwValue, 2&
End Function

Public Function readCsvFile(FilePath As String) As Boolean
    readCsvFile = False
    

End Function

 'CSVファイルを指定シートに出力
Public Sub CsvToVListView(ByVal strFile As String, Optional ByVal CharSet As String = "Auto")
    Dim myArray() As Variant
 
    'readCsvでCSVを読み込み
    Dim strRec As String
    strRec = readCsv(strFile, CharSet)
 
    'CsvToJaggedで行・フィールドに分割してジャグ配列に
'    Dim jagArray() As Variant
    Call CsvToJagged(strRec)
'    Dim jagArray() As Variant
'    jagArray = CsvToJagged(strRec)
 
'    'JaggedTo2Dでジャグ配列を2次元配列に変換
'    Call JaggedTo2D(jagArray, myArray)
'
'
'    '2次元配列→シート
'    ws.Range("A1").Resize(UBound(myArray, 1), UBound(myArray, 2)) = myArray
    Debug.Print VLItems(3).sSubItems(4)
End Sub

Private Sub JaggedTo2D(ByRef jagArray() As Variant, _
                       ByRef twoDArray As Variant)
    'ジャグ配列の最大列数取得
    Dim maxCol As Long, v As Variant
    maxCol = 0
    For Each v In jagArray
        If UBound(v) > maxCol Then
            maxCol = UBound(v)
        End If
    Next
  
    'ジャグ配列→2次元配列
    Dim i1 As Long, i2 As Long
    ReDim twoDArray(1 To UBound(jagArray), 1 To maxCol)
    For i1 = 1 To UBound(jagArray)
        For i2 = 1 To UBound(jagArray(i1))
            twoDArray(i1, i2) = jagArray(i1)(i2)
        Next
    Next
End Sub

'Private Function CsvToJagged(ByVal strRec As String) As Variant()
Private Sub CsvToJagged(ByVal strRec As String)
'    Dim childArray() As Variant 'ジャグ配列の子配列
    Dim lngQuate As Long 'ダブルクォーテーション数
    Dim strCell As String '1フィールド文字列
    Dim blnCrLf As Boolean '改行判定
    Dim i As Long '行位置
    Dim j As Long '列位置
    Dim k As Long

    ReDim VLItems(1 To 1)
    ReDim VLItems(1).sSubItems(1 To 1)
    
'    ReDim CsvToJagged(1 To 1) 'ジャグ配列の初期化
'    ReDim childArray(1 To 1) 'ジャグ配列の子配列の初期化
    i = 1 'シートの1行目から出力
    j = 0 '列位置はputChildArrayでカウントアップ
    lngQuate = 0 'ダブルクォーテーションの数
    strCell = ""
    For k = 1 To Len(strRec)
        Select Case Mid(strRec, k, 1)
            Case vbLf, vbCr '「"」が偶数なら改行、奇数ならただの文字
                If lngQuate Mod 2 = 0 Then
                    blnCrLf = False
                    If k > 1 Then '改行のCrLfはCrで改行判定済なので無視する
                        If Mid(strRec, k - 1, 2) = vbCrLf Then
                            blnCrLf = True
                        End If
                    End If
                    If blnCrLf = False Then
                        ReDim Preserve VLItems(1 To i)
                        ReDim VLItems(i).sSubItems(1 To 1) '子配列の初期化
                        i = i + 1 '列位置
                        j = 0 '列位置
                        lngQuate = 0 'ダブルクォーテーション数

'                        Call putChildArray(VLItems(i).sSubItems, j, strCell, lngQuate)
'                        Call putjagArray(i, j, lngQuate, strCell)
'                        Call putChildArray(childArray, j, strCell, lngQuate)
                        'これが改行となる
'                        Call putjagArray(CsvToJagged, childArray, i, j, lngQuate, strCell)
                    End If
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case ",", vbTab '「"」が偶数なら区切り、奇数ならただの文字
                If lngQuate Mod 2 = 0 Then
'                    Call putChildArray(childArray, j, strCell, lngQuate)
'                    Debug.Print "NewField:" & i & ":" & j
                    Call putChildArray(VLItems(i).sSubItems, j, strCell, lngQuate)
                Else
                    strCell = strCell & Mid(strRec, k, 1)
                End If
            Case """" '「"」のカウントをとる
                lngQuate = lngQuate + 1
                strCell = strCell & Mid(strRec, k, 1)
            Case Else
                strCell = strCell & Mid(strRec, k, 1)
        End Select
    Next

    '最終行の最終列の処理
    If j > 0 And strCell <> "" Then
        Call putjagArray(i, j, lngQuate, strCell)
        Call putChildArray(VLItems(i).sSubItems, j, strCell, lngQuate)
'        Call putChildArray(childArray, j, strCell, lngQuate)
'        Call putjagArray(CsvToJagged, childArray, i, j, lngQuate, strCell)
    End If
End Sub

Private Sub putjagArray(ByRef i As Long, ByRef j As Long, ByRef lngQuate As Long, ByRef strCell As String)
'Private Sub putjagArray(ByRef jagArray() As Variant, _
                        ByRef childArray() As Variant, _
                        ByRef i As Long, _
                        ByRef j As Long, _
                        ByRef lngQuate As Long, _
                        ByRef strCell As String)
'    If i > UBound(jagArray) Then '常に成立するが一応記述
'    If i > UBound(VLItems) Then '常に成立するが一応記述
'        ReDim Preserve VLItems(1 To i)
''        ReDim Preserve jagArray(1 To i)
'    End If
''    VLItems(i) = childArray '子配列をジャグ配列に入れる
''    jagArray(i) = childArray '子配列をジャグ配列に入れる
''    ReDim childArray(1 To 1) '子配列の初期化
'    ReDim VLItems(i).sSubItems(1 To 1) '子配列の初期化
    
    strCell = Replace(strCell, """""", """")
    '前後の「"」を削除
    If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
        If Len(strCell) <= 2 Then
            strCell = ""
        Else
            strCell = Mid(strCell, 2, Len(strCell) - 2)
        End If
    End If
    VLItems(i).sSubItems(1) = strCell
    
    i = i + 1 '列位置
    j = 0 '列位置
    lngQuate = 0 'ダブルクォーテーション数
    strCell = "" '1フィールド文字列
End Sub

'1フィールドごとにセルに出力
'Private Sub putChildArray(ByRef childArray() As Variant, ByRef j As Long, ByRef strCell As String, ByRef lngQuate As Long)
Private Sub putChildArray(ByRef childArray() As String, ByRef j As Long, ByRef strCell As String, ByRef lngQuate As Long)
    j = j + 1
    '「""」を「"」で置換
    strCell = Replace(strCell, """""", """")
    '前後の「"」を削除
    If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
        If Len(strCell) <= 2 Then
            strCell = ""
        Else
            strCell = Mid(strCell, 2, Len(strCell) - 2)
        End If
    End If
    If j > UBound(childArray) Then
        ReDim Preserve childArray(1 To j)
    End If
    childArray(j) = strCell
    strCell = ""
    lngQuate = 0
End Sub

'文字コードを自動判別し、全行をCrLf区切りに統一してStringに入れる
Private Function readCsv(ByVal strFile As String, ByVal CharSet As String) As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim inTS As Object
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
 
    Dim strRec As String
    Dim i As Long
    Dim aryRec() As String

    If CharSet = "Auto" Then CharSet = getCharSet(strFile)
    Debug.Print "CharSet:" & CharSet
    Select Case UCase(CharSet)
        Case "UTF-8", "UTF-8N"
            'ADOを使って読込、その後の処理を統一するため全レコードをCrLfで結合
            'Set inTS = objFSO.OpenTextFile(strFile, ForAppending)
            Set inTS = objFSO.OpenTextFile(strFile, 8)
            i = inTS.Line - 1
            inTS.Close
            ReDim aryRec(i)
            With adoSt
                '.Type = adTypeText
                .Type = 2
                .CharSet = "UTF-8"
                .Open
                .LoadFromFile strFile
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
            Set inTS = objFSO.OpenTextFile(strFile, , , -1)
            strRec = inTS.ReadAll
            inTS.Close
        Case "SHIFT-JIS"
            Set inTS = objFSO.OpenTextFile(strFile)
            strRec = inTS.ReadAll
            inTS.Close
        Case Else
            'EUC-JP、UTF-32については未テスト
            MsgBox "文字コードを確認してください。" & vbLf & CharSet
            Stop
    End Select
    Set inTS = Nothing
    Set objFSO = Nothing
    readCsv = strRec
End Function

'文字コードの自動判別
Private Function getCharSet(strFileName As String) As String
    Dim bytes() As Byte
    Dim intFileNo As Integer
    ReDim bytes(FileLen(strFileName))
    intFileNo = FreeFile
    Open strFileName For Binary As #intFileNo
    Get #intFileNo, , bytes
    Close intFileNo
 
    'BOMによる判断
    getCharSet = getCharFromBOM(bytes)
 
    'BOMなしをデータの文字コードで判別
    If getCharSet = "" Then
        getCharSet = getCharFromCode(bytes)
    End If
 
    Debug.Print strFileName & " : " & getCharSet
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
                getCharFromCode = "Shift_JIS"
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
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bDollar And b3 = bB Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bOpen And (b3 = bB Or b3 = bJ) Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            ElseIf b2 = bOpen And b3 = bI Then
                getCharFromCode = "Shift_JIS"
                Exit Function
            End If
            If i < bLen - 3 Then
                b4 = bytes(i + 3)
                If b2 = bDollar And b3 = bOpen And b4 = bD Then
                    getCharFromCode = "Shift_JIS"
                    Exit Function
                End If
                If i < bLen - 5 And _
                    b2 = bAnd And b3 = bAt And b4 = bEscape And _
                    bytes(i + 4) = bDollar And bytes(i + 5) = bB Then
                    getCharFromCode = "Shift_JIS"
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


