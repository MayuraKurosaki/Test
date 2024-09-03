Attribute VB_Name = "ListViewModule"
Option Explicit

Public hLVVG As LongPtr
Public hLVVGHdr As LongPtr
Public Const IDD_LISTVIEW = 101
Public Const WC_LISTVIEW = "SysListView32"

Public Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Enum LVStyles
  LVS_ICON = &H0
  LVS_REPORT = &H1
  LVS_SMALLICON = &H2
  LVS_LIST = &H3
  LVS_TYPEMASK = &H3
  LVS_SINGLESEL = &H4
  LVS_SHOWSELALWAYS = &H8
  LVS_SORTASCENDING = &H10
  LVS_SORTDESCENDING = &H20
  LVS_SHAREIMAGELISTS = &H40
  LVS_NOLABELWRAP = &H80
  LVS_AUTOARRANGE = &H100
  LVS_EDITLABELS = &H200
  LVS_OWNERDATA = &H1000
  LVS_NOSCROLL = &H2000
  LVS_TYPESTYLEMASK = &HFC00
  LVS_ALIGNTOP = &H0
  LVS_ALIGNLEFT = &H800
  LVS_ALIGNMASK = &HC00
  LVS_OWNERDRAWFIXED = &H400
  LVS_NOCOLUMNHEADER = &H4000
  LVS_NOSORTHEADER = &H8000&
End Enum   ' LVStyles

Public Enum LVStylesEx
  LVS_EX_GRIDLINES = &H1
  LVS_EX_SUBITEMIMAGES = &H2
  LVS_EX_CHECKBOXES = &H4
  LVS_EX_TRACKSELECT = &H8
  LVS_EX_HEADERDRAGDROP = &H10
  LVS_EX_FULLROWSELECT = &H20         ' // applies to report mode only
  LVS_EX_ONECLICKACTIVATE = &H40
  LVS_EX_TWOCLICKACTIVATE = &H80
  LVS_EX_FLATSB = &H100
  LVS_EX_REGIONAL = &H200
  LVS_EX_INFOTIP = &H400              ' listview does InfoTips for you
  LVS_EX_UNDERLINEHOT = &H800
  LVS_EX_UNDERLINECOLD = &H1000
  LVS_EX_MULTIWORKAREAS = &H2000
  LVS_EX_LABELTIP = &H4000
  LVS_EX_BORDERSELECT = &H8000
  LVS_EX_DOUBLEBUFFER = &H10000
  LVS_EX_HIDELABELS = &H20000
  LVS_EX_SINGLEROW = &H40000
  LVS_EX_SNAPTOGRID = &H80000 '// Icons automatically snap to grid.
  LVS_EX_SIMPLESELECT = &H100000        '// Also changes overlay rendering to top right for icon mode.
  LVS_EX_JUSTIFYCOLUMNS = &H200000      '// Icons are lined up in columns that use up the whole view area.
  LVS_EX_TRANSPARENTBKGND = &H400000    '// Background is painted by the parent via WM_PRINTCLIENT
  LVS_EX_TRANSPARENTSHADOWTEXT = &H800000    '// Enable shadow text on transparent backgrounds only (useful with bitmaps)
  LVS_EX_AUTOAUTOARRANGE = &H1000000    '// Icons automatically arrange if no icon positions have been set
  LVS_EX_HEADERINALLVIEWS = &H2000000   '// Display column header in all view modes
  LVS_EX_AUTOCHECKSELECT = &H8000000
  LVS_EX_AUTOSIZECOLUMNS = &H10000000
  LVS_EX_COLUMNSNAPPOINTS = &H40000000
  LVS_EX_COLUMNOVERFLOW = &H80000000
End Enum

' value returned by many listview messages indicating
' the index of no listview item (user defined)
Public Const LVI_NOITEM = &HFFFFFFFF

' messages
Public Const LVM_FIRST = &H1000
Public Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_INSERTITEM = (LVM_FIRST + 7)
Public Const LVM_DELETEITEM = (LVM_FIRST + 8)
Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Public Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
Public Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_FINDITEM = (LVM_FIRST + 13)
Public Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Public Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_GETSTRINGWIDTH = (LVM_FIRST + 17)
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SCROLL = (LVM_FIRST + 20)
Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
Public Const LVM_ARRANGE = (LVM_FIRST + 22)
Public Const LVM_EDITLABEL = (LVM_FIRST + 23)
Public Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)
Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_SETCOLUMN = (LVM_FIRST + 26)
Public Const LVM_INSERTCOLUMN = (LVM_FIRST + 27)
Public Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Public Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Public Const LVM_GETORIGIN = (LVM_FIRST + 41)
Public Const LVM_UPDATE = (LVM_FIRST + 42)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT = (LVM_FIRST + 45)
Public Const LVM_SETITEMTEXT = (LVM_FIRST + 46)
Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
Public Const LVM_SORTITEMS = (LVM_FIRST + 48)
Public Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Public Const LVM_GETITEMSPACING = (LVM_FIRST + 51)
Public Const LVM_GETISEARCHSTRING = (LVM_FIRST + 52)
Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_SETHOTITEM = (LVM_FIRST + 60)
Public Const LVM_GETHOTITEM = (LVM_FIRST + 61)
Public Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
Public Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
Public Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
Public Const LVM_SETWORKAREAS = (LVM_FIRST + 65)
Public Const LVM_GETSELECTIONMARK = (LVM_FIRST + 66)
Public Const LVM_SETSELECTIONMARK = (LVM_FIRST + 67)
Public Const LVM_SETBKIMAGE As Long = (LVM_FIRST + 68)
Public Const LVM_GETBKIMAGE As Long = (LVM_FIRST + 69)
Public Const LVM_GETWORKAREAS As Long = (LVM_FIRST + 70)
Public Const LVM_SETHOVERTIME As Long = (LVM_FIRST + 71)
Public Const LVM_GETHOVERTIME As Long = (LVM_FIRST + 72)
Public Const LVM_GETNUMBEROFWORKAREAS As Long = (LVM_FIRST + 73)
Public Const LVM_SETTOOLTIPS As Long = (LVM_FIRST + 74)
Public Const LVM_GETITEMW = (LVM_FIRST + 75)
Public Const LVM_SETITEMW = (LVM_FIRST + 76)  'Unicode
Public Const LVM_INSERTITEMW = (LVM_FIRST + 77) 'Unicode
Public Const LVM_GETTOOLTIPS As Long = (LVM_FIRST + 78)
Public Const LVM_GETHOTLIGHTCOLOR = (LVM_FIRST + 79) 'UNDOCUMENTED
Public Const LVM_SETHOTLIGHTCOLOR = (LVM_FIRST + 80) 'UNDOCUMENTED
Public Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
Public Const LVM_SETRANGEOBJECT = (LVM_FIRST + 82) 'UNDOCUMENTED
Public Const LVM_FINDITEMW                  As Long = (LVM_FIRST + 83) 'Unicode
Public Const LVM_RESETEMPTYTEXT = (LVM_FIRST + 84) 'UNDOCUMENTED
Public Const LVM_SETFROZENITEM = (LVM_FIRST + 85) 'UNDOCUMENTED
Public Const LVM_GETFROZENITEM = (LVM_FIRST + 86) 'UNDOCUMENTED
Public Const LVM_GETSTRINGWIDTHW = (LVM_FIRST + 87)
Public Const LVM_SETFROZENSLOT = (LVM_FIRST + 88) 'UNDOCUMENTED
Public Const LVM_GETFROZENSLOT = (LVM_FIRST + 89) 'UNDOCUMENTED
Public Const LVM_SETVIEWMARGIN = (LVM_FIRST + 90) 'UNDOCUMENTED
Public Const LVM_GETVIEWMARGIN = (LVM_FIRST + 91) 'UNDOCUMENTED
Public Const LVM_GETGROUPSTATE = (LVM_FIRST + 92)
Public Const LVM_GETFOCUSEDGROUP = (LVM_FIRST + 93)
Public Const LVM_EDITGROUPLABEL = (LVM_FIRST + 94) 'UNDOCUMENTED
Public Const LVM_GETCOLUMNW                As Long = (LVM_FIRST + 95) 'Unicode
Public Const LVM_SETCOLUMNW                As Long = (LVM_FIRST + 96) 'Unicode
Public Const LVM_INSERTCOLUMNW             As Long = (LVM_FIRST + 97) 'Unicode
Public Const LVM_GETGROUPRECT             As Long = (LVM_FIRST + 98)

Public Const LVM_GETITEMTEXTW = (LVM_FIRST + 115)     'Unicode
Public Const LVM_SETITEMTEXTW = (LVM_FIRST + 116)           'Unicode
Public Const LVM_GETISEARCHSTRINGW = (LVM_FIRST + 117)
Public Const LVM_EDITLABELW = (LVM_FIRST + 118)

Public Const LVM_SETBKIMAGEW = (LVM_FIRST + 138)
Public Const LVM_GETBKIMAGEW = (LVM_FIRST + 139)
Public Const LVM_SETSELECTEDCOLUMN = (LVM_FIRST + 140)
Public Const LVM_SETTILEWIDTH = (LVM_FIRST + 141)
Public Const LVM_SETVIEW As Long = (LVM_FIRST + 142)
Public Const LVM_GETVIEW As Long = (LVM_FIRST + 143)

Public Const LVM_INSERTGROUP = (LVM_FIRST + 145)

Public Const LVM_SETGROUPINFO = (LVM_FIRST + 147)

Public Const LVM_GETGROUPINFO = (LVM_FIRST + 149)
Public Const LVM_REMOVEGROUP = (LVM_FIRST + 150)
Public Const LVM_MOVEGROUP = (LVM_FIRST + 151)
Public Const LVM_GETGROUPCOUNT            As Long = (LVM_FIRST + 152)
Public Const LVM_GETGROUPINFOBYINDEX      As Long = (LVM_FIRST + 153)
Public Const LVM_MOVEITEMTOGROUP = (LVM_FIRST + 154)
Public Const LVM_SETGROUPMETRICS = (LVM_FIRST + 155)
Public Const LVM_GETGROUPMETRICS = (LVM_FIRST + 156)
Public Const LVM_ENABLEGROUPVIEW = (LVM_FIRST + 157)
Public Const LVM_SORTGROUPS = (LVM_FIRST + 158)
Public Const LVM_INSERTGROUPSORTED = (LVM_FIRST + 159)
Public Const LVM_REMOVEALLGROUPS = (LVM_FIRST + 160)
Public Const LVM_HASGROUP = (LVM_FIRST + 161)
Public Const LVM_SETTILEVIEWINFO = (LVM_FIRST + 162)
Public Const LVM_GETTILEVIEWINFO = (LVM_FIRST + 163)
Public Const LVM_SETTILEINFO = (LVM_FIRST + 164)
Public Const LVM_GETTILEINFO = (LVM_FIRST + 165)
Public Const LVM_SETINSERTMARK = (LVM_FIRST + 166)
Public Const LVM_GETINSERTMARK = (LVM_FIRST + 167)
Public Const LVM_INSERTMARKHITTEST = (LVM_FIRST + 168)
Public Const LVM_GETINSERTMARKRECT = (LVM_FIRST + 169)
Public Const LVM_SETINSERTMARKCOLOR = (LVM_FIRST + 170)
Public Const LVM_GETINSERTMARKCOLOR = (LVM_FIRST + 171)

Public Const LVM_SETINFOTIP = (LVM_FIRST + 173)
Public Const LVM_GETSELECTEDCOLUMN = (LVM_FIRST + 174)
Public Const LVM_ISGROUPVIEWENABLED = (LVM_FIRST + 175)
Public Const LVM_GETOUTLINECOLOR = (LVM_FIRST + 176)
Public Const LVM_SETOUTLINECOLOR = (LVM_FIRST + 177)
Public Const LVM_SETKEYBOARDSELECTED = (LVM_FIRST + 178)  'UNDOCUMENTED
Public Const LVM_CANCELEDITLABEL = (LVM_FIRST + 179)
Public Const LVM_MAPINDEXTOID = (LVM_FIRST + 180)
Public Const LVM_MAPIDTOINDEX = (LVM_FIRST + 181)
Public Const LVM_ISITEMVISIBLE = (LVM_FIRST + 182)
Public Const LVM_EDITSUBITEM = (LVM_FIRST + 183)          'UNDOCUMENTED
Public Const LVM_ENSURESUBITEMVISIBLE = (LVM_FIRST + 184) 'UNDOCUMENTED
Public Const LVM_GETCLIENTRECT = (LVM_FIRST + 185)        'UNDOCUMENTED
Public Const LVM_GETFOCUSEDCOLUMN = (LVM_FIRST + 186)     'UNDOCUMENTED
Public Const LVM_SETOWNERDATACALLBACK = (LVM_FIRST + 187) 'UNDOCUMENTED
Public Const LVM_RECOMPUTEITEMS = (LVM_FIRST + 188)      'UNDOCUMENTED
Public Const LVM_QUERYINTERFACE = (LVM_FIRST + 189)      'UNDOCUMENTED: NOT OFFICIAL NAME
Public Const LVM_SETGROUPSUBSETCOUNT = (LVM_FIRST + 190) 'UNDOCUMENTED
Public Const LVM_GETGROUPSUBSETCOUNT = (LVM_FIRST + 191) 'UNDOCUMENTED
Public Const LVM_ORDERTOINDEX = (LVM_FIRST + 192)        'UNDOCUMENTED
Public Const LVM_GETACCVERSION = (LVM_FIRST + 193)       'UNDOCUMENTED
Public Const LVM_MAPACCIDTOACCINDEX = (LVM_FIRST + 194)  'UNDOCUMENTED
Public Const LVM_MAPACCINDEXTOACCID = (LVM_FIRST + 195)  'UNDOCUMENTED
Public Const LVM_GETOBJECTCOUNT = (LVM_FIRST + 196)      'UNDOCUMENTED
Public Const LVM_GETOBJECTRECT = (LVM_FIRST + 197)       'UNDOCUMENTED
Public Const LVM_ACCHITTEST = (LVM_FIRST + 198)          'UNDOCUMENTED
Public Const LVM_GETFOCUSEDOBJECT = (LVM_FIRST + 199)    'UNDOCUMENTED
Public Const LVM_GETOBJECTROLE = (LVM_FIRST + 200)       'UNDOCUMENTED
Public Const LVM_GETOBJECTSTATE = (LVM_FIRST + 201)      'UNDOCUMENTED
Public Const LVM_ACCNAVIGATE = (LVM_FIRST + 202)         'UNDOCUMENTED
Public Const LVM_INVOKEDEFAULTACTION = (LVM_FIRST + 203) 'UNDOCUMENTED
Public Const LVM_GETEMPTYTEXT = (LVM_FIRST + 204)
Public Const LVM_GETFOOTERRECT = (LVM_FIRST + 205)
Public Const LVM_GETFOOTERINFO = (LVM_FIRST + 206)
Public Const LVM_GETFOOTERITEMRECT = (LVM_FIRST + 207)
Public Const LVM_GETFOOTERITEM = (LVM_FIRST + 208)
Public Const LVM_GETITEMINDEXRECT = (LVM_FIRST + 209)
Public Const LVM_SETITEMINDEXSTATE = (LVM_FIRST + 210)
Public Const LVM_GETNEXTITEMINDEX = (LVM_FIRST + 211)
Public Const LVM_SETPRESERVEALPHA = (LVM_FIRST + 212)    'UNDOCUMENTED

Public Const LVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
Public Const LVM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT

Public Const I_IMAGECALLBACK As Long = (-1)
Public Const I_IMAGENONE = (-2)
Public Const I_COLUMNSCALLBACK As Long = (-1)
Public Const I_GROUPIDCALLBACK As Long = (-1)
Public Const I_GROUPIDNONE As Long = (-2)
Public Const LPSTR_TEXTCALLBACKA = (-1)
Public Const LPSTR_TEXTCALLBACKW = (-1)

Public Enum LVTVI_Flags
    LVTVIF_AUTOSIZE = &H0
    LVTVIF_FIXEDWIDTH = &H1
    LVTVIF_FIXEDHEIGHT = &H2
    LVTVIF_FIXEDSIZE = &H3
    '6.0
    LVTVIF_EXTENDED = &H4
End Enum
Public Enum LVTVI_Mask
    LVTVIM_TILESIZE = &H1
    LVTVIM_COLUMNS = &H2
    LVTVIM_LABELMARGIN = &H4
End Enum

Public Type SIZELVT
    CX As Long
    CY As Long
End Type
Public Type LVTILEVIEWINFO
    cbSize As Long
    dwMask As LVTVI_Mask ';     //LVTVIM_*
    dwFlags As LVTVI_Flags ';    //LVTVIF_*
    SizeTile As SIZELVT ' ;
    cLines As Long
    RCLabelMargin As RECT
End Type

Public Type LVTILEINFO
    cbSize As Long
    iItem As Long
    cColumns As Long
    puColumns As Long
'#if (_WIN32_WINNT >= 0x0600)
    piColFmt As Long
'#End If
End Type


Public Const LV_VIEW_ICON As Long = &H0
Public Const LV_VIEW_DETAILS As Long = &H1
Public Const LV_VIEW_SMALLICON As Long = &H2
Public Const LV_VIEW_LIST As Long = &H3
Public Const LV_VIEW_TILE As Long = &H4&
'below are not part of API, but are valid views nonetheless
Public Const LV_VIEW_XLICON As Long = &H5&
Public Const LV_VIEW_THUMBNAIL As Long = &H6&

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

Public Enum LVNSCH 'LVN_INCREMENTALSEARCH codes
    LVNSCH_DEFAULT = -1
    LVNSCH_ERROR = -2
    LVNSCH_IGNORE = -3
End Enum
' LVM_GET/SETIMAGELIST wParam

Public Enum LV_ImageList
    LVSIL_NORMAL = 0
    LVSIL_SMALL = 1
    LVSIL_STATE = 2
    LVSIL_GROUPHEADER = 3
    LVSIL_FOOTER = 4 'UNDOCUMENTED: For footer items... see IListViewFooter
End Enum

' LVM_GET/SETITEM lParam
Public Type LVITEM   ' was LV_ITEM
  mask As LVITEM_mask
  iItem As Long
  iSubItem As Long
  State As LVITEM_state
  StateMask As Long
  pszText As LongPtr  ' if String, must be pre-allocated
  cchTextMax As Long
  iImage As Long
  lParam As Long
'#If (WIN32_IE >= &H300) Then
  iIndent As Long
  iGroupId As Long
  cColumns As Long
  puColumns As Long
'#End If
End Type
Public Type LVITEM_S   ' LVITEM with pszText as string
    mask As LVITEM_mask
    iItem As Long
    iSubItem As Long
    State As LVITEM_state
  StateMask As Long
  pszText As String  ' if String, must be pre-allocated
  cchTextMax As Long
  iImage As Long
  lParam As Long
'#If (WIN32_IE >= &H300) Then
  iIndent As Long
  iGroupId As Long
  cColumns As Long
  puColumns As Long
'#End If
End Type
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

' LVITEM state, stateMask, LVM_SETCALLBACKMASK wParam
Public Enum LVITEM_state
  LVIS_FOCUSED = &H1
  LVIS_SELECTED = &H2
  LVIS_CUT = &H4
  LVIS_DROPHILITED = &H8
  LVIS_GLOW = &H10
  LVIS_ACTIVATING = &H20
 
  LVIS_OVERLAYMASK = &HF00
  LVIS_STATEIMAGEMASK = &HF000
End Enum
Public Type LVBKIMAGE
  ulFlags As LVBKIMAGE_Flags
  hBm As Long
  pszImage As Long  ' if String, must be pre-allocated
  cchImageMax As Long
  XOffsetPercent As Long
  YOffsetPercent As Long
End Type
Public Enum LVBKIMAGE_Flags
    LVBKIF_SOURCE_NONE = &H0
    LVBKIF_SOURCE_HBITMAP = &H1
    LVBKIF_SOURCE_URL = &H2
    LVBKIF_SOURCE_MASK = &H3
    LVBKIF_STYLE_NORMAL = &H0
    LVBKIF_STYLE_TILE = &H10
    LVBKIF_STYLE_MASK = &H10
  '5.0
    LVBKIF_FLAG_TILEOFFSET = &H100
    LVBKIF_TYPE_WATERMARK = &H10000000
    LVBKIF_FLAG_ALPHABLEND = &H20000000
End Enum

' LVM_GETNEXTITEM LOWORD(lParam)
Public Enum LVNI_Flags
    LVNI_ALL = &H0
    LVNI_FOCUSED = &H1
    LVNI_SELECTED = &H2
    LVNI_CUT = &H4
    LVNI_DROPHILITED = &H8
    
    LVNI_ABOVE = &H100
    LVNI_BELOW = &H200
    LVNI_TOLEFT = &H400
    LVNI_TORIGHT = &H800
'#If (WIN32_IE >= &H600) Then
    LVNI_STATEMASK = (LVNI_FOCUSED Or LVNI_SELECTED Or LVNI_CUT Or LVNI_DROPHILITED)
    LVNI_DIRECTIONMASK = (LVNI_ABOVE Or LVNI_BELOW Or LVNI_TOLEFT Or LVNI_TORIGHT)

    LVNI_PREVIOUS = &H20
    LVNI_VISIBLEORDER = &H10
    LVNI_VISIBLEONLY = &H40
    LVNI_SAMEGROUPONLY = &H80
'#End If
End Enum
' LVM_GETITEMRECT rc.Left (lParam)
Public Enum LVIR_Flags
    LVIR_BOUNDS = 0
    LVIR_ICON = 1
    LVIR_LABEL = 2
    LVIR_SELECTBOUNDS = 3
End Enum
Public Enum LVM_SETITEMCOUNT_lParam
    LVSICF_NOINVALIDATEALL = &H1
    LVSICF_NOSCROLL = &H2
End Enum

' LVM_HITTEST lParam
Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINT
  Flags As LVHT_flags
  iItem As Long
'#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
'#End If
'#If (WIN32_IE >= &H600) then
  iGroup As Long
'#End If
End Type
Public Enum LVA_Flags
  LVA_DEFAULT = &H0
  LVA_ALIGNLEFT = &H1
  LVA_ALIGNTOP = &H2
  LVA_SNAPTOGRID = &H5
End Enum
Public Enum LVHT_flags
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
Public Type LVFINDINFO   ' was LV_FINDINFO
  Flags As LVFINDINFO_flags
  psz As String  ' if String, must be pre-allocated
  lParam As Long
  pt As POINT
  VKDirection As Long
End Type
 
Public Enum LVFINDINFO_flags
  LVFI_PARAM = &H1
  LVFI_STRING = &H2
  LVFI_SUBSTRING = &H4 'same as LVFI_PARTIAL
  LVFI_PARTIAL = &H8
  LVFI_WRAP = &H20
  LVFI_NEARESTXY = &H40
End Enum
Public Const LVFF_ITEMCOUNT = &H1
Public Type LVFOOTERINFO
     mask As Long 'must be LVFF_ITEMCOUNT
     pszText As Long 'not supported, must be 0
     cchText As Long 'not supported, must be 0
     cItems As Long
End Type
Public Enum LVFOOTERITEM_Flags
    LVFIF_TEXT = &H1
    LVFIF_STATE = &H2
End Enum
' footer item state
Public Const LVFIS_FOCUSED = &H1

Public Type LVFOOTERITEM
    mask As LVFOOTERITEM_Flags
    iItem As Long
    pszText As Long
    cchTextMax As Long
    State As Long
    StateMask As Long
End Type

Public Const LVIM_AFTER = &H1
Public Type LVINSERTMARK
    cbSize As Long
    dwFlags As Long 'must be LVIM_AFTER
    iItem As Long
    dwReserved As Long 'must be 0
End Type

Public Type LVITEMINDEX
    iItem As Long '          // listview item index
    iGroup As Long
End Type
Public Type LVSETINFOTIP
    cbSize As Long
    dwFlags As Long
    pszText As Long ' LPWSTR
    iItem As Long
    iSubItem As Long
End Type


' key flags stored in uKeyFlags
Public Const LVKF_ALT = &H1
Public Const LVKF_CONTROL = &H2
Public Const LVKF_SHIFT = &H4
' #end If '(_WIN32_IE >= =&H0400)

Public Type LVCOLUMN   ' was LV_COLUMN
  mask As LVCOLUMN_mask
  fmt As LVCOLUMN_fmt
  CX As Long
  pszText As String  ' if String, must be pre-allocated
  cchTextMax As Long
  iSubItem As Long
'#If (WIN32_IE >= &H300) Then
  iImage As Long
  iOrder As Long
'#End If
'#if (WIN32_IE >= &H600)
  cxMin As Long
  cxDefault As Long
  cxIdeal As Long
'#End If
End Type
Public Type LVCOLUMNW   ' was LV_COLUMN
  mask As LVCOLUMN_mask
  fmt As LVCOLUMN_fmt
  CX As Long
  pszText As Long  ' if String, must be pre-allocated
  cchTextMax As Long
  iSubItem As Long
'#If (WIN32_IE >= &H300) Then
  iImage As Long
  iOrder As Long
'#End If
'#if (WIN32_IE >= &H600)
  cxMin As Long
  cxDefault As Long
  cxIdeal As Long
'#End If
End Type
Public Enum LVCOLUMN_mask
  LVCF_FMT = &H1
  LVCF_WIDTH = &H2
  LVCF_TEXT = &H4
  LVCF_SUBITEM = &H8
'#If (WIN32_IE >= &H300) Then
  LVCF_IMAGE = &H10
  LVCF_ORDER = &H20
'#End If
'#If (WIN32_IE >= &H600) Then
  LVCF_MINWIDTH = &H40
  LVCF_DEFAULTWIDTH = &H80
  LVCF_IDEALWIDTH = &H100
'#End If
End Enum
 
Public Enum LVCOLUMN_fmt
  LVCFMT_LEFT = &H0
  LVCFMT_RIGHT = &H1
  LVCFMT_CENTER = &H2
  LVCFMT_JUSTIFYMASK = &H3
'#If (WIN32_IE >= &H300) Then
  LVCFMT_IMAGE = &H800
  LVCFMT_BITMAP_ON_RIGHT = &H1000
  LVCFMT_COL_HAS_IMAGES = &H8000&
'#End If
'#If (WIN32_IE >= &H600) Then
  LVCFMT_FIXED_WIDTH = &H100
  LVCFMT_NO_DPI_SCALE = &H40000
  LVCFMT_FIXED_RATIO = &H80000
  LVCFMT_LINE_BREAK = &H100000
  LVCFMT_FILL = &H200000
  LVCFMT_WRAP = &H400000
  LVCFMT_NO_TITLE = &H800000
  LVCFMT_TILE_PLACEMENTMASK = (LVCFMT_LINE_BREAK Or LVCFMT_FILL)
  LVCFMT_SPLITBUTTON = &H1000000
'#End If
End Enum

Public Enum LVM_SETCOLUMNWIDTH_lParam
  LVSCW_AUTOSIZE = -1
  LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Public Enum LVGROUPRECT
    LVGGR_GROUP = 0                      'Entire expanded group
    LVGGR_HEADER = 1                     'Header only (collapsed group)
    LVGGR_LABEL = 2                      'Label only
    LVGGR_SUBSETLINK = 3                 'subset link only
End Enum
Public Enum LVGROUPMETRICFLAGS
    LVGMF_NONE = 0
    LVGMF_BORDERSIZE = 1
    LVGMF_BORDERCOLOR = 2
    LVGMF_TEXTCOLOR = 4
End Enum
Public Enum LVGROUPMASK
     LVGF_NONE = 0
     LVGF_HEADER = &H1
     LVGF_FOOTER = &H2
     LVGF_STATE = &H4
     LVGF_ALIGN = &H8
     LVGF_GROUPID = &H10
    ' If SO >= WinVista Then
     LVGF_SUBTITLE = &H100
     LVGF_TASK = &H200
     LVGF_DESCRIPTIONTOP = &H400
     LVGF_DESCRIPTIONBOTTOM = &H800
     LVGF_TITLEIMAGE = &H1000
     LVGF_EXTENDEDIMAGE = &H2000
     LVGF_ITEMS = &H4000
     LVGF_SUBSET = &H8000
     LVGF_SUBSETITEMS = &H10000               'readonly, cItems holds count of items in visible subset, iFirstItem is valid
End Enum

Public Enum LVGROUPSTATE
     LVGS_NORMAL = &H0
     LVGS_COLLAPSED = &H1
     LVGS_HIDDEN = &H2
    
    ' SO >= WinVista
     LVGS_NOHEADER = &H4
     LVGS_COLLAPSIBLE = &H8
     LVGS_FOCUSED = &H10
     LVGS_SELECTED = &H20
     LVGS_SUBSETED = &H40
     LVGS_SUBSETLINKFOCUSED = &H80
End Enum
Public Enum LVGROUPALIGN
     LVGA_HEADER_LEFT = &H1
     LVGA_HEADER_CENTER = &H2
     LVGA_HEADER_RIGHT = &H4             ' Don't forget to validate exclusivity
    ' SO >= WinVista
     LVGA_FOOTER_LEFT = &H8
     LVGA_FOOTER_CENTER = &H10
     LVGA_FOOTER_RIGHT = &H20             ' Don't forget to validate exclusivity
End Enum

Public Type LVGROUP
    cbSize                  As Long
    mask                    As LVGROUPMASK
    pszHeader               As Long
    cchHeader               As Long
    
    pszFooter               As Long
    cchFooter               As Long
    
    iGroupId                As Long
    
    StateMask               As LVGROUPSTATE
    State                   As LVGROUPSTATE
    uAlign                  As LVGROUPALIGN
' SO >= WinVista
    pszSubtitle            As Long
    cchSubtitle            As Long
    pszTask                As Long
    cchTask                As Long
    pszDescriptionTop      As Long
    cchDescriptionTop      As Long
    pszDescriptionBottom   As Long
    cchDescriptionBottom   As Long
    iTitleImage            As Long
    iExtendedImage         As Long
    iFirstItem             As Long     ' Read only
    cItems                 As Long     ' Read only
    pszSubsetTitle         As Long   ' NULL if group is not subset
    cchSubsetTitle         As Long
End Type
Public Type LVINSERTGROUPSORTED
    pfnGroupCompare As Long
    pvData As Long
    LVG As LVGROUP
End Type

Public Type LVGROUPMETRICS
    cbSize      As Long
    mask        As LVGROUPMETRICFLAGS
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
    crLeft      As Long
    crTop       As Long
    crRigth     As Long
    crBottom    As Long
    crHeader    As Long
    crFooter    As Long
End Type
' Notify Message Header for Listview


Public Type NMHEADER
     hdr As NMHDR
     iItem As Long
     iButton As Long
     lPtrHDItem As Long ' HDITEM FAR* pItem
End Type
Public Type NMLISTVIEW   ' was NM_LISTVIEW
  hdr As NMHDR
  iItem As Long
  iSubItem As Long
  uNewState As Long
  uOldState As Long
  uChanged As Long
  PTAction As POINT
  lParam As Long
End Type
Public Enum LVCD_ItemType
    LVCDI_ITEM = &H0
    LVCDI_GROUP = &H1
    LVCDI_ITEMSLIST = &H2
End Enum
Public Const LVCDRF_NOSELECT = &H10000
Public Const LVCDRF_NOGROUPFRAME = &H20000
  Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
  End Type

Public Type NMLVCUSTOMDRAW
  NMCD As NMCUSTOMDRAW
  ClrText As Long
  ClrTextBk As Long
  ' if IE >= 4.0 this member of the struct can be used
  iSubItem As Integer
  '>=5.01
  dwItemType As LVCD_ItemType
  clrFace As Long
  iIconEffect As Integer
  iIconPhase As Integer
  iPartId As Integer
  iStateId As Integer
  rcText As RECT
  uAlign As Long
End Type
Public Type NMLVKEYDOWN   ' was LV_KEYDOWN
   hdr As NMHDR
   wVKey As Integer   ' can't be KeyCodeConstants, enums are Longs!
   Flags As Long   ' Always zero.
End Type
Public Type NMLVDISPINFO   ' was LV_DISPINFO
  hdr As NMHDR
  Item As LVITEM
End Type
Private Const L_MAX_URL_LENGTH = 2084
Private Const MAX_LINKID_TEXT = 48
Public Type LITEM
    mask As Long
    iLink As Long
    State As Long
    StateMask As Long
    szID(0 To ((MAX_LINKID_TEXT * 2) - 1)) As Byte
    szURL(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type
Public Type NMLVLINK
    hdr As NMHDR
    Item As LITEM
    iItem As Long
    iGroupId As Long
End Type

Public Const EMF_CENTERED = &H1
Public Type NMLVEMPTYMARKUP
    hdr As NMHDR
    dwFlags As Long
    szMarkup(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type

Public Type NMLVSCROLL
    hdr As NMHDR
    DX As Long
    DY As Long
End Type

Public Type NMLVGROUP
    hdr As NMHDR
    iGroupId As Long
    uNewState As Long
    uOldState As Long
End Type

Public Type NMLVODSTATECHANGE
    hdr As NMHDR
    iFrom As Long
    iTo As Long
    uNewState As Long
    uOldState As Long
End Type
Public Const LVGIT_UNFOLDED = &H1
Public Type NMLVGETINFOTIP
    hdr As NMHDR
    dwFlags As Long
    pszText As Long
    cchTextMax As Long
    iItem As Long
    iSubItem As Long
    lParam As Long
End Type

Public Type NMLVFINDITEM
    hdr As NMHDR
    iStart As Long
    LVFI As LVFINDINFO
End Type

Public Type NMLVCACHEHINT
    hdr As NMHDR
    iFrom As Long
    iTo As Long
End Type

Public Type NMITEMACTIVATE
    hdr As NMHDR
    iItem As Long
    iSubItem As Long
    uNewState As Long
    uOldState As Long
    uChanged As Long
    PTAction As POINT
    lParam As Long
    uKeyFlags As Long
End Type

'ListView Rating Column - will require lvrHIML to be setup
Public lvrEnable As Boolean
Public lvrHIML As Long
Public lvrCol As Long
Public lvrData() As Long
' ============================================================
' listview macros
Public Function ListView_ApproximateViewRect(hWnd As Long, iWidth As Long, _
                                                                      iHeight As Long, iCount As Long) As Long
  ListView_ApproximateViewRect = SendMessage(hWnd, _
                                                                          LVM_APPROXIMATEVIEWRECT, _
                                                                          ByVal iCount, _
                                                                          ByVal MAKELPARAM(iWidth, iHeight))
End Function
Public Function ListView_Arrange(hwndLV As Long, code As LVA_Flags) As Boolean
  ListView_Arrange = SendMessage(hwndLV, LVM_ARRANGE, ByVal code, 0)
End Function
Public Function ListView_CreateDragImage(hWnd As Long, i As Long, lpptUpLeft As POINT) As Long
  ListView_CreateDragImage = SendMessage(hWnd, LVM_CREATEDRAGIMAGE, ByVal i, lpptUpLeft)
End Function
Public Function ListView_DeleteItem(hWnd As Long, i As Long) As Boolean
  ListView_DeleteItem = SendMessage(hWnd, LVM_DELETEITEM, ByVal i, 0)
End Function
Public Function ListView_EditLabel(hwndLV As Long, i As Long) As Long
  ListView_EditLabel = SendMessage(hwndLV, LVM_EDITLABEL, ByVal i, 0)
End Function
Public Function ListView_GetBkColor(hWnd As Long) As Long
  ListView_GetBkColor = SendMessage(hWnd, LVM_GETBKCOLOR, 0, 0)
End Function
 
Public Function ListView_SetBkColor(hWnd As Long, clrBk As Long) As Boolean
  ListView_SetBkColor = SendMessage(hWnd, LVM_SETBKCOLOR, 0, ByVal clrBk)
End Function
Public Function ListView_SetWorkAreas(hWnd As Long, nWorkAreas As Long, prc() As RECT) As Boolean
  ListView_SetWorkAreas = SendMessage(hWnd, LVM_SETWORKAREAS, ByVal nWorkAreas, prc(0))
End Function

Public Function ListView_GetWorkAreas(hWnd As Long, nWorkAreas, prc() As RECT) As Boolean
  ListView_GetWorkAreas = SendMessage(hWnd, LVM_GETWORKAREAS, ByVal nWorkAreas, prc(0))
End Function

Public Function ListView_GetNumberOfWorkAreas(hWnd As Long, pnWorkAreas As Long) As Boolean
  ListView_GetNumberOfWorkAreas = SendMessage(hWnd, LVM_GETNUMBEROFWORKAREAS, 0, pnWorkAreas)
End Function

Public Function ListView_GetSelectionMark(hWnd As Long) As Long
  ListView_GetSelectionMark = SendMessage(hWnd, LVM_GETSELECTIONMARK, 0, 0)
End Function

Public Function ListView_SetSelectionMark(hWnd As Long, i As Long) As Long
  ListView_SetSelectionMark = SendMessage(hWnd, LVM_SETSELECTIONMARK, 0, ByVal i)
End Function

Public Function ListView_SetHoverTime(hwndLV As Long, dwHoverTimeMs As Long) As Long
  ListView_SetHoverTime = SendMessage(hwndLV, LVM_SETHOVERTIME, 0, ByVal dwHoverTimeMs)
End Function

Public Function ListView_GetHoverTime(hwndLV As Long) As Long
  ListView_GetHoverTime = SendMessage(hwndLV, LVM_GETHOVERTIME, 0, 0)
End Function
Public Function ListView_GetStringWidth(hwndLV As Long, psz As String) As Long
  ListView_GetStringWidth = SendMessage(hwndLV, LVM_GETSTRINGWIDTH, 0, ByVal psz)
End Function
 Public Function ListView_GetSubItemRect(hWnd As Long, iItem As Long, iSubItem As Long, _
                                                              code As LVIR_Flags, prc As RECT) As Boolean
  prc.Top = iSubItem
  prc.Left = code
  ListView_GetSubItemRect = SendMessage(hWnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
End Function
Public Function ListView_GetTextBkColor(hWnd As Long) As Long
  ListView_GetTextBkColor = SendMessage(hWnd, LVM_GETTEXTBKCOLOR, 0, 0)
End Function
 
Public Function ListView_SetTextBkColor(hWnd As Long, ClrTextBk As Long) As Boolean
  ListView_SetTextBkColor = SendMessage(hWnd, LVM_SETTEXTBKCOLOR, 0, ByVal ClrTextBk)
End Function
Public Function ListView_GetTextColor(hWnd As Long) As Long
  ListView_GetTextColor = SendMessage(hWnd, LVM_GETTEXTCOLOR, 0, 0)
End Function
 
Public Function ListView_SetTextColor(hWnd As Long, ClrText As Long) As Boolean
  ListView_SetTextColor = SendMessage(hWnd, LVM_SETTEXTCOLOR, 0, ByVal ClrText)
End Function
Public Function ListView_GetTopIndex(hwndLV As Long) As Long
  ListView_GetTopIndex = SendMessage(hwndLV, LVM_GETTOPINDEX, 0, 0)
End Function
 
Public Function ListView_SubItemHitTest(hWnd As Long, plvhti As LVHITTESTINFO) As Long
  ListView_SubItemHitTest = SendMessage(hWnd, LVM_SUBITEMHITTEST, 0, plvhti)
End Function


Public Function ListView_SetToolTips(hwndLV As Long, hwndNewHwnd As Long) As Long
  ListView_SetToolTips = SendMessage(hwndLV, LVM_SETTOOLTIPS, ByVal hwndNewHwnd, 0)
End Function

Public Function ListView_GetToolTips(hwndLV As Long) As Long
  ListView_GetToolTips = SendMessage(hwndLV, LVM_GETTOOLTIPS, 0, 0)
End Function
Public Function ListView_GetISearchString(hwndLV As Long, lpsz As String) As Boolean
  ListView_GetISearchString = SendMessage(hwndLV, LVM_GETISEARCHSTRING, 0, ByVal lpsz)
End Function


Public Function ListView_SetBkImage(hWnd As Long, plvbki As LVBKIMAGE) As Boolean
  ListView_SetBkImage = SendMessage(hWnd, LVM_SETBKIMAGE, 0, plvbki)
End Function

Public Function ListView_GetBkImage(hWnd As Long, plvbki As LVBKIMAGE) As Boolean
  ListView_GetBkImage = SendMessage(hWnd, LVM_GETBKIMAGE, 0, plvbki)
End Function
Public Function ListView_SetUnicodeFormat(hWnd As Long, fUnicode As Boolean) As Boolean
  ListView_SetUnicodeFormat = SendMessage(hWnd, LVM_SETUNICODEFORMAT, ByVal fUnicode, 0)
End Function

Public Function ListView_GetUnicodeFormat(hWnd As Long) As Boolean
  ListView_GetUnicodeFormat = SendMessage(hWnd, LVM_GETUNICODEFORMAT, 0, 0)
End Function

Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, dwMask As Long, dw As Long) As Long
  ListView_SetExtendedListViewStyleEx = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                                                                                    ByVal dwMask, ByVal dw)
End Function

Public Function ListView_SetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
  ListView_SetColumnOrderArray = SendMessage(hWnd, LVM_SETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function

Public Function ListView_GetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
  ListView_GetColumnOrderArray = SendMessage(hWnd, LVM_GETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function
Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As LV_ImageList) As Long
  ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
End Function
Public Function ListView_GetImageList(hWnd As Long, iImageList As LV_ImageList) As Long
  ListView_GetImageList = SendMessage(hWnd, LVM_GETIMAGELIST, ByVal iImageList, 0)
End Function
 
Public Function ListView_GetHeader(hWnd As Long) As Long
  ListView_GetHeader = SendMessage(hWnd, LVM_GETHEADER, 0, 0)
End Function
Public Function ListView_GetItem(hWnd As Long, pItem As LVITEM) As Boolean
  ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pItem)
End Function
 
Public Function ListView_SetItem(hWnd As Long, pItem As LVITEM) As Boolean
  ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pItem)
End Function
'
Public Function ListView_SetCallbackMask(hWnd As Long, mask As Long) As Boolean
  ListView_SetCallbackMask = SendMessage(hWnd, LVM_SETCALLBACKMASK, ByVal mask, 0)
End Function
Public Function ListView_GetCallbackMask(hWnd As Long) As Long   ' LVStyles
  ListView_GetCallbackMask = SendMessage(hWnd, LVM_GETCALLBACKMASK, 0, 0)
End Function
Public Function ListView_GetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
  ListView_GetColumn = SendMessage(hWnd, LVM_GETCOLUMN, ByVal iCol, pcol)
End Function
 
Public Function ListView_SetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
  ListView_SetColumn = SendMessage(hWnd, LVM_SETCOLUMN, ByVal iCol, pcol)
End Function
Public Function ListView_GetCountPerPage(hwndLV As Long) As Long
  ListView_GetCountPerPage = SendMessage(hwndLV, LVM_GETCOUNTPERPAGE, 0, 0)
End Function
 
Public Function ListView_GetOrigin(hwndLV As Long, ppt As POINT) As Boolean
  ListView_GetOrigin = SendMessage(hwndLV, LVM_GETORIGIN, 0, ppt)
End Function
Public Function ListView_GetEditControl(hwndLV As Long) As Long
  ListView_GetEditControl = SendMessage(hwndLV, LVM_GETEDITCONTROL, 0, 0)
End Function
Public Function ListView_GetExtendedListViewStyle(hwndLV As Long) As Long
  ListView_GetExtendedListViewStyle = SendMessage(hwndLV, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
End Function
Public Function ListView_SetHotItem(hWnd As Long, i As Long) As Long
  ListView_SetHotItem = SendMessage(hWnd, LVM_SETHOTITEM, ByVal i, 0)
End Function
 
Public Function ListView_GetHotItem(hWnd As Long) As Long
  ListView_GetHotItem = SendMessage(hWnd, LVM_GETHOTITEM, 0, 0)
End Function
 
Public Function ListView_SetHotCursor(hWnd As Long, hcur As Long) As Long
  ListView_SetHotCursor = SendMessage(hWnd, LVM_SETHOTCURSOR, 0, ByVal hcur)
End Function
 
Public Function ListView_GetHotCursor(hWnd As Long) As Long
  ListView_GetHotCursor = SendMessage(hWnd, LVM_GETHOTCURSOR, 0, 0)
End Function

Public Sub ListView_SetItemText(hwndLV As Long, i As Long, iSubItem As Long, pszText As String)
  Dim lvi As LVITEM
  lvi.iSubItem = iSubItem
  lvi.pszText = pszText
  lvi.cchTextMax = Len(pszText) + 1
  SendMessage hwndLV, LVM_SETITEMTEXT, ByVal i, lvi
End Sub
Public Function ListView_SetIconSpacing(hwndLV As Long, CX As Long, CY As Long) As Long
  ListView_SetIconSpacing = SendMessage(hwndLV, LVM_SETICONSPACING, 0, ByVal MAKELONG(CX, CY))
End Function
Public Sub ListView_SetItemCount(hwndLV As Long, cItems As Long)
  SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, 0
End Sub

#If (WIN32_IE >= &H300) Then

Public Sub ListView_SetItemCountEx(hwndLV As Long, cItems As Long, dwFlags As Long)
  SendMessage hwndLV, LVM_SETITEMCOUNT, ByVal cItems, ByVal dwFlags
End Sub
'
#End If

' ListView_GetNextItem

Public Function ListView_GetNextItem(hWnd As Long, i As Long, Flags As LVNI_Flags) As Long
  ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal Flags)    ' ByVal MAKELPARAM(flags, 0))
End Function

' Returns the index of the item that is selected and has the focus rectangle (user-defined)

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
  ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function
Public Function ListView_FindItem(hWnd As Long, iStart, plvfi As LVFINDINFO) As Long
  ListView_FindItem = SendMessage(hWnd, LVM_FINDITEM, ByVal iStart, plvfi)
End Function
Public Function ListView_GetItemRect(hWnd As Long, i As Long, prc As RECT, code As LVIR_Flags) As Boolean
  prc.Left = code
  ListView_GetItemRect = SendMessage(hWnd, LVM_GETITEMRECT, ByVal i, prc)
End Function
Public Function ListView_GetCheckState(hwndLV As Long, iIndex As Long) As Long   ' updated
  Dim dwState As Long
  dwState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal iIndex, ByVal LVIS_STATEIMAGEMASK)
  ListView_GetCheckState = (dwState \ 2 ^ 12) - 1
  '((((UINT)(SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, LVIS_STATEIMAGEMASK))) >> 12) -1)
End Function
Public Function ListView_SetCheckState(hwndLV As Long, i As Long, fCheck As Long) As Long
'#define ListView_SetCheckState(hwndLV, i, fCheck) \
'  ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK((fCheck)?2:1), LVIS_STATEIMAGEMASK)
ListView_SetCheckState = ListView_SetItemState(hwndLV, i, IndexToStateImageMask(IIf(fCheck, 2, 1)), LVIS_STATEIMAGEMASK)
End Function

Public Function ListView_GetItemCount(hWnd As Long) As Long
  ListView_GetItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function
Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINT) As Boolean
  ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
End Function
Public Function ListView_SetItemPosition(hwndLV As Long, i As Long, x As Long, y As Long) As Boolean
  ListView_SetItemPosition = SendMessage(hwndLV, LVM_SETITEMPOSITION, ByVal i, ByVal MAKELPARAM(x, y))
End Function
Public Sub ListView_SetItemPosition32(hwndLV As Long, i As Long, x As Long, y As Long)
  Dim ptNewPos As POINT
  ptNewPos.x = x
  ptNewPos.y = y
  SendMessage hwndLV, LVM_SETITEMPOSITION32, ByVal i, ptNewPos
End Sub
Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
  ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                                                                     LVIS_FOCUSED Or LVIS_SELECTED)
End Function
Public Function ListView_Update(hwndLV As Long, i As Long) As Boolean
  ListView_Update = SendMessage(hwndLV, LVM_UPDATE, ByVal i, 0)
End Function

Public Function ListView_GetItemSpacing(hwndLV As Long, fSmall As Boolean) As Long
  ListView_GetItemSpacing = SendMessage(hwndLV, LVM_GETITEMSPACING, ByVal fSmall, 0)
End Function
Public Function ListView_GetItemState(hwndLV As Long, i As Long, mask As LVITEM_state) As Long   ' LVITEM_state
  ListView_GetItemState = SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, ByVal mask)
End Function
Public Sub ListView_GetItemText(hwndLV As Long, i As Long, iSubItem As Long, _
                                                     pszText As Long, cchTextMax As Long)
  Dim lvi As LVITEM
  lvi.iSubItem = iSubItem
  lvi.cchTextMax = cchTextMax
  lvi.pszText = pszText
  SendMessage hwndLV, LVM_GETITEMTEXT, ByVal i, lvi
  pszText = lvi.pszText   ' fills pszText w/ pointer
End Sub


Public Function ListView_HitTest(hwndLV As Long, pInfo As LVHITTESTINFO) As Long
  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pInfo)
End Function
 
Public Function ListView_InsertItem(hWnd As Long, pItem As LVITEM) As Long
  ListView_InsertItem = SendMessage(hWnd, LVM_INSERTITEM, 0, pItem)
End Function

Public Function ListView_DeleteColumn(hWnd As Long, iCol As Long) As Boolean
  ListView_DeleteColumn = SendMessage(hWnd, LVM_DELETECOLUMN, ByVal iCol, 0)
End Function

Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As CBoolean) As Boolean
  ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal fPartialOK)   ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_InsertColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Long
  ListView_InsertColumn = SendMessage(hWnd, LVM_INSERTCOLUMN, ByVal iCol, pcol)
End Function
Public Function ListView_Scroll(hwndLV As Long, DX As Long, DY As Long) As Boolean
  ListView_Scroll = SendMessage(hwndLV, LVM_SCROLL, ByVal DX, ByVal DY)
End Function
 

 Public Function ListView_DeleteAllItems(hWnd As Long) As Boolean
  ListView_DeleteAllItems = SendMessage(hWnd, LVM_DELETEALLITEMS, 0, 0)
End Function

Public Function ListView_GetColumnWidth(hWnd As Long, iCol As Long) As Long
  ListView_GetColumnWidth = SendMessage(hWnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0)
End Function
 
Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, CX As Long) As Boolean
  ListView_SetColumnWidth = SendMessage(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(CX, 0))
End Function
Public Function ListView_RedrawItems(hwndLV As Long, iFirst, iLast) As Boolean
  ListView_RedrawItems = SendMessage(hwndLV, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast)
End Function

Public Function ListView_GetSelectedCount(hwndLV As Long) As Long
  ListView_GetSelectedCount = SendMessage(hwndLV, LVM_GETSELECTEDCOUNT, 0, 0)
End Function
Public Function ListView_GetView(hWnd As Long) As Long

ListView_GetView = SendMessage(hWnd, LVM_GETVIEW, 0, ByVal 0&)

End Function
Public Function ListView_GetViewRect(hWnd As Long, prc As RECT) As Boolean
  ListView_GetViewRect = SendMessage(hWnd, LVM_GETVIEWRECT, 0, prc)
End Function
' ListView_SetItemState

Public Function ListView_SetItemState(hwndLV As Long, i As Long, State As Long, mask As Long) As Boolean
  Dim lvi As LVITEM
  lvi.State = State
  lvi.StateMask = mask
  ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

' Selects all listview items. The item with the focus rectangle maintains it (user-defined).

Public Function ListView_SelectAll(hwndLV As Long) As Boolean
  ListView_SelectAll = ListView_SetItemState(hwndLV, -1, LVIS_SELECTED, LVIS_SELECTED)
End Function
Public Function ListView_SelectNone(hwndLV As Long) As Boolean
  Dim lv As LVITEM
   
   With lv
      .mask = LVIF_STATE
      .State = False
      .StateMask = LVIS_SELECTED
   End With
      
   ListView_SelectNone = SendMessage(hwndLV, LVM_SETITEMSTATE, -1, lv)

End Function
 
' Selects the specified item and gives it the focus rectangle.
' does not de-select any currently selected items (user-defined).

Public Function ListView_SetFocusedItem(hwndLV As LongPtr, i As Long) As Boolean
  ListView_SetFocusedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)
End Function

Public Function ListView_SortItems(hwndLV As LongPtr, pfnCompare As Long, lParamSort As Long) As Boolean
  ListView_SortItems = SendMessage(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
End Function
Public Function ListView_SortItemsEx(hwndLV As LongPtr, pfnCompare As Long, lParamSort As Long) As Boolean
  ListView_SortItemsEx = SendMessage(hwndLV, LVM_SORTITEMSEX, ByVal lParamSort, ByVal pfnCompare)
End Function



Public Function ListView_SetExtendedStyle(hWnd As LongPtr, lST As LVStylesEx) As Long
    Dim lStyle As Long
    
    lStyle = SendMessage(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or lST
    Call SendMessage(hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)
End Function

Public Function ListView_GetStyle(hWnd As LongPtr) As LVStyles
    ListView_GetStyle = GetWindowLongPtr(hWnd, GWL_STYLE)
End Function

Public Function ListView_SetStyle(hWnd As LongPtr, dwStyle As LVStyles) As LongPtr
    ListView_SetStyle = SetWindowLongPtr(hWnd, GWL_STYLE, dwStyle)
End Function

'THE MACROS BELOW ARE ONLY FOR VISTA AND HIGHER
Public Function ListView_CancelEditLabel(hWnd As LongPtr) As LongPtr
    ListView_CancelEditLabel = SendMessage(hWnd, LVM_CANCELEDITLABEL, 0, ByVal 0&)
End Function

Public Function ListView_EnableGroupView(hWnd As LongPtr, fEnable As Long) As LongPtr
    ListView_EnableGroupView = SendMessage(hWnd, LVM_ENABLEGROUPVIEW, fEnable, ByVal 0&)
End Function

Public Function ListView_GetEmptyText(hWnd As LongPtr, cchText As Long, pszText As String) As LongPtr
    ListView_GetEmptyText = SendMessage(hWnd, LVM_GETEMPTYTEXT, cchText, ByVal pszText)
End Function

Public Function ListView_GetFocusedGroup(hWnd As LongPtr) As LongPtr
    '#define ListView_GetFocusedGroup(hwnd) \
    '    SNDMSG((hwnd), LVM_GETFOCUSEDGROUP, 0, 0)
    ListView_GetFocusedGroup = SendMessage(hWnd, LVM_GETFOCUSEDGROUP, 0, ByVal 0&)
End Function

Public Function ListView_GetFooterInfo(hWnd As LongPtr, plvfi As Long) As LongPtr
    '#define ListView_GetFooterInfo(hwnd, plvfi) \
    '    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERINFO, (WPARAM)(0), (LPARAM)(plvfi))
    ListView_GetFooterInfo = SendMessage(hWnd, LVM_GETFOOTERINFO, 0, ByVal plvfi)
End Function

Public Function ListView_GetFooterItem(hWnd As LongPtr, iItem As Long, pfi As LVFOOTERITEM) As LongPtr
    '#define ListView_GetFooterItem(hwnd, iItem, pfi) \
    '    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEM, (WPARAM)(iItem), (LPARAM)(pfi))
    ListView_GetFooterItem = SendMessage(hWnd, LVM_GETFOOTERITEM, iItem, pfi)
End Function

Public Function ListView_GetFooterItemRect(hWnd As LongPtr, iItem As Long, prc As RECT) As LongPtr
    '#define ListView_GetFooterItemRect(hwnd, iItem, prc) \
    '    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERITEMRECT, (WPARAM)(iItem), (LPARAM)(prc))
    ListView_GetFooterItemRect = SendMessage(hWnd, LVM_GETFOOTERITEMRECT, iItem, prc)
End Function

Public Function ListView_GetFooterRect(hWnd As LongPtr, prc As RECT) As LongPtr
    '#define ListView_GetFooterRect(hwnd, prc) \
    '    (BOOL)SNDMSG((hwnd), LVM_GETFOOTERRECT, (WPARAM)(0), (LPARAM)(prc))
    ListView_GetFooterRect = SendMessage(hWnd, LVM_GETFOOTERRECT, 0, prc)
End Function

Public Function ListView_GetGroupHeaderImageList(hWnd As LongPtr) As LongPtr
    '#define ListView_GetGroupHeaderImageList(hwnd) \
    '    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, 0L)
    ListView_GetGroupHeaderImageList = SendMessage(hWnd, LVM_GETIMAGELIST, LVSIL_GROUPHEADER, ByVal 0&)
End Function

Public Function ListView_SetGroupHeaderImageList(hWnd As LongPtr, himl As Long) As LongPtr
    '#define ListView_GetGroupHeaderImageList(hwnd) \
    '    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)LVSIL_GROUPHEADER, 0L)
    ListView_SetGroupHeaderImageList = SendMessage(hWnd, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal himl)
End Function
Public Function ListView_GetGroupInfo(hWnd As Long, iGroupId As Long, pgrp As LVGROUP) As Long
'#define ListView_GetGroupInfo(hwnd, iGroupId, pgrp) \
'    SNDMSG((hwnd), LVM_GETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
ListView_GetGroupInfo = SendMessage(hWnd, LVM_GETGROUPINFO, iGroupId, pgrp)
End Function
Public Function ListView_SetGroupInfo(hWnd As Long, iGroupId As Long, pgrp As LVGROUP) As Long
'#define ListView_SetGroupInfo(hwnd, iGroupId, pgrp) \
'    SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)(iGroupId), (LPARAM)(pgrp))
ListView_SetGroupInfo = SendMessage(hWnd, LVM_SETGROUPINFO, iGroupId, pgrp)
End Function
Public Function ListView_GetGroupInfoByIndex(hWnd As Long, iIndex As Long, pgrp As LVGROUP) As Long
'#define ListView_GetGroupInfoByIndex(hwnd, iIndex, pgrp) \
'    SNDMSG((hwnd), LVM_GETGROUPINFOBYINDEX, (WPARAM)(iIndex), (LPARAM)(pgrp))
ListView_GetGroupInfoByIndex = SendMessage(hWnd, LVM_GETGROUPINFOBYINDEX, iIndex, pgrp)
End Function
Public Function ListView_SetGroupMetrics(hWnd As Long, pGroupMetrics As LVGROUPMETRICS) As Long
'#define ListView_SetGroupMetrics(hwnd, pGroupMetrics) \
'    SNDMSG((hwnd), LVM_SETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
ListView_SetGroupMetrics = SendMessage(hWnd, LVM_SETGROUPMETRICS, 0, pGroupMetrics)
End Function
Public Function ListView_GetGroupMetrics(hWnd As Long, pGroupMetrics As LVGROUPMETRICS) As Long
'#define ListView_GetGroupMetrics(hwnd, pGroupMetrics) \
'    SNDMSG((hwnd), LVM_GETGROUPMETRICS, 0, (LPARAM)(pGroupMetrics))
ListView_GetGroupMetrics = SendMessage(hWnd, LVM_GETGROUPMETRICS, 0, pGroupMetrics)
End Function
Public Function ListView_GetGroupRect(hWnd As Long, iGroup As Long, Item As LVGROUPRECT, rc As RECT) As Long
        rc.Top = Item
        ListView_GetGroupRect = SendMessage(hWnd, LVM_GETGROUPRECT, iGroup, rc)
End Function
Public Function ListView_GetGroupCount(hWnd As Long, iGroup As Long) As Long
Dim LVG As LVGROUP
    LVG.mask = LVGF_ITEMS
    LVG.cbSize = LenB(LVG)
    Call SendMessage(hWnd, LVM_GETGROUPINFO, iGroup, LVG)
ListView_GetGroupCount = LVG.cItems
End Function
Public Function ListView_GetGroupState(hWnd As Long, dwGroupId As Long, dwMask As Long) As Long
ListView_GetGroupState = SendMessage(hWnd, LVM_GETGROUPSTATE, dwGroupId, ByVal dwMask)
End Function

'#define ListView_SetGroupState(hwnd, dwGroupId, dwMask, dwState) \
'{ LVGROUP _macro_lvg;\
'  _macro_lvg.cbSize = sizeof(_macro_lvg);\
'  _macro_lvg.mask = LVGF_STATE;\
'  _macro_lvg.stateMask = dwMask;\
'  _macro_lvg.state = dwState;\
'  SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)(dwGroupId), (LPARAM)(LVGROUP *)&_macro_lvg);\
Public Function ListView_SetGroupState(hWnd As Long, dwGroupId As Long, dwMask As Long, dwState As Long) As Long
Dim LVG As LVGROUP
LVG.cbSize = LenB(LVG)
LVG.mask = LVGF_STATE
LVG.StateMask = dwMask
LVG.State = dwState
ListView_SetGroupState = SendMessage(hWnd, LVM_SETGROUPINFO, dwGroupId, LVG)
End Function
Public Function ListView_GetInsertMark(hWnd As Long, LVIM As LVINSERTMARK) As Long
'#define ListView_GetInsertMark(hwnd, lvim) \
'    (BOOL)SNDMSG((hwnd), LVM_GETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
ListView_GetInsertMark = SendMessage(hWnd, LVM_GETINSERTMARK, 0, LVIM)
End Function
Public Function ListView_SetInsertMark(hWnd As Long, LVIM As LVINSERTMARK) As Long
'#define ListView_SetInsertMark(hwnd, lvim) \
'    (BOOL)SNDMSG((hwnd), LVM_SETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))
ListView_SetInsertMark = SendMessage(hWnd, LVM_SETINSERTMARK, 0, LVIM)
End Function
Public Function ListView_GetInsertMarkColor(hWnd As Long) As Long
'#define ListView_GetInsertMarkColor(hwnd) \
'    (COLORREF)SNDMSG((hwnd), LVM_GETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)0)
ListView_GetInsertMarkColor = SendMessage(hWnd, LVM_GETINSERTMARKCOLOR, 0, ByVal 0&)
End Function
Public Function ListView_SetInsertMarkColor(hWnd As Long, Color As Long) As Long
'#define ListView_SetInsertMarkColor(hwnd, color) \
'    (COLORREF)SNDMSG((hwnd), LVM_SETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
ListView_SetInsertMarkColor = SendMessage(hWnd, LVM_SETINSERTMARKCOLOR, 0, ByVal Color)
End Function
Public Function ListView_GetInsertMarkRect(hWnd As Long, rc As RECT) As Long
'#define ListView_GetInsertMarkRect(hwnd, rc) \
'    (int)SNDMSG((hwnd), LVM_GETINSERTMARKRECT, (WPARAM)0, (LPARAM)(LPRECT)(rc))
ListView_GetInsertMarkRect = SendMessage(hWnd, LVM_GETINSERTMARKRECT, 0, rc)
End Function
Public Function ListView_InsertMarkHitTest(hWnd As Long, POINT As POINT, LVIM As LVINSERTMARK) As Long
'#define ListView_InsertMarkHitTest(hwnd, point, lvim) \
'    (int)SNDMSG((hwnd), LVM_INSERTMARKHITTEST, (WPARAM)(LPPOINT)(point), (LPARAM)(LPLVINSERTMARK)(lvim))
ListView_InsertMarkHitTest = SendMessage(hWnd, LVM_INSERTMARKHITTEST, VarPtr(POINT), LVIM)
End Function
Public Function ListView_GetItemIndexRect(hWnd As Long, lvii As LVITEMINDEX, iSubItem As Long, code As Long, prc As RECT) As Long
'#define ListView_GetItemIndexRect(hwnd, plvii, iSubItem, code, prc) \
'        (BOOL)SNDMSG((hwnd), LVM_GETITEMINDEXRECT, (WPARAM)(LVITEMINDEX*)(plvii), \
'                ((prc) ? ((((LPRECT)(prc))->top = (iSubItem)), (((LPRECT)(prc))->left = (code)), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))
prc.Top = iSubItem
prc.Left = code
ListView_GetItemIndexRect = SendMessage(hWnd, LVM_GETITEMINDEXRECT, VarPtr(lvii), prc)
End Function
Public Function ListView_GetNextItemIndex(hWnd As Long, plvii As LVITEMINDEX, ByVal Flags As LVNI_Flags) As Long
 '#define ListView_GetNextItemIndex(hwnd, plvii, flags) \
 '    (BOOL)SNDMSG((hwnd), LVM_GETNEXTITEMINDEX, (WPARAM)(LVITEMINDEX*)(plvii), MAKELPARAM((flags), 0))
 ListView_GetNextItemIndex = SendMessage(hWnd, LVM_GETNEXTITEMINDEX, VarPtr(plvii), ByVal Flags)
End Function
Public Function ListView_GetOutlineColor(hWnd As Long) As Long
'#define ListView_GetOutlineColor(hwnd) \
'    (COLORREF)SNDMSG((hwnd), LVM_GETOUTLINECOLOR, 0, 0)
ListView_GetOutlineColor = SendMessage(hWnd, LVM_GETOUTLINECOLOR, 0, ByVal 0&)
End Function
Public Function ListView_SetOutlineColor(hWnd As Long, Color As Long) As Long
'#define ListView_SetOutlineColor(hwnd, color) \
'    (COLORREF)SNDMSG((hwnd), LVM_SETOUTLINECOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))
ListView_SetOutlineColor = SendMessage(hWnd, LVM_SETOUTLINECOLOR, 0, ByVal Color)
End Function
Public Function ListView_GetSelectedColumn(hWnd As Long) As Long
'#define ListView_GetSelectedColumn(hwnd) \
'    (UINT)SNDMSG((hwnd), LVM_GETSELECTEDCOLUMN, 0, 0)
ListView_GetSelectedColumn = SendMessage(hWnd, LVM_GETSELECTEDCOLUMN, 0, ByVal 0&)
End Function
Public Function ListView_GetTileInfo(hWnd As Long, pTI As LVTILEINFO) As Long
'#define ListView_GetTileInfo(hwnd, pti) \
'    SNDMSG((hwnd), LVM_GETTILEINFO, 0, (LPARAM)(pti))
ListView_GetTileInfo = SendMessage(hWnd, LVM_GETTILEINFO, 0, pTI)
End Function
Public Function ListView_SetTileInfo(hWnd As Long, pTI As LVTILEINFO) As Long
'#define ListView_SetTileInfo(hwnd, pti) \
'    SNDMSG((hwnd), LVM_SETTILEINFO, 0, (LPARAM)(pti))
ListView_SetTileInfo = SendMessage(hWnd, LVM_SETTILEINFO, 0, pTI)
End Function
Public Function ListView_GetTileViewInfo(hWnd As Long, ptvi As LVTILEVIEWINFO) As Long
'#define ListView_GetTileViewInfo(hwnd, ptvi) \
'    SNDMSG((hwnd), LVM_GETTILEVIEWINFO, 0, (LPARAM)(ptvi))
ListView_GetTileViewInfo = SendMessage(hWnd, LVM_GETTILEVIEWINFO, 0, ptvi)
End Function
Public Function ListView_SetTileViewInfo(hWnd As Long, ptvi As LVTILEVIEWINFO) As Long
'#define ListView_SetTileViewInfo(hwnd, ptvi) \
'    SNDMSG((hwnd), LVM_SETTILEVIEWINFO, 0, (LPARAM)(ptvi))
ListView_SetTileViewInfo = SendMessage(hWnd, LVM_SETTILEVIEWINFO, 0, ptvi)
End Function
Public Function ListView_HasGroup(hWnd As Long, dwGroupId As Long) As Long
    ListView_HasGroup = SendMessage(hWnd, LVM_HASGROUP, dwGroupId, ByVal 0&)
End Function
Public Function ListView_HitTestEx(hwndLV As Long, pInfo As LVHITTESTINFO) As Long
'HitTestEx is used if you need the iGroup and iSubItem members filled
  ListView_HitTestEx = SendMessage(hwndLV, LVM_HITTEST, -1, pInfo)
End Function
Public Function ListView_InsertGroup(hWnd As Long, Index As Long, pgrp As LVGROUP) As Long
'#define ListView_InsertGroup(hwnd, index, pgrp) \
'    SNDMSG((hwnd), LVM_INSERTGROUP, (WPARAM)(index), (LPARAM)(pgrp))
ListView_InsertGroup = SendMessage(hWnd, LVM_INSERTGROUP, Index, pgrp)
End Function
Public Function ListView_InsertGroupSorted(hWnd As Long, structInsert As LVINSERTGROUPSORTED) As Long
'#define ListView_InsertGroupSorted(hwnd, structInsert) \
'    SNDMSG((hwnd), LVM_INSERTGROUPSORTED, (WPARAM)(structInsert), 0)
ListView_InsertGroupSorted = SendMessage(hWnd, LVM_INSERTGROUPSORTED, VarPtr(structInsert), ByVal 0&)
End Function
Public Function ListView_IsGroupViewEnabled(hWnd As Long) As Long
'#define ListView_IsGroupViewEnabled(hwnd) \
'    (BOOL)SNDMSG((hwnd), LVM_ISGROUPVIEWENABLED, 0, 0)
ListView_IsGroupViewEnabled = SendMessage(hWnd, LVM_ISGROUPVIEWENABLED, 0, ByVal 0&)
End Function
Public Function ListView_IsItemVisible(hWnd As Long, Index As Long) As Long
'#define ListView_IsItemVisible(hwnd, index) \
'    (UINT)SNDMSG((hwnd), LVM_ISITEMVISIBLE, (WPARAM)(index), (LPARAM)0)
ListView_IsItemVisible = SendMessage(hWnd, LVM_ISITEMVISIBLE, Index, ByVal 0&)
End Function
Public Function ListView_MapIDToIndex(hWnd As Long, id As Long) As Long
'#define ListView_MapIDToIndex(hwnd, id) \
'    (UINT)SNDMSG((hwnd), LVM_MAPIDTOINDEX, (WPARAM)(id), (LPARAM)0)
ListView_MapIDToIndex = SendMessage(hWnd, LVM_MAPIDTOINDEX, id, ByVal 0&)
End Function
Public Function ListView_MapIndexToID(hWnd As Long, Index As Long) As Long
'#define ListView_MapIndexToID(hwnd, index) \
'    (UINT)SNDMSG((hwnd), LVM_MAPINDEXTOID, (WPARAM)(index), (LPARAM)0)
ListView_MapIndexToID = SendMessage(hWnd, LVM_MAPINDEXTOID, Index, ByVal 0&)
End Function
Public Function ListView_MoveGroup(hWnd As Long, iGroupId As Long, toIndex As Long) As Long
'NOT IMPLEMENTED
'#define ListView_MoveGroup(hwnd, iGroupId, toIndex) \
'    SNDMSG((hwnd), LVM_MOVEGROUP, (WPARAM)(iGroupId), (LPARAM)(toIndex))
ListView_MoveGroup = SendMessage(hWnd, LVM_MOVEGROUP, iGroupId, ByVal toIndex)
End Function
Public Function ListView_MoveItemToGroup(hWnd As Long, idItemFrom As Long, idGroupTo As Long) As Long
'NOT IMPLEMENTED
'#define ListView_MoveItemToGroup(hwnd, idItemFrom, idGroupTo) \
'    SNDMSG((hwnd), LVM_MOVEITEMTOGROUP, (WPARAM)(idItemFrom), (LPARAM)(idGroupTo))
ListView_MoveItemToGroup = SendMessage(hWnd, LVM_MOVEITEMTOGROUP, idItemFrom, ByVal idGroupTo)
End Function
Public Function ListView_RemoveAllGroups(hWnd As Long) As Long
'#define ListView_RemoveAllGroups(hwnd) \
'    SNDMSG((hwnd), LVM_REMOVEALLGROUPS, 0, 0)
ListView_RemoveAllGroups = SendMessage(hWnd, LVM_REMOVEALLGROUPS, 0, ByVal 0&)
End Function
Public Function ListView_RemoveGroup(hWnd As Long, iGroupId As Long) As Long
'#define ListView_RemoveGroup(hwnd, iGroupId) \
'    SNDMSG((hwnd), LVM_REMOVEGROUP, (WPARAM)(iGroupId), 0)
ListView_RemoveGroup = SendMessage(hWnd, LVM_REMOVEGROUP, iGroupId, ByVal 0&)

End Function
Public Function ListView_SetInfoTip(hWnd As Long, plvInfoTip As LVSETINFOTIP) As Long
'#define ListView_SetInfoTip(hwndLV, plvInfoTip)\
'        (BOOL)SNDMSG((hwndLV), LVM_SETINFOTIP, (WPARAM)0, (LPARAM)(plvInfoTip))
ListView_SetInfoTip = SendMessage(hWnd, LVM_SETINFOTIP, 0, plvInfoTip)
End Function
Public Function ListView_SetItemIndexState(hwndLV As Long, plvii As LVITEMINDEX, Data As Long, mask As Long) As Long
'#define ListView_SetItemIndexState(hwndLV, plvii, data, mask) \
'{ LV_ITEM _macro_lvi;\
'  _macro_lvi.stateMask = (mask);\
'  _macro_lvi.state = (data);\
'  SNDMSG((hwndLV), LVM_SETITEMINDEXSTATE, (WPARAM)(LVITEMINDEX*)(plvii), (LPARAM)(LV_ITEM *)&_macro_lvi);\}

Dim lvi As LVITEM
lvi.StateMask = mask
lvi.State = Data
ListView_SetItemIndexState = SendMessage(hwndLV, LVM_SETITEMINDEXSTATE, VarPtr(plvii), lvi)
End Function
Public Function ListView_SetSelectedColumn(hWnd As Long, iCol As Long) As Long
'#define ListView_SetSelectedColumn(hwnd, iCol) \
'    SNDMSG((hwnd), LVM_SETSELECTEDCOLUMN, (WPARAM)(iCol), 0)
ListView_SetSelectedColumn = SendMessage(hWnd, LVM_SETSELECTEDCOLUMN, iCol, ByVal 0&)
End Function

Public Function IID_IListView() As UUID
'{E5B16AF2-3990-4681-A609-1F060CD14269}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5B16AF2, CInt(&H3990), CInt(&H4681), &HA6, &H9, &H1F, &H6, &HC, &HD1, &H42, &H69)
 IID_IListView = iid
End Function
Public Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub
Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  '   INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function
Private Function IndexToStateImageMask(ByVal Index As Long) As Long
IndexToStateImageMask = Index * (2 ^ 12)
End Function
Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Private Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
    MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Public Function LoWord(ByVal dwValue As Long) As Integer
' Returns the low 16-bit integer from a 32-bit long integer
    CopyMemory LoWord, dwValue, 2&
End Function


