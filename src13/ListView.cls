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

'--------------Constants----------------
Private Const ICC_LISTVIEW_CLASSES = &H1          ' listview, header
Private Const IDD_LISTVIEW = 101
Private Const WC_LISTVIEW = "SysListView32"

'ListView Window Styles
Private Const LVS_ICON                       As Long = &H0
Private Const LVS_REPORT                     As Long = &H1
Private Const LVS_SMALLICON                  As Long = &H2
Private Const LVS_LIST                       As Long = &H3
Private Const LVS_TYPEMASK                   As Long = &H3
Private Const LVS_SINGLESEL                  As Long = &H4
Private Const LVS_SHOWSELALWAYS              As Long = &H8
Private Const LVS_SORTASCENDING              As Long = &H10
Private Const LVS_SORTDESCENDING             As Long = &H20
Private Const LVS_SHAREIMAGELISTS            As Long = &H40
Private Const LVS_NOLABELWRAP                As Long = &H80
Private Const LVS_AUTOARRANGE                As Long = &H100
Private Const LVS_EDITLABELS                 As Long = &H200
Private Const LVS_OWNERDATA                  As Long = &H1000
Private Const LVS_NOSCROLL                   As Long = &H2000

Private Const LVS_TYPESTYLEMASK              As Long = &HFC00

Private Const LVS_ALIGNTOP                   As Long = &H0
Private Const LVS_ALIGNLEFT                  As Long = &H800
Private Const LVS_ALIGNMASK                  As Long = &HC00

Private Const LVS_OWNERDRAWFIXED             As Long = &H400
Private Const LVS_NOCOLUMNHEADER             As Long = &H4000
Private Const LVS_NOSORTHEADER               As Long = &H8000

'Extended ListView Styles
Private Const LVS_EX_GRIDLINES               As Long = &H1
Private Const LVS_EX_SUBITEMIMAGES           As Long = &H2
Private Const LVS_EX_CHECKBOXES              As Long = &H4
Private Const LVS_EX_TRACKSELECT             As Long = &H8
Private Const LVS_EX_HEADERDRAGDROP          As Long = &H10
Private Const LVS_EX_FULLROWSELECT           As Long = &H20          'applies to report mode only
Private Const LVS_EX_ONECLICKACTIVATE        As Long = &H40
Private Const LVS_EX_TWOCLICKACTIVATE        As Long = &H80
Private Const LVS_EX_FLATSB                  As Long = &H100
Private Const LVS_EX_REGIONAL                As Long = &H200
Private Const LVS_EX_INFOTIP                 As Long = &H400         'listview does InfoTips for you
Private Const LVS_EX_UNDERLINEHOT            As Long = &H800
Private Const LVS_EX_UNDERLINECOLD           As Long = &H1000
Private Const LVS_EX_MULTIWORKAREAS          As Long = &H2000
Private Const LVS_EX_LABELTIP                As Long = &H4000        'listview unfolds partly hidden labels if it does not have infotip text
Private Const LVS_EX_BORDERSELECT            As Long = &H8000        'border selection style instead of highlight
Private Const LVS_EX_DOUBLEBUFFER            As Long = &H10000
Private Const LVS_EX_HIDELABELS              As Long = &H20000
Private Const LVS_EX_SINGLEROW               As Long = &H40000
Private Const LVS_EX_SNAPTOGRID              As Long = &H80000       'Icons automatically snap to grid.
Private Const LVS_EX_SIMPLESELECT            As Long = &H100000      'Also changes overlay rendering to top right for icon mode.
Private Const LVS_EX_JUSTIFYCOLUMNS          As Long = &H200000      'Icons are lined up in columns that use up the whole view area.
Private Const LVS_EX_TRANSPARENTBKGND        As Long = &H400000      'Background is painted by the parent via WM_PRINTCLIENT
Private Const LVS_EX_TRANSPARENTSHADOWTEXT   As Long = &H800000      'Enable shadow text on transparent backgrounds only (useful with bitmaps)
Private Const LVS_EX_AUTOAUTOARRANGE         As Long = &H1000000     'Icons automatically arrange if no icon positions have been set
Private Const LVS_EX_HEADERINALLVIEWS        As Long = &H2000000     'Display column header in all view modes
Private Const LVS_EX_AUTOCHECKSELECT         As Long = &H8000000
Private Const LVS_EX_AUTOSIZECOLUMNS         As Long = &H10000000
Private Const LVS_EX_COLUMNSNAPPOINTS        As Long = &H40000000
Private Const LVS_EX_COLUMNOVERFLOW          As Long = &H80000000


'these flags only apply to LVS_OWNERDATA listviews in report or list mode
Private Const LVSICF_NOINVALIDATEALL         As Long = &H1
Private Const LVSICF_NOSCROLL                As Long = &H2


'ListView Messages
Private Const LVM_FIRST                      As Long = &H1000
Private Const LVM_GETBKCOLOR                 As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR                 As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST               As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST               As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT               As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEM                    As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEM                    As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEM                 As Long = (LVM_FIRST + 7)
Private Const LVM_DELETEITEM                 As Long = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS             As Long = (LVM_FIRST + 9)
Private Const LVM_GETCALLBACKMASK            As Long = (LVM_FIRST + 10)
Private Const LVM_SETCALLBACKMASK            As Long = (LVM_FIRST + 11)
Private Const LVM_GETNEXTITEM                As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEM                   As Long = (LVM_FIRST + 13)
Private Const LVM_GETITEMRECT                As Long = (LVM_FIRST + 14)
Private Const LVM_SETITEMPOSITION            As Long = (LVM_FIRST + 15)
Private Const LVM_GETITEMPOSITION            As Long = (LVM_FIRST + 16)
Private Const LVM_GETSTRINGWIDTH             As Long = (LVM_FIRST + 17)
Private Const LVM_HITTEST                    As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE              As Long = (LVM_FIRST + 19)
Private Const LVM_SCROLL                     As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS                As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE                    As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABEL                  As Long = (LVM_FIRST + 23)
Private Const LVM_GETEDITCONTROL             As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMN                  As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMN                  As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMN               As Long = (LVM_FIRST + 27)
Private Const LVM_DELETECOLUMN               As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH             As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH             As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                  As Long = (LVM_FIRST + 31)
Private Const LVM_CREATEDRAGIMAGE            As Long = (LVM_FIRST + 33)
Private Const LVM_GETVIEWRECT                As Long = (LVM_FIRST + 34)
Private Const LVM_GETTEXTCOLOR               As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR               As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR             As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR             As Long = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX                As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE            As Long = (LVM_FIRST + 40)
Private Const LVM_GETORIGIN                  As Long = (LVM_FIRST + 41)
Private Const LVM_UPDATE                     As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE               As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE               As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT                As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXT                As Long = (LVM_FIRST + 46)
Private Const LVM_SETITEMCOUNT               As Long = (LVM_FIRST + 47)
Private Const LVM_SORTITEMS                  As Long = (LVM_FIRST + 48)
Private Const LVM_SETITEMPOSITION32          As Long = (LVM_FIRST + 49)
Private Const LVM_GETSELECTEDCOUNT           As Long = (LVM_FIRST + 50)
Private Const LVM_GETITEMSPACING             As Long = (LVM_FIRST + 51)
Private Const LVM_GETISEARCHSTRING           As Long = (LVM_FIRST + 52)
Private Const LVM_SETICONSPACING             As Long = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE   As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE   As Long = (LVM_FIRST + 55)
Private Const LVM_GETSUBITEMRECT             As Long = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST             As Long = (LVM_FIRST + 57)
Private Const LVM_SETCOLUMNORDERARRAY        As Long = (LVM_FIRST + 58)
Private Const LVM_GETCOLUMNORDERARRAY        As Long = (LVM_FIRST + 59)
Private Const LVM_SETHOTITEM                 As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM                 As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR               As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR               As Long = (LVM_FIRST + 63)
Private Const LVM_APPROXIMATEVIEWRECT        As Long = (LVM_FIRST + 64)
Private Const LVM_SETWORKAREAS               As Long = (LVM_FIRST + 65)
Private Const LVM_GETSELECTIONMARK           As Long = (LVM_FIRST + 66)
Private Const LVM_SETSELECTIONMARK           As Long = (LVM_FIRST + 67)
Private Const LVM_SETBKIMAGE                 As Long = (LVM_FIRST + 68)
Private Const LVM_GETBKIMAGE                 As Long = (LVM_FIRST + 69)
Private Const LVM_GETWORKAREAS               As Long = (LVM_FIRST + 70)
Private Const LVM_SETHOVERTIME               As Long = (LVM_FIRST + 71)
Private Const LVM_GETHOVERTIME               As Long = (LVM_FIRST + 72)
Private Const LVM_GETNUMBEROFWORKAREAS       As Long = (LVM_FIRST + 73)
Private Const LVM_SETTOOLTIPS                As Long = (LVM_FIRST + 74)
Private Const LVM_GETITEMW                   As Long = (LVM_FIRST + 75)
Private Const LVM_SETITEMW                   As Long = (LVM_FIRST + 76)      'Unicode
Private Const LVM_INSERTITEMW                As Long = (LVM_FIRST + 77)      'Unicode
Private Const LVM_GETTOOLTIPS                As Long = (LVM_FIRST + 78)
Private Const LVM_GETHOTLIGHTCOLOR           As Long = (LVM_FIRST + 79)      'UNDOCUMENTED
Private Const LVM_SETHOTLIGHTCOLOR           As Long = (LVM_FIRST + 80)      'UNDOCUMENTED
Private Const LVM_SORTITEMSEX                As Long = (LVM_FIRST + 81)
Private Const LVM_SETRANGEOBJECT             As Long = (LVM_FIRST + 82)      'UNDOCUMENTED
Private Const LVM_FINDITEMW                  As Long = (LVM_FIRST + 83)      'Unicode
Private Const LVM_RESETEMPTYTEXT             As Long = (LVM_FIRST + 84)      'UNDOCUMENTED
Private Const LVM_SETFROZENITEM              As Long = (LVM_FIRST + 85)      'UNDOCUMENTED
Private Const LVM_GETFROZENITEM              As Long = (LVM_FIRST + 86)      'UNDOCUMENTED
Private Const LVM_GETSTRINGWIDTHW            As Long = (LVM_FIRST + 87)
Private Const LVM_SETFROZENSLOT              As Long = (LVM_FIRST + 88)      'UNDOCUMENTED
Private Const LVM_GETFROZENSLOT              As Long = (LVM_FIRST + 89)      'UNDOCUMENTED
Private Const LVM_SETVIEWMARGIN              As Long = (LVM_FIRST + 90)      'UNDOCUMENTED
Private Const LVM_GETVIEWMARGIN              As Long = (LVM_FIRST + 91)      'UNDOCUMENTED
Private Const LVM_GETGROUPSTATE              As Long = (LVM_FIRST + 92)
Private Const LVM_GETFOCUSEDGROUP            As Long = (LVM_FIRST + 93)
Private Const LVM_EDITGROUPLABEL             As Long = (LVM_FIRST + 94)      'UNDOCUMENTED
Private Const LVM_GETCOLUMNW                 As Long = (LVM_FIRST + 95)      'Unicode
Private Const LVM_SETCOLUMNW                 As Long = (LVM_FIRST + 96)      'Unicode
Private Const LVM_INSERTCOLUMNW              As Long = (LVM_FIRST + 97)      'Unicode
Private Const LVM_GETGROUPRECT               As Long = (LVM_FIRST + 98)

Private Const LVM_GETITEMTEXTW               As Long = (LVM_FIRST + 115)     'Unicode
Private Const LVM_SETITEMTEXTW               As Long = (LVM_FIRST + 116)     'Unicode
Private Const LVM_GETISEARCHSTRINGW          As Long = (LVM_FIRST + 117)
Private Const LVM_EDITLABELW                 As Long = (LVM_FIRST + 118)

Private Const LVM_SETBKIMAGEW                As Long = (LVM_FIRST + 138)
Private Const LVM_GETBKIMAGEW                As Long = (LVM_FIRST + 139)
Private Const LVM_SETSELECTEDCOLUMN          As Long = (LVM_FIRST + 140)
Private Const LVM_SETTILEWIDTH               As Long = (LVM_FIRST + 141)
Private Const LVM_SETVIEW                    As Long = (LVM_FIRST + 142)
Private Const LVM_GETVIEW                    As Long = (LVM_FIRST + 143)

Private Const LVM_INSERTGROUP                As Long = (LVM_FIRST + 145)

Private Const LVM_SETGROUPINFO               As Long = (LVM_FIRST + 147)

Private Const LVM_GETGROUPINFO               As Long = (LVM_FIRST + 149)
Private Const LVM_REMOVEGROUP                As Long = (LVM_FIRST + 150)
Private Const LVM_MOVEGROUP                  As Long = (LVM_FIRST + 151)
Private Const LVM_GETGROUPCOUNT              As Long = (LVM_FIRST + 152)
Private Const LVM_GETGROUPINFOBYINDEX        As Long = (LVM_FIRST + 153)
Private Const LVM_MOVEITEMTOGROUP            As Long = (LVM_FIRST + 154)
Private Const LVM_SETGROUPMETRICS            As Long = (LVM_FIRST + 155)
Private Const LVM_GETGROUPMETRICS            As Long = (LVM_FIRST + 156)
Private Const LVM_ENABLEGROUPVIEW            As Long = (LVM_FIRST + 157)
Private Const LVM_SORTGROUPS                 As Long = (LVM_FIRST + 158)
Private Const LVM_INSERTGROUPSORTED          As Long = (LVM_FIRST + 159)
Private Const LVM_REMOVEALLGROUPS            As Long = (LVM_FIRST + 160)
Private Const LVM_HASGROUP                   As Long = (LVM_FIRST + 161)
Private Const LVM_SETTILEVIEWINFO            As Long = (LVM_FIRST + 162)
Private Const LVM_GETTILEVIEWINFO            As Long = (LVM_FIRST + 163)
Private Const LVM_SETTILEINFO                As Long = (LVM_FIRST + 164)
Private Const LVM_GETTILEINFO                As Long = (LVM_FIRST + 165)
Private Const LVM_SETINSERTMARK              As Long = (LVM_FIRST + 166)
Private Const LVM_GETINSERTMARK              As Long = (LVM_FIRST + 167)
Private Const LVM_INSERTMARKHITTEST          As Long = (LVM_FIRST + 168)
Private Const LVM_GETINSERTMARKRECT          As Long = (LVM_FIRST + 169)
Private Const LVM_SETINSERTMARKCOLOR         As Long = (LVM_FIRST + 170)
Private Const LVM_GETINSERTMARKCOLOR         As Long = (LVM_FIRST + 171)

Private Const LVM_SETINFOTIP                 As Long = (LVM_FIRST + 173)
Private Const LVM_GETSELECTEDCOLUMN          As Long = (LVM_FIRST + 174)
Private Const LVM_ISGROUPVIEWENABLED         As Long = (LVM_FIRST + 175)
Private Const LVM_GETOUTLINECOLOR            As Long = (LVM_FIRST + 176)
Private Const LVM_SETOUTLINECOLOR            As Long = (LVM_FIRST + 177)
Private Const LVM_SETKEYBOARDSELECTED        As Long = (LVM_FIRST + 178)     'UNDOCUMENTED
Private Const LVM_CANCELEDITLABEL            As Long = (LVM_FIRST + 179)
Private Const LVM_MAPINDEXTOID               As Long = (LVM_FIRST + 180)
Private Const LVM_MAPIDTOINDEX               As Long = (LVM_FIRST + 181)
Private Const LVM_ISITEMVISIBLE              As Long = (LVM_FIRST + 182)
Private Const LVM_EDITSUBITEM                As Long = (LVM_FIRST + 183)     'UNDOCUMENTED
Private Const LVM_ENSURESUBITEMVISIBLE       As Long = (LVM_FIRST + 184)     'UNDOCUMENTED
Private Const LVM_GETCLIENTRECT              As Long = (LVM_FIRST + 185)     'UNDOCUMENTED
Private Const LVM_GETFOCUSEDCOLUMN           As Long = (LVM_FIRST + 186)     'UNDOCUMENTED
Private Const LVM_SETOWNERDATACALLBACK       As Long = (LVM_FIRST + 187)     'UNDOCUMENTED
Private Const LVM_RECOMPUTEITEMS             As Long = (LVM_FIRST + 188)     'UNDOCUMENTED
Private Const LVM_QUERYINTERFACE             As Long = (LVM_FIRST + 189)     'UNDOCUMENTED: NOT OFFICIAL NAME
Private Const LVM_SETGROUPSUBSETCOUNT        As Long = (LVM_FIRST + 190)     'UNDOCUMENTED
Private Const LVM_GETGROUPSUBSETCOUNT        As Long = (LVM_FIRST + 191)     'UNDOCUMENTED
Private Const LVM_ORDERTOINDEX               As Long = (LVM_FIRST + 192)     'UNDOCUMENTED
Private Const LVM_GETACCVERSION              As Long = (LVM_FIRST + 193)     'UNDOCUMENTED
Private Const LVM_MAPACCIDTOACCINDEX         As Long = (LVM_FIRST + 194)     'UNDOCUMENTED
Private Const LVM_MAPACCINDEXTOACCID         As Long = (LVM_FIRST + 195)     'UNDOCUMENTED
Private Const LVM_GETOBJECTCOUNT             As Long = (LVM_FIRST + 196)     'UNDOCUMENTED
Private Const LVM_GETOBJECTRECT              As Long = (LVM_FIRST + 197)     'UNDOCUMENTED
Private Const LVM_ACCHITTEST                 As Long = (LVM_FIRST + 198)     'UNDOCUMENTED
Private Const LVM_GETFOCUSEDOBJECT           As Long = (LVM_FIRST + 199)     'UNDOCUMENTED
Private Const LVM_GETOBJECTROLE              As Long = (LVM_FIRST + 200)     'UNDOCUMENTED
Private Const LVM_GETOBJECTSTATE             As Long = (LVM_FIRST + 201)     'UNDOCUMENTED
Private Const LVM_ACCNAVIGATE                As Long = (LVM_FIRST + 202)     'UNDOCUMENTED
Private Const LVM_INVOKEDEFAULTACTION        As Long = (LVM_FIRST + 203)     'UNDOCUMENTED
Private Const LVM_GETEMPTYTEXT               As Long = (LVM_FIRST + 204)
Private Const LVM_GETFOOTERRECT              As Long = (LVM_FIRST + 205)
Private Const LVM_GETFOOTERINFO              As Long = (LVM_FIRST + 206)
Private Const LVM_GETFOOTERITEMRECT          As Long = (LVM_FIRST + 207)
Private Const LVM_GETFOOTERITEM              As Long = (LVM_FIRST + 208)
Private Const LVM_GETITEMINDEXRECT           As Long = (LVM_FIRST + 209)
Private Const LVM_SETITEMINDEXSTATE          As Long = (LVM_FIRST + 210)
Private Const LVM_GETNEXTITEMINDEX           As Long = (LVM_FIRST + 211)
Private Const LVM_SETPRESERVEALPHA           As Long = (LVM_FIRST + 212)     'UNDOCUMENTED

Private Const LVM_SETUNICODEFORMAT           As Long = CCM_SETUNICODEFORMAT
Private Const LVM_GETUNICODEFORMAT           As Long = CCM_GETUNICODEFORMAT


' ============================================
' Notifications

'ListView Notifications
Private Const LVN_FIRST                      As Long = -100&                     ' &HFFFFFF9C   ' (0U-100U)
Private Const LVN_LAST                       As Long = -199&                     ' &HFFFFFF39   ' (0U-199U)
                                                                                ' lParam points to:
Private Const LVN_ITEMCHANGING               As Long = (LVN_FIRST - 0)           ' NMLISTVIEW, ?, rtn T/F
Private Const LVN_ITEMCHANGED                As Long = (LVN_FIRST - 1)           ' NMLISTVIEW, ?
Private Const LVN_INSERTITEM                 As Long = (LVN_FIRST - 2)           ' NMLISTVIEW, iItem
Private Const LVN_DELETEITEM                 As Long = (LVN_FIRST - 3)           ' NMLISTVIEW, iItem
Private Const LVN_DELETEALLITEMS             As Long = (LVN_FIRST - 4)           ' NMLISTVIEW, iItem = -1, rtn T/F

Private Const LVN_COLUMNCLICK                As Long = (LVN_FIRST - 8)           ' NMLISTVIEW, iItem = -1, iSubItem = column
Private Const LVN_BEGINDRAG                  As Long = (LVN_FIRST - 9)           ' NMLISTVIEW, iItem
Private Const LVN_BEGINRDRAG                 As Long = (LVN_FIRST - 11)          ' NMLISTVIEW, iItem

Private Const LVN_ODCACHEHINT                As Long = (LVN_FIRST - 13)          ' NMLVCACHEHINT
Private Const LVN_ITEMACTIVATE               As Long = (LVN_FIRST - 14)          ' v4.70 = NMHDR, v4.71 = NMITEMACTIVATE
Private Const LVN_ODSTATECHANGED             As Long = (LVN_FIRST - 15)          ' NMLVODSTATECHANGE, rtn T/F
Private Const LVN_HOTTRACK                   As Long = (LVN_FIRST - 21)          ' NMLISTVIEW, see docs, rtn T/F
Private Const LVN_BEGINLABELEDITA            As Long = (LVN_FIRST - 5)           ' NMLVDISPINFO, iItem, rtn T/F
Private Const LVN_ENDLABELEDITA              As Long = (LVN_FIRST - 6)           ' NMLVDISPINFO, see docs
 
Private Const LVN_GETDISPINFOA               As Long = (LVN_FIRST - 50)          ' NMLVDISPINFO, see docs
Private Const LVN_SETDISPINFOA               As Long = (LVN_FIRST - 51)          ' NMLVDISPINFO, see docs
Private Const LVN_ODFINDITEMA                As Long = (LVN_FIRST - 52)          ' NMLVFINDITEM
 
Private Const LVN_KEYDOWN                    As Long = (LVN_FIRST - 55)          ' NMLVKEYDOWN
Private Const LVN_MARQUEEBEGIN               As Long = (LVN_FIRST - 56)          ' NMLISTVIEW, rtn T/F
Private Const LVN_GETINFOTIPA                As Long = (LVN_FIRST - 57)          ' NMLVGETINFOTIP
Private Const LVN_INCREMENTALSEARCHA         As Long = (LVN_FIRST - 62)
Private Const LVN_INCREMENTALSEARCHW         As Long = (LVN_FIRST - 63)
  
Private Const LVN_COLUMNDROPDOWN             As Long = (LVN_FIRST - 64)
Private Const LVN_COLUMNOVERFLOWCLICK        As Long = (LVN_FIRST - 66)
  
Private Const LVN_BEGINSCROLL                As Long = (LVN_FIRST - 80)
Private Const LVN_ENDSCROLL                  As Long = (LVN_FIRST - 81)
Private Const LVN_LINKCLICK                  As Long = (LVN_FIRST - 84)
Private Const LVN_GETEMPTYMARKUP             As Long = (LVN_FIRST - 87)
Private Const LVN_GROUPCHANGED               As Long = (LVN_FIRST - 88)          ' Undocumented
Private Const LVN_BEGINLABELEDITW            As Long = (LVN_FIRST - 75)
Private Const LVN_ENDLABELEDITW              As Long = (LVN_FIRST - 76)
Private Const LVN_GETDISPINFOW               As Long = (LVN_FIRST - 77)
Private Const LVN_SETDISPINFOW               As Long = (LVN_FIRST - 78)
Private Const LVN_ODFINDITEMW                As Long = (LVN_FIRST - 79)          ' NMLVFINDITEM
Private Const LVN_GETINFOTIPW                As Long = (LVN_FIRST - 58)          ' NMLVGETINFOTIP

#If Unicode Then
Private Const LVN_BEGINLABELEDIT             As Long = LVN_BEGINLABELEDITW
Private Const LVN_ENDLABELEDIT               As Long = LVN_ENDLABELEDITW
Private Const LVN_GETDISPINFO                As Long = LVN_GETDISPINFOW
Private Const LVN_SETDISPINFO                As Long = LVN_SETDISPINFOW
Private Const LVN_ODFINDITEM                 As Long = LVN_ODFINDITEMW           ' NMLVFINDITEM
Private Const LVN_GETINFOTIP                 As Long = LVN_GETINFOTIPW           ' NMLVGETINFOTIP
Private Const LVN_INCREMENTALSEARCH          As Long = LVN_INCREMENTALSEARCHW
#Else
Private Const LVN_BEGINLABELEDIT             As Long = LVN_BEGINLABELEDITA
Private Const LVN_ENDLABELEDIT               As Long = LVN_ENDLABELEDITA
Private Const LVN_GETDISPINFO                As Long = LVN_GETDISPINFOA
Private Const LVN_SETDISPINFO                As Long = LVN_SETDISPINFOA
Private Const LVN_ODFINDITEM                 As Long = LVN_ODFINDITEMA           ' NMLVFINDITEM
Private Const LVN_GETINFOTIP                 As Long = LVN_GETINFOTIPA           ' NMLVGETINFOTIP
Private Const LVN_INCREMENTALSEARCH          As Long = LVN_INCREMENTALSEARCHA
#End If   ' UNICODE


'Private Const LVN_ITEMCHANGED  As Long = -100 - 1

' LVITEM mask
Private Const LVIF_TEXT                      As Long = &H1
Private Const LVIF_IMAGE                     As Long = &H2
Private Const LVIF_PARAM                     As Long = &H4
Private Const LVIF_STATE                     As Long = &H8
Private Const LVIF_INDENT                    As Long = &H10
Private Const LVIF_NORECOMPUTE               As Long = &H800
Private Const LVIF_GROUPID                   As Long = &H100
Private Const LVIF_COLUMNS                   As Long = &H200
Private Const LVIF_DI_SETITEM                As Long = &H1000    'used only with the LVN_GETDISPINFO notification code.
Private Const LVIF_COLFMT                    As Long = &H10000   'The piColFmt member is valid in addition to puColumns

Private Const LVCF_FMT                       As Long = &H1
Private Const LVCF_WIDTH                     As Long = &H2
Private Const LVCF_TEXT                      As Long = &H4
Private Const LVCF_SUBITEM                   As Long = &H8
Private Const LVCF_IMAGE                     As Long = &H10
Private Const LVCF_ORDER                     As Long = &H20
Private Const LVCF_MINWIDTH                  As Long = &H40
Private Const LVCF_DEFAULTWIDTH              As Long = &H80
Private Const LVCF_IDEALWIDTH                As Long = &H100

'Private Const LVIF_TEXT  As Long = 1

Private Const LVSCW_AUTOSIZE   As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER   As Long = -2


Private Type LVCOLUMNA
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

Private Type LVCOLUMNW
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

Private Type LVITEMA
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

Private Type LVITEMW
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

' LVM_HITTEST lParam
Private Type LVHITTESTINFO   ' was LV_HITTESTINFO
    pt As POINTAPI
    Flags As LVHT_FLAGS
    iItem As Long
    iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
    iGroup As Long
End Type

Private Type NMLVDISPINFOA   ' was LV_DISPINFO
    hdr As NMHDR
    Item As LVITEMA
End Type

Private Type NMLVDISPINFOW   ' was LV_DISPINFO
    hdr As NMHDR
    Item As LVITEMW
End Type

Private Type NMLISTVIEW   ' was NM_LISTVIEW
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

Private Type NMITEMACTIVATE
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

'--------------Enums----------------
Private Enum ListPictureAlignmentConstants
    lvwTopLeft = 0
    lvwTopRight = 1
    lvwBottomLeft = 2
    lvwBottomRight = 3
    lvwCenter = 4
    lvwTile = 5
End Enum

Private Enum ListSortOrderConstants
    lvwAscending = 1
    lvwDescending = 2
End Enum

Private Enum ListTextBackgroundConstants
    lvwTransparent = 0
    lvwOpaque = 1
End Enum

Private Enum ListViewConstants
    lvwIcon = 0
    lvwSmallIcon = 1
    lvwList = 2
    lvwReport = 3
End Enum

Private Enum LISTVIEWITEMRECT
    LVIR_BOUNDS = 0
    LVIR_ICON = 1
    LVIR_LABEL = 2
    LVIR_SELECTBOUNDS = 3
End Enum

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

Private Type Fields
    Frame As IControl
    hFont As LongPtr
    Font As StdFont
    columnCount As Long
    ListItemsCount As Long
    hParent As LongPtr
    hChild As LongPtr
    pfn As LongPtr
    ColumnHeaders As ColumnHeaders
    ListItems As ListItems
    ItemCountPerPage As Long
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
        .columnCount = 0
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
        MeasureProcTime 0
        Set .ColumnHeaders = Nothing
        Debug.Print "ColumnHeaders破棄:" & MeasureProcTime(1)
        MeasureProcTime 0
        Set .ListItems = Nothing
        Debug.Print "ListItems破棄:" & MeasureProcTime(1)
    End With
'    Set VLItems = Nothing
End Sub

Public Function Init(ByVal ctl As IOptionFrame) As Boolean
    Init = False
    Dim mWS As Long
    
    Dim iccex As InitCommonControlsExType
    Dim wd As Long, ht As Long
    
    Dim flbkColor As Long
    Dim buf As Long
    
    If This.hParent Then Exit Function
    Set This.Frame = ctl
       
    This.hParent = This.Frame.[_GethWnd]
    
'    Call WindowFromAccessibleObject(fr, This.hParent)
    
    If This.hParent = 0 Then
        err.Raise 91
    End If
    
    Set acc = This.Frame
    acc.accLocation 0, 0, wd, ht
    
    iccex.dwSize = Len(iccex)
'    Debug.Print iccex.dwSize
    iccex.dwICC = ICC_LISTVIEW_CLASSES
    InitCommonControlsEx iccex
    
    With This
'        .pfn = GetPtr(AddressOf ListViewModule.Redirect)
        .pfn = GetPtr(AddressOf Redirect)
'        Debug.Print "hParent:" & .hParent & " / pfn:" & .pfn
        If SetWindowSubclass(.hParent, .pfn, .hParent, ObjPtr(Me)) = 0 Then Debug.Print "FrameSubclassError": Exit Function
               
'        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hParent), "Unicode", "ANCI")
        
        'コモンコントロールのウィンドウスタイル
        mWS = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
        mWS = mWS Or LVS_REPORT Or LVS_SHOWSELALWAYS Or LVS_OWNERDATA
        
        On Error GoTo ErrorHandler
        
'        .hChild = CreateWindowEx(WS_EX_CLIENTEDGE, _
                            WC_LISTVIEW, 0, mWS, _
                            0, 0, wd, ht, .hParent, IDD_LISTVIEW, 0, 0)
'        Dim WindowClass As String
''        WindowClass = StrConv("SysListView32", vbUnicode)
'        WindowClass = "SysListView32"
'        Debug.Print StrPtr(WindowClass)
        .hChild = CreateWindowExW(WS_EX_CLIENTEDGE, _
                            StrPtr(WC_LISTVIEW), 0, mWS, _
                            0, 0, wd, ht, .hParent, IDD_LISTVIEW, 0, 0)
        If .hChild = 0 Then Debug.Print "CreateWindow Error": Exit Function
'        Debug.Print "Frame:" & IIf(IsWindowUnicode(.hChild), "Unicode", "ANCI")
'        Debug.Print "ｈChild" & .hChild
        If SetWindowSubclass(.hChild, .pfn, .hChild, ObjPtr(Me)) = 0 Then Debug.Print "ListViewSubclassError": Exit Function
                                    
        '拡張スタイル
        mWS = LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER 'Or LVS_EX_CHECKBOXES
        SendMessageW .hChild, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal mWS

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
       .hFont = CreateFontW(16, _
                         FW_NORMAL, _
                         0, _
                         0, _
                         0, _
                         False, _
                         False, _
                         False, _
                         SHIFTJIS_CHARSET, _
                         OUT_DEFAULT_PRECIS, _
                         CLIP_DEFAULT_PRECIS, _
                         DEFAULT_QUALITY, _
                         DEFAULT_PITCH, _
                         StrPtr("Yu Gothic UI"))
        
'        SendMessageW .hChild, WM_SETFONT, .hFont, ByVal 1&
        
'        SendMessage .hChild, WM_SETFONT, This.hFont, ByVal 1&
'        SendMessage hChild, WM_SETFONT, This.hFont, ByVal 1&

'        SendMessageW .hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
''    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]
    
        .ItemCountPerPage = GetCountPerPage
        Debug.Print .ItemCountPerPage
    End With
    Init = True
    Exit Function
    
ErrorHandler:
    Debug.Print GetAPIError(err.LastDllError)
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
    Dim ColumnText As String
    Dim TextWidth As Long
    
    This.columnCount = This.ColumnHeaders.Count
    
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 0, 0))

'    i = 1
'    For i = 0 To This.ColumnHeaders.Count - 1
    For i = 0 To This.columnCount - 1
'    For i = This.columnCount - 1 To 0 Step -1
        ColumnText = This.ColumnHeaders(i).Text
'        TextWidth = GetStringWidthPixel(ColumnText)
        This.ColumnHeaders(i).Width = GetStringWidth(ColumnText) + 12
        Call InsertColumn(ColumnText, i, This.ColumnHeaders(i).Width)
'        Call InsertColumn(ColumnText, i, 10)
'        Debug.Print "SetHeader:" & i & ":" & This.ColumnHeaders(i)
'        Call SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER)
'        Call SetColumnWidth(i, LVSCW_AUTOSIZE)
    Next i
'    Dim i As Long
'    This.ColumnCount = UBound(Items)
'    For i = 1 To This.ColumnHeaders.Count
'        Call InsertColumn(Items(i), i, 0)
'    Next i

'    For i = This.columnCount - 1 To 0 Step -1
'        Call SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER)
'    Next i

    Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 1, 0))

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

'Public Function SetReportRecord(Records As Dictionary) As Long
'    Dim i As Long
'    Call InsertItem(Records.Item(i), This.ListItemsCount)
'    If This.ColumnCount < Records.Count Then
'        For i = This.ColumnCount + 1 To Records.Count
'            Call InsertColumn("", i, 0)
'        Next i
'        This.ColumnCount = Records.Count
'    End If
'    For i = 1 To This.ColumnCount
'        Call SetItem(Records.Item(i), This.ListItemsCount, i)
'    Next i
'    This.ListItemsCount = This.ListItemsCount + 1
'End Function

Public Function InsertColumn(Title As String, ByVal iCol As Long, ByVal Width As Long) As Long
    Dim lvc As LVCOLUMNW
    With lvc
        .mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_SUBITEM
        .cx = Width
        .pszText = StrPtr(Title)
'        Debug.Print Title
    End With
    With This
        If .hChild = 0 Then Exit Function
        InsertColumn = CLng(SendMessageW(.hChild, LVM_INSERTCOLUMNW, ByVal iCol, lvc))
    End With
End Function

Public Function InsertItem(sItem As String, ByVal iItem As Long) As Long
    Dim TextWidth As Long
    Dim lvi As LVITEMW
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .pszText = StrPtr(sItem)
    End With
    With This
        If .hChild = 0 Then Exit Function
        InsertItem = CLng(SendMessageW(.hChild, LVM_INSERTITEMW, 0, lvi))
        
        TextWidth = GetStringWidth(sItem) + 12
        Debug.Print TextWidth & "/" & This.ColumnHeaders(0).Width
        If TextWidth > This.ColumnHeaders(0).Width Then
            Debug.Print "InsertItem WidthSet"
            Call SetColumnWidth(iItem, TextWidth)
        End If
        
    End With
End Function

Public Function SetItem(sItem As String, ByVal iItem As Long, ByVal iSubItem As Long) As Long
    Dim TextWidth As Long
    Dim lvi As LVITEMW
    With lvi
        .mask = LVIF_TEXT
        .iItem = iItem
        .iSubItem = iSubItem
        .pszText = StrPtr(sItem)
    End With
    With This
        If .hChild = 0 Then Exit Function
        SetItem = CLng(SendMessageW(.hChild, LVM_SETITEMW, 0, lvi))
    
        TextWidth = GetStringWidth(sItem) + 12
        Debug.Print TextWidth & "/" & This.ColumnHeaders(iSubItem).Width
        If TextWidth > This.ColumnHeaders(iSubItem).Width Then
            Debug.Print "SetItem WidthSet"
            Call SetColumnWidth(iSubItem, TextWidth)
        End If
    
    End With
End Function

Public Function DeleteAllItems() As Long
    DeleteAllItems = CLng(SendMessageW(This.hChild, LVM_DELETEALLITEMS, 0, 0))
End Function

Public Function GetItemCount() As Long
    GetItemCount = CLng(SendMessageW(This.hChild, LVM_GETITEMCOUNT, 0, 0))
End Function

Public Sub SetItemCount(cItems As Long)
'    Debug.Print cItems
    SendMessageW This.hChild, LVM_SETITEMCOUNT, ByVal cItems, 0
End Sub

Public Function LabelText(ByVal iItem As Long, ByVal iSubItem As Long) As String
'    Debug.Print "Item:" & iItem & " / SubItem:" & iSubItem
    Dim lv As LVITEMW
'    Dim buf As String * 255
    Dim buf(255) As Integer
'    buf = String(255, vbNullChar)
'    Debug.Print "Len(buf):" & Len(buf)
    
    With lv
        .cchTextMax = 256
        .iSubItem = iSubItem
        .pszText = VarPtr(buf(0))
    End With
    SendMessageW This.hChild, LVM_GETITEMTEXTW, iItem, lv
    
    MoveMemory LabelText, lv.pszText, 2 * 256
    
'    LabelText = lv.pszText
'    Debug.Print "Len(LabelText):" & Len(LabelText)
'    Debug.Print "LabelText:" & LabelText
End Function

Public Function GetCheckState(ByVal iItem As Long) As Boolean
' チェック状態を取得（ii = インデックス）...
    ' LVM_GETITEMSTATE...
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_GETITEMSTATE, iItem, &H1000&))
    If Ret <> 0 Then
        GetCheckState = True
    Else
        GetCheckState = False
    End If
'    ii = CLng(SendMessage(This.hChild, LVM_GETITEMSTATE, ii, &H1000&))
'    GetCheckState = ii  '= 0
End Function

Public Sub SetCheckState(ByVal iItem As Long, ByVal f As Boolean)
' チェック状態をセット（ii = インデックス)
' iiが「-1」の時は全ての項目が対象
    Dim buf As LVITEMW
    With buf
        .stateMask = &HF000&    'LVIS_STATEIMAGEMASK
        .State = &H1000&
        If f Then .State = &H2000&
    End With
    ' LVM_SETITEMSTATE
    SendMessageW This.hChild, LVM_SETITEMSTATE, iItem, VarPtr(buf)
End Sub

Public Function SetTextColor(ClrText As Long) As Boolean
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_SETTEXTCOLOR, 0, ByVal ClrText))
    If Ret <> 0 Then
        SetTextColor = True
    Else
        SetTextColor = False
    End If
End Function

Public Function SetTextBkColor(ClrTextBk As Long) As Boolean
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_SETTEXTBKCOLOR, 0, ClrTextBk))
    If Ret <> 0 Then
        SetTextBkColor = True
    Else
        SetTextBkColor = False
    End If
End Function

Public Function RedrawItems(iFirst As Long, iLast As Long) As Boolean
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast))
    If Ret <> 0 Then
        RedrawItems = True
    Else
        RedrawItems = False
    End If
End Function

Public Function Update(i As Long) As Boolean
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_UPDATE, ByVal i, 0))
    If Ret <> 0 Then
        Update = True
    Else
        Update = False
    End If
End Function

Public Function GetCountPerPage() As Long
    GetCountPerPage = CLng(SendMessage(This.hChild, LVM_GETCOUNTPERPAGE, 0, 0))
End Function

Public Function GetStringWidth(psz As String) As Long
    GetStringWidth = CLng(SendMessageW(This.hChild, LVM_GETSTRINGWIDTHW, 0, ByVal StrPtr(psz)))
End Function

Public Function SetColumnWidth(iCol As Long, cx As Long) As Boolean
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(cx, 0)))
    If Ret <> 0 Then
        SetColumnWidth = True
    Else
        SetColumnWidth = False
    End If
End Function

Public Function SetAllColumnWidthAuto() As Boolean
    SetAllColumnWidthAuto = False
    Dim Ret As Long
    Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 0, 0))
    
    Dim i As Long
    For i = This.columnCount - 1 To 0 Step -1
'        SetAllColumnWidthAuto = SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER)
'        SetAllColumnWidthAuto = CLng(SendMessageW(This.hChild, LVM_SETCOLUMNWIDTH, ByVal i, ByVal MAKELPARAM(LVSCW_AUTOSIZE_USEHEADER, 0)))
'        SetAllColumnWidthAuto = CLng(SendMessageW(This.hChild, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER))
        SetAllColumnWidthAuto = SetColumnWidth(i, 20)
        If SetAllColumnWidthAuto = False Then Exit For
    Next i
    
    Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 1, 0))
End Function

' 文字列の幅(ピクセル)を取得する関数
Private Function GetStringWidthPixel(ByVal Text As String) As Long
    Dim hDC As LongPtr
    Dim sz As SIZE
    
    hDC = GetDC(This.hChild)
    
    ' 幅を取得
    Call GetTextExtentPoint32(hDC, Text, Len(Text), sz)
    
    ' 後処理
    Call ReleaseDC(This.hChild, hDC)
    
    GetStringWidthPixel = sz.cx
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
            Dim Ret As Long
            Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 0, 0))
            Dim LVDIW As NMLVDISPINFOW
            MoveMemory ByVal VarPtr(LVDIW), ByVal lParam, LenB(LVDIW)
            With LVDIW.Item
'            Debug.Print "GetDispInfoW"
                If (.mask And LVIF_TEXT) Then
'                    Debug.Print "LVIF_TEXT:Item:" & .iItem & " SubItem:" & .iSubItem
                    Select Case .iSubItem
                        Case 0
                            .pszText = This.ListItems(.iItem).TextPtr
                            
'                            Dim FieldText As String
'                            FieldText = This.ListItems(.iItem).Text
'                            Call SetColumnWidth(.iSubItem, GetStringWidth(FieldText) + 5)
                        Case Else
                            .pszText = This.ListItems(.iItem).ListSubItems(.iSubItem).TextPtr
'                            Call SetColumnWidth(.iSubItem, LVSCW_AUTOSIZE_USEHEADER)
                    End Select
                End If
                
                If (.mask And LVIF_IMAGE) Then
'                    Debug.Print "LVIF_IMAGE"
                    Select Case .iSubItem
                        Case 0
'                            .iImage = VLItems(.iItem).iImage
'                            .iImage = This.ListItems(.iItem).Icon
                    End Select
                End If
                MoveMemory ByVal lParam, ByVal VarPtr(LVDIW), LenB(LVDIW)
'                Call SetColumnWidth(.iSubItem, LVSCW_AUTOSIZE_USEHEADER)
                Ret = CLng(SendMessageW(This.hChild, WM_SETREDRAW, 1, 0))
            End With
           
        Case LVN_GETDISPINFOA
'            Debug.Print "GetDispInfoA"
            Dim LVDI As NMLVDISPINFOA
            MoveMemory ByVal VarPtr(LVDI), ByVal lParam, LenB(LVDI)
            With LVDI.Item
                
                If (.mask And LVIF_TEXT) Then
'                    Debug.Print "LVIF_TEXT"
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
'                    Debug.Print "LVIF_IMAGE"
                    Select Case .iSubItem
                        Case 0
                            .iImage = VLItems(.iItem).iImage
                    End Select
                End If
                MoveMemory ByVal lParam, ByVal VarPtr(LVDI), LenB(LVDI)
            End With
        Case Else
        
    End Select
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
    MeasureProcTime 0
    
    With CSV
        .FilePath = FileName
        .CharSet = CharSet
        .IncludeHeader = HasHeader
    End With
    
    'readCsvでCSVを読み込み
    Dim CSVText As String
    CSVText = readCsv
    Debug.Print "ReadCSV:" & MeasureProcTime(1)
 
     
    MeasureProcTime 0
'    ReDim VLItems(0 To 0)
    'CsvToJaggedで行・フィールドに分割してジャグ配列に
    Call CsvToJagged(CSVText)
    Debug.Print "SetListItems:" & MeasureProcTime(1)
    
    MeasureProcTime 0
    SetItemCount This.ListItems.Count
'    This.columnCount = This.ColumnHeaders.Count
    Debug.Print "SetItemCount:" & MeasureProcTime(1)
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
    
    MeasureProcTime 0
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
                            Debug.Print "SetHeader:" & MeasureProcTime(1)
                            MeasureProcTime 0
                        Else
                            Call SetFieldText(i, j, FieldText)
'                            Call SetRecord(ItemText, FieldArray, i, j, FieldText)
                            i = i + 1 '列位置
'                            Dim Item As VListItem
'                            VLItems.Add i, Item
'                            ReDim Preserve VLItems(0 To i)
'                            Debug.Print "SetRecord:" & MeasureProcTime(1)
'                            MeasureProcTime 0
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
                        If j = 0 Then
                            This.ColumnHeaders.Add j, , "L/N", 0
                        Else
                            This.ColumnHeaders.Add j + 1, , FieldText, 0
                        End If
'                    ElseIf j = 0 Then
'                        Call AddRecord(i, FieldText)
'                        ItemText = FieldText
                    Else
                        If j = 0 Then
                            Call SetFieldText(i, j, CStr(i))
                        Else
                            Call SetFieldText(i, j, FieldText)
                        End If
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

'1フィールドごとにセルに出力
Private Sub SetFieldText(ByRef i As Long, ByRef j As Long, ByRef FieldText As String)
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
    
    If j = 0 Then
'        Debug.Print "AddListItems:" & i & ":" & MeasureProcTime(1)
        This.ListItems.Add Index:=i, Key:="", Text:=FieldText
        This.ListItems(i).ListSubItems.Add Index:=j, Key:="", Width:=0, Text:=""
'        Debug.Print "AddListItems:" & i & ":" & MeasureProcTime(1)
'        MeasureProcTime 0
'        Debug.Print GetStringWidth(FieldText)
    Else
        This.ListItems(i).ListSubItems.Add Index:=j, Key:="", Width:=0, Text:=FieldText
'        Debug.Print GetStringWidth(FieldText)
    End If
    
    If i < This.ItemCountPerPage + 1 And j < This.columnCount Then
        Dim TextWidth As Long
        TextWidth = GetStringWidth(FieldText) + 12
'        Debug.Print TextWidth & "/" & This.ColumnHeaders(j).Width
        If TextWidth > This.ColumnHeaders(j).Width Then
'            Debug.Print "SetFieldText WidthSet:i:" & i & "/j:" & j
            Call SetColumnWidth(j, TextWidth)
        End If
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
'    Debug.Print "CharSet:" & CSV.CharSet
    Select Case UCase(CSV.CharSet)
        Case "UTF-8", "UTF-8N"
            'ADOを使って読込、その後の処理を統一するため全レコードをCrLfで結合
            'Set inTS = objFSO.OpenTextFile(strFile, ForAppending)
            Set inTS = objFSO.OpenTextFile(CSV.FilePath, 8)
            i = inTS.Line - 1
            inTS.Close
'            Debug.Print i & "行"
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



