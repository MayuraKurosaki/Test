Attribute VB_Name = "HeaderExModule"
Option Explicit

'--------------Constants----------------
' Header Window Class Names
Public Const HEADER32_CLASS             As String = "SysHeader32"
Public Const HEADER_CLASS               As String = "SysHeader"

' Common Control Messages
Public Const HDM_FIRST                  As Long = &H1200
Public Const HDM_GETITEMCOUNT           As Long = (HDM_FIRST + 0)
Public Const HDM_INSERTITEMA            As Long = (HDM_FIRST + 1)
Public Const HDM_DELETEITEM             As Long = (HDM_FIRST + 2)
Public Const HDM_GETITEMA               As Long = (HDM_FIRST + 3)
Public Const HDM_SETITEMA               As Long = (HDM_FIRST + 4)
Public Const HDM_LAYOUT                 As Long = (HDM_FIRST + 5)
Public Const HDM_HITTEST                As Long = (HDM_FIRST + 6)
Public Const HDM_GETITEMRECT            As Long = (HDM_FIRST + 7)
Public Const HDM_SETIMAGELIST           As Long = (HDM_FIRST + 8)
Public Const HDM_GETIMAGELIST           As Long = (HDM_FIRST + 9)
Public Const HDM_INSERTITEMW            As Long = (HDM_FIRST + 10)
Public Const HDM_GETITEMW               As Long = (HDM_FIRST + 11)
Public Const HDM_SETITEMW               As Long = (HDM_FIRST + 12)

Public Const HDM_ORDERTOINDEX           As Long = (HDM_FIRST + 15)
Public Const HDM_CREATEDRAGIMAGE        As Long = (HDM_FIRST + 16)      '// wparam = which item (by index)
Public Const HDM_GETORDERARRAY          As Long = (HDM_FIRST + 17)
Public Const HDM_SETORDERARRAY          As Long = (HDM_FIRST + 18)
Public Const HDM_SETHOTDIVIDER          As Long = (HDM_FIRST + 19)
Public Const HDM_SETBITMAPMARGIN        As Long = (HDM_FIRST + 20)
Public Const HDM_GETBITMAPMARGIN        As Long = (HDM_FIRST + 21)
Public Const HDM_SETFILTERCHANGETIMEOUT As Long = (HDM_FIRST + 22)
Public Const HDM_EDITFILTER             As Long = (HDM_FIRST + 23)
Public Const HDM_CLEARFILTER            As Long = (HDM_FIRST + 24)
Public Const HDM_GETITEMDROPDOWNRECT    As Long = (HDM_FIRST + 25)      ' // rect of item's drop down button
Public Const HDM_GETOVERFLOWRECT        As Long = (HDM_FIRST + 26)      '// rect of overflow button
Public Const HDM_GETFOCUSEDITEM         As Long = (HDM_FIRST + 27)
Public Const HDM_SETFOCUSEDITEM         As Long = (HDM_FIRST + 28)
Public Const HDM_TRANSLATEACCELERATOR   As Long = &H0461                ' CCM_TRANSLATEACCELERATOR

Public Const HDM_GETITEM                As Long = HDM_GETITEMA
Public Const HDM_SETITEM                As Long = HDM_SETITEMA
Public Const HDM_INSERTITEM             As Long = HDM_INSERTITEMA
Public Const HDM_SETUNICODEFORMAT       As Long = CCM_SETUNICODEFORMAT
Public Const HDM_GETUNICODEFORMAT       As Long = CCM_GETUNICODEFORMAT

' Header Notification Codes
Public Const HDN_FIRST                  As Long = -300
Public Const HDN_ITEMCLICK              As Long = (HDN_FIRST - 2)
Public Const HDN_DIVIDERDBLCLICK        As Long = (HDN_FIRST - 5)
Public Const HDN_BEGINTRACK             As Long = (HDN_FIRST - 6)
Public Const HDN_ENDTRACK               As Long = (HDN_FIRST - 7)
Public Const HDN_TRACK                  As Long = (HDN_FIRST - 8)
Public Const HDN_GETDISPINFO            As Long = (HDN_FIRST - 9)
Public Const HDN_ITEMCHANGING           As Long = (HDN_FIRST - 0)
Public Const HDN_ITEMDBLCLICK           As Long = (HDN_FIRST - 3)
Public Const HDN_ITEMCHANGINGA          As Long = (HDN_FIRST - 0)
Public Const HDN_ITEMCHANGINGW          As Long = (HDN_FIRST - 20)
Public Const HDN_ITEMCHANGEDA           As Long = (HDN_FIRST - 1)
Public Const HDN_ITEMCHANGEDW           As Long = (HDN_FIRST - 21)
Public Const HDN_ITEMCLICKA             As Long = (HDN_FIRST - 2)
Public Const HDN_ITEMCLICKW             As Long = (HDN_FIRST - 22)
Public Const HDN_ITEMDBLCLICKA          As Long = (HDN_FIRST - 3)
Public Const HDN_ITEMDBLCLICKW          As Long = (HDN_FIRST - 23)
Public Const HDN_DIVIDERDBLCLICKA       As Long = (HDN_FIRST - 5)
Public Const HDN_DIVIDERDBLCLICKW       As Long = (HDN_FIRST - 25)
Public Const HDN_BEGINTRACKA            As Long = (HDN_FIRST - 6)
Public Const HDN_BEGINTRACKW            As Long = (HDN_FIRST - 26)
Public Const HDN_ENDTRACKA              As Long = (HDN_FIRST - 7)
Public Const HDN_ENDTRACKW              As Long = (HDN_FIRST - 27)
Public Const HDN_TRACKA                 As Long = (HDN_FIRST - 8)
Public Const HDN_TRACKW                 As Long = (HDN_FIRST - 28)
Public Const HDN_GETDISPINFOA           As Long = (HDN_FIRST - 9)
Public Const HDN_GETDISPINFOW           As Long = (HDN_FIRST - 29)
Public Const HDN_BEGINDRAG              As Long = (HDN_FIRST - 10)
Public Const HDN_ENDDRAG                As Long = (HDN_FIRST - 11)
Public Const HDN_FILTERCHANGE           As Long = (HDN_FIRST - 12)
Public Const HDN_FILTERBTNCLICK         As Long = (HDN_FIRST - 13)
Public Const HDN_BEGINFILTEREDIT        As Long = (HDN_FIRST - 14)
Public Const HDN_ENDFILTEREDIT          As Long = (HDN_FIRST - 15)
Public Const HDN_ITEMSTATEICONCLICK     As Long = (HDN_FIRST - 16)
Public Const HDN_ITEMKEYDOWN            As Long = (HDN_FIRST - 17)
Public Const HDN_DROPDOWN               As Long = (HDN_FIRST - 18)
Public Const HDN_OVERFLOWCLICK          As Long = (HDN_FIRST - 19)

' Header Control Styles
Public Const HDS_HORZ                   As Long = &H0000
Public Const HDS_BUTTONS                As Long = &H0002
Public Const HDS_HOTTRACK               As Long = &H0004
Public Const HDS_HIDDEN                 As Long = &H0008
Public Const HDS_DRAGDROP               As Long = &H0040
Public Const HDS_FULLDRAG               As Long = &H0080
Public Const HDS_FILTERBAR              As Long = &H0100
Public Const HDS_FLAT                   As Long = &H0200
Public Const HDS_CHECKBOXES             As Long = &H0400
Public Const HDS_NOSIZING               As Long = &H0800
Public Const HDS_OVERFLOW               As Long = &H1000

' Header Item Format Constants
Public Const HDF_LEFT                   As Long = &H0000        '// Same as LVCFMT_LEFT
Public Const HDF_RIGHT                  As Long = &H0001        '// Same as LVCFMT_RIGHT
Public Const HDF_CENTER                 As Long = &H0002        '// Same as LVCFMT_CENTER
Public Const HDF_JUSTIFYMASK            As Long = &H0003        '// Same as LVCFMT_JUSTIFYMASK
Public Const HDF_RTLREADING             As Long = &H0004        '// Same as LVCFMT_LEFT
Public Const HDF_CHECKBOX               As Long = &H0040
Public Const HDF_CHECKED                As Long = &H0080
Public Const HDF_FIXEDWIDTH             As Long = &H0100        '// Can't resize the column; same as LVCFMT_FIXED_WIDTH
Public Const HDF_SORTDOWN               As Long = &H0200
Public Const HDF_SORTUP                 As Long = &H0400
Public Const HDF_IMAGE                  As Long = &H0800        '// Same as LVCFMT_IMAGE
Public Const HDF_BITMAP_ON_RIGHT        As Long = &H1000        '// Same as LVCFMT_BITMAP_ON_RIGHT
Public Const HDF_BITMAP                 As Long = &H2000
Public Const HDF_STRING                 As Long = &H4000
Public Const HDF_OWNERDRAW              As Long = &H8000        '// Same as LVCFMT_COL_HAS_IMAGES
Public Const HDF_SPLITBUTTON            As Long = &H1000000     '// Column is a split button; same as LVCFMT_SPLITBUTTON

' Header Item Filter Type Constants
Public Const HDFT_ISSTRING              As Long = &H0000        '// HD_ITEM.pvFilter points to a HD_TEXTFILTER
Public Const HDFT_ISNUMBER              As Long = &H0001        '// HD_ITEM.pvFilter points to a INT
Public Const HDFT_ISDATE                As Long = &H0002        '// HD_ITEM.pvFilter points to a DWORD (dos date)
Public Const HDFT_HASNOVALUE            As Long = &H8000        '// clear the filter, by setting this bit

' Header Item State Constants
Public Const HDIS_FOCUSED               As Long = &H0001
Public Const HDIS_SELECTED              As Long = &H0002
Public Const HDIS_HOTTRACKED            As Long = &H0004

' Header Item Mask Constants
Public Const HDI_WIDTH                  As Long = &H0001
Public Const HDI_HEIGHT                 As Long = HDI_WIDTH
Public Const HDI_TEXT                   As Long = &H0002
Public Const HDI_FORMAT                 As Long = &H0004
Public Const HDI_LPARAM                 As Long = &H0008
Public Const HDI_BITMAP                 As Long = &H0010
Public Const HDI_IMAGE                  As Long = &H0020
Public Const HDI_DI_SETITEM             As Long = &H0040
Public Const HDI_ORDER                  As Long = &H0080
Public Const HDI_FILTER                 As Long = &H0100
Public Const HDI_STATE                  As Long = &H0200

' Header Hit Test Constants
Public Const HHT_NOWHERE                As Long = &H0001
Public Const HHT_ONHEADER               As Long = &H0002
Public Const HHT_ONDIVIDER              As Long = &H0004
Public Const HHT_ONDIVOPEN              As Long = &H0008
Public Const HHT_ONFILTER               As Long = &H0010
Public Const HHT_ONFILTERBUTTON         As Long = &H0020
Public Const HHT_ABOVE                  As Long = &H0100
Public Const HHT_BELOW                  As Long = &H0200
Public Const HHT_TORIGHT                As Long = &H0400
Public Const HHT_TOLEFT                 As Long = &H0800
Public Const HHT_ONITEMSTATEICON        As Long = &H1000
Public Const HHT_ONDROPDOWN             As Long = &H2000
Public Const HHT_ONOVERFLOW             As Long = &H4000


'--------------Enums----------------
Public Type HD_TEXTFILTERA
    pszText As LongPtr                      ' [in] pointer to the buffer containing the filter (ANSI)
    cchTextMax As Long                      ' [in] max size of buffer/edit control buffer
End Type

Public Type HD_TEXTFILTERW
    pszText As LongPtr                      ' [in] pointer to the buffer containing the filter (UNICODE)
    cchTextMax As Long                      ' [in] max size of buffer/edit control buffer
End Type

Public Type HD_ITEMA
    mask As Long
    cxy As Long
    pszText As LongPtr
    hbm As LongPtr
    cchTextMax As Long
    fmt As Long
    lParam As LongPtr
    iImage As Long
    iOrder As Long
    type As Long
    pvFilter As LongPtr
    state As Long
End Type

Public Type HD_ITEMW
    mask As Long
    cxy As Long
    pszText As LongPtr
    hbm As LongPtr
    cchTextMax As Long
    fmt As Long
    lParam As LongPtr
    iImage As Long
    iOrder As Long
    type As Long
    pvFilter As LongPtr
    state As Long
End Type

Public Type HD_LAYOUT
    prc As LongPtr
    pwpos As LongPtr
End Type

Public Type HD_HITTESTINFO
    pt As POINT
    flags As Long
    iItem As Long
End Type

Public Type NMHEADERA
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pitem As HDITEMA
End Type

Public Type NMHEADERW
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pitem As HDITEMW
End Type

Public Type NMHDDISPINFOW
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As LongPtr
    cchTextMax As Long
    iImage As Long
    lParam As LongPtr
End Type

Public Type NMHDDISPINFOA
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As LongPtr
    cchTextMax As Long
    iImage As Long
    lParam As LongPtr
End Type

Public Type NMHDFILTERBTNCLICK
    hdr As NMHDR
    iItem As Long
    rc As RECT
End Type


Public Enum HeaderItemStateConstants
    hdisFocused = HDIS_FOCUSED
    hdisSelected = HDIS_SELECTED
    hdisHotTracked = HDIS_HOTTRACKED
End Enum



'--------------APIs----------------
