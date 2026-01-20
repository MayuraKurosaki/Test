VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeaderEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Extended Header"
Option Explicit

'--------------Constants----------------
' Header Window Class Names
Private Const HEADER32_CLASS             As String = "SysHeader32"
Private Const HEADER_CLASS               As String = "SysHeader"

' Common Control Messages
Private Const HDM_FIRST                  As Long = &H1200
Private Const HDM_GETITEMCOUNT           As Long = (HDM_FIRST + 0)
Private Const HDM_INSERTITEMA            As Long = (HDM_FIRST + 1)
Private Const HDM_DELETEITEM             As Long = (HDM_FIRST + 2)
Private Const HDM_GETITEMA               As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA               As Long = (HDM_FIRST + 4)
Private Const HDM_LAYOUT                 As Long = (HDM_FIRST + 5)
Private Const HDM_HITTEST                As Long = (HDM_FIRST + 6)
Private Const HDM_GETITEMRECT            As Long = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST           As Long = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST           As Long = (HDM_FIRST + 9)
Private Const HDM_INSERTITEMW            As Long = (HDM_FIRST + 10)
Private Const HDM_GETITEMW               As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW               As Long = (HDM_FIRST + 12)

Private Const HDM_ORDERTOINDEX           As Long = (HDM_FIRST + 15)
Private Const HDM_CREATEDRAGIMAGE        As Long = (HDM_FIRST + 16)      '// wparam = which item (by index)
Private Const HDM_GETORDERARRAY          As Long = (HDM_FIRST + 17)
Private Const HDM_SETORDERARRAY          As Long = (HDM_FIRST + 18)
Private Const HDM_SETHOTDIVIDER          As Long = (HDM_FIRST + 19)
Private Const HDM_SETBITMAPMARGIN        As Long = (HDM_FIRST + 20)
Private Const HDM_GETBITMAPMARGIN        As Long = (HDM_FIRST + 21)
Private Const HDM_SETFILTERCHANGETIMEOUT As Long = (HDM_FIRST + 22)
Private Const HDM_EDITFILTER             As Long = (HDM_FIRST + 23)
Private Const HDM_CLEARFILTER            As Long = (HDM_FIRST + 24)
Private Const HDM_GETITEMDROPDOWNRECT    As Long = (HDM_FIRST + 25)      ' // rect of item's drop down button
Private Const HDM_GETOVERFLOWRECT        As Long = (HDM_FIRST + 26)      '// rect of overflow button
Private Const HDM_GETFOCUSEDITEM         As Long = (HDM_FIRST + 27)
Private Const HDM_SETFOCUSEDITEM         As Long = (HDM_FIRST + 28)
Private Const HDM_TRANSLATEACCELERATOR   As Long = &H461                 ' CCM_TRANSLATEACCELERATOR

Private Const HDM_GETITEM                As Long = HDM_GETITEMA
Private Const HDM_SETITEM                As Long = HDM_SETITEMA
Private Const HDM_INSERTITEM             As Long = HDM_INSERTITEMA
Private Const HDM_SETUNICODEFORMAT       As Long = CCM_SETUNICODEFORMAT
Private Const HDM_GETUNICODEFORMAT       As Long = CCM_GETUNICODEFORMAT

' Header Notification Codes
Private Const HDN_FIRST                  As Long = -300
Private Const HDN_ITEMCLICK              As Long = (HDN_FIRST - 2)
Private Const HDN_DIVIDERDBLCLICK        As Long = (HDN_FIRST - 5)
Private Const HDN_BEGINTRACK             As Long = (HDN_FIRST - 6)
Private Const HDN_ENDTRACK               As Long = (HDN_FIRST - 7)
Private Const HDN_TRACK                  As Long = (HDN_FIRST - 8)
Private Const HDN_GETDISPINFO            As Long = (HDN_FIRST - 9)
Private Const HDN_ITEMCHANGING           As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMDBLCLICK           As Long = (HDN_FIRST - 3)
Private Const HDN_ITEMCHANGINGA          As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCHANGINGW          As Long = (HDN_FIRST - 20)
Private Const HDN_ITEMCHANGEDA           As Long = (HDN_FIRST - 1)
Private Const HDN_ITEMCHANGEDW           As Long = (HDN_FIRST - 21)
Private Const HDN_ITEMCLICKA             As Long = (HDN_FIRST - 2)
Private Const HDN_ITEMCLICKW             As Long = (HDN_FIRST - 22)
Private Const HDN_ITEMDBLCLICKA          As Long = (HDN_FIRST - 3)
Private Const HDN_ITEMDBLCLICKW          As Long = (HDN_FIRST - 23)
Private Const HDN_DIVIDERDBLCLICKA       As Long = (HDN_FIRST - 5)
Private Const HDN_DIVIDERDBLCLICKW       As Long = (HDN_FIRST - 25)
Private Const HDN_BEGINTRACKA            As Long = (HDN_FIRST - 6)
Private Const HDN_BEGINTRACKW            As Long = (HDN_FIRST - 26)
Private Const HDN_ENDTRACKA              As Long = (HDN_FIRST - 7)
Private Const HDN_ENDTRACKW              As Long = (HDN_FIRST - 27)
Private Const HDN_TRACKA                 As Long = (HDN_FIRST - 8)
Private Const HDN_TRACKW                 As Long = (HDN_FIRST - 28)
Private Const HDN_GETDISPINFOA           As Long = (HDN_FIRST - 9)
Private Const HDN_GETDISPINFOW           As Long = (HDN_FIRST - 29)
Private Const HDN_BEGINDRAG              As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG                As Long = (HDN_FIRST - 11)
Private Const HDN_FILTERCHANGE           As Long = (HDN_FIRST - 12)
Private Const HDN_FILTERBTNCLICK         As Long = (HDN_FIRST - 13)
Private Const HDN_BEGINFILTEREDIT        As Long = (HDN_FIRST - 14)
Private Const HDN_ENDFILTEREDIT          As Long = (HDN_FIRST - 15)
Private Const HDN_ITEMSTATEICONCLICK     As Long = (HDN_FIRST - 16)
Private Const HDN_ITEMKEYDOWN            As Long = (HDN_FIRST - 17)
Private Const HDN_DROPDOWN               As Long = (HDN_FIRST - 18)
Private Const HDN_OVERFLOWCLICK          As Long = (HDN_FIRST - 19)

' Header Control Styles
Private Const HDS_HORZ                   As Long = &H0
Private Const HDS_BUTTONS                As Long = &H2
Private Const HDS_HOTTRACK               As Long = &H4
Private Const HDS_HIDDEN                 As Long = &H8
Private Const HDS_DRAGDROP               As Long = &H40
Private Const HDS_FULLDRAG               As Long = &H80
Private Const HDS_FILTERBAR              As Long = &H100
Private Const HDS_FLAT                   As Long = &H200
Private Const HDS_CHECKBOXES             As Long = &H400
Private Const HDS_NOSIZING               As Long = &H800
Private Const HDS_OVERFLOW               As Long = &H1000

' Header Item Format Constants
Private Const HDF_LEFT                   As Long = &H0           '// Same as LVCFMT_LEFT
Private Const HDF_RIGHT                  As Long = &H1           '// Same as LVCFMT_RIGHT
Private Const HDF_CENTER                 As Long = &H2           '// Same as LVCFMT_CENTER
Private Const HDF_JUSTIFYMASK            As Long = &H3           '// Same as LVCFMT_JUSTIFYMASK
Private Const HDF_RTLREADING             As Long = &H4           '// Same as LVCFMT_LEFT
Private Const HDF_CHECKBOX               As Long = &H40
Private Const HDF_CHECKED                As Long = &H80
Private Const HDF_FIXEDWIDTH             As Long = &H100         '// Can't resize the column; same as LVCFMT_FIXED_WIDTH
Private Const HDF_SORTDOWN               As Long = &H200
Private Const HDF_SORTUP                 As Long = &H400
Private Const HDF_IMAGE                  As Long = &H800         '// Same as LVCFMT_IMAGE
Private Const HDF_BITMAP_ON_RIGHT        As Long = &H1000        '// Same as LVCFMT_BITMAP_ON_RIGHT
Private Const HDF_BITMAP                 As Long = &H2000
Private Const HDF_STRING                 As Long = &H4000
Private Const HDF_OWNERDRAW              As Long = &H8000        '// Same as LVCFMT_COL_HAS_IMAGES
Private Const HDF_SPLITBUTTON            As Long = &H1000000     '// Column is a split button; same as LVCFMT_SPLITBUTTON

' Header Item Filter Type Constants
Private Const HDFT_ISSTRING              As Long = &H0           '// HD_ITEM.pvFilter points to a HD_TEXTFILTER
Private Const HDFT_ISNUMBER              As Long = &H1           '// HD_ITEM.pvFilter points to a INT
Private Const HDFT_ISDATE                As Long = &H2           '// HD_ITEM.pvFilter points to a DWORD (dos date)
Private Const HDFT_HASNOVALUE            As Long = &H8000        '// clear the filter, by setting this bit

' Header Item State Constants
Private Const HDIS_FOCUSED               As Long = &H1
Private Const HDIS_SELECTED              As Long = &H2
Private Const HDIS_HOTTRACKED            As Long = &H4

' Header Item Mask Constants
Private Const HDI_WIDTH                  As Long = &H1
Private Const HDI_HEIGHT                 As Long = HDI_WIDTH
Private Const HDI_TEXT                   As Long = &H2
Private Const HDI_FORMAT                 As Long = &H4
Private Const HDI_LPARAM                 As Long = &H8
Private Const HDI_BITMAP                 As Long = &H10
Private Const HDI_IMAGE                  As Long = &H20
Private Const HDI_DI_SETITEM             As Long = &H40
Private Const HDI_ORDER                  As Long = &H80
Private Const HDI_FILTER                 As Long = &H100
Private Const HDI_STATE                  As Long = &H200

' Header Hit Test Constants
Private Const HHT_NOWHERE                As Long = &H1
Private Const HHT_ONHEADER               As Long = &H2
Private Const HHT_ONDIVIDER              As Long = &H4
Private Const HHT_ONDIVOPEN              As Long = &H8
Private Const HHT_ONFILTER               As Long = &H10
Private Const HHT_ONFILTERBUTTON         As Long = &H20
Private Const HHT_ABOVE                  As Long = &H100
Private Const HHT_BELOW                  As Long = &H200
Private Const HHT_TORIGHT                As Long = &H400
Private Const HHT_TOLEFT                 As Long = &H800
Private Const HHT_ONITEMSTATEICON        As Long = &H1000
Private Const HHT_ONDROPDOWN             As Long = &H2000
Private Const HHT_ONOVERFLOW             As Long = &H4000


Private Const HDSIL_NORMAL               As Long = &H0
Private Const HDSIL_STATE                As Long = &H1

'--------------Enums----------------
Private Enum HeaderItemStateConstants
    hdisFocused = HDIS_FOCUSED
    hdisSelected = HDIS_SELECTED
    hdisHotTracked = HDIS_HOTTRACKED
End Enum

''ColumnAlignment constants
'Public Enum ListColumnAlignmentConstants
'        lvwColumnLeft = 0
'        lvwColumnRight = 1
'        lvwColumnCenter = 2
'End Enum



Private Type HDTEXTFILTERA
    pszText As LongPtr                      ' [in] pointer to the buffer containing the filter (ANSI)
    cchTextMax As Long                      ' [in] max size of buffer/edit control buffer
End Type

Private Type HDTEXTFILTERW
    pszText As LongPtr                      ' [in] pointer to the buffer containing the filter (UNICODE)
    cchTextMax As Long                      ' [in] max size of buffer/edit control buffer
End Type

Private Type HDITEMA
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
    State As Long
End Type

Private Type HDITEMW
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
    State As Long
End Type

Private Type HDLAYOUT
    prc As LongPtr
    pwpos As LongPtr
End Type

Private Type HDHITTESTINFO
    pt As POINT
    Flags As Long
    iItem As Long
End Type

Private Type NMHEADERA
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pitem As HDITEMA
End Type

Private Type NMHEADERW
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pitem As HDITEMW
End Type

Private Type NMHDDISPINFOW
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As LongPtr
    cchTextMax As Long
    iImage As Long
    lParam As LongPtr
End Type

Private Type NMHDDISPINFOA
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As LongPtr
    cchTextMax As Long
    iImage As Long
    lParam As LongPtr
End Type

Private Type NMHDFILTERBTNCLICK
    hdr As NMHDR
    iItem As Long
    rc As RECT
End Type

'--------------Compatible Properties----------------

'--------------Compatible Methods----------------
' #define Header_GetItemCount(hwndHD) \
'     (int)SNDMSG((hwndHD), HDM_GETITEMCOUNT, 0, 0L)
Public Function Header_GetItemCount(hwnd As LongPtr) As Long
  Header_GetItemCount = CLng(SendMessage(hwnd, HDM_GETITEMCOUNT, 0, 0))
End Function

' #define Header_InsertItem(hwndHD, i, phdi) \
'     (int)SNDMSG((hwndHD), HDM_INSERTITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM *)(phdi))
Public Function Header_InsertItem(hwnd As LongPtr, i As Long, phdi As HDITEMA) As Long
  Header_InsertItem = CLng(SendMessage(hwnd, HDM_INSERTITEM, i, phdi))
End Function

' #define Header_DeleteItem(hwndHD, i) \
'     (BOOL)SNDMSG((hwndHD), HDM_DELETEITEM, (WPARAM)(int)(i), 0L)
Public Function Header_DeleteItem(hwnd As LongPtr, i As Long) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_DELETEITEM, i, ByVal 0&))
  If Ret <> 0 Then
      Header_DeleteItem = True
  Else
      Header_DeleteItem = False
  End If
End Function

' #define Header_GetItem(hwndHD, i, phdi) \
'     (BOOL)SNDMSG((hwndHD), HDM_GETITEM, (WPARAM)(int)(i), (LPARAM)(HD_ITEM *)(phdi))
Public Function Header_GetItem(hwnd As LongPtr, iItem As Long, phdi As HDITEMA) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETITEM, iItem, phdi))
  If Ret <> 0 Then
      Header_GetItem = True
  Else
      Header_GetItem = False
  End If
End Function

' #define Header_SetItem(hwndHD, i, phdi) \
'     (BOOL)SNDMSG((hwndHD), HDM_SETITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM *)(phdi))
Public Function Header_SetItem(hwnd As LongPtr, i As Long, phdi As HDITEMA) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_SETITEM, ByVal i, phdi))
  If Ret <> 0 Then
      Header_SetItem = True
  Else
      Header_SetItem = False
  End If
End Function

' #define Header_Layout(hwndHD, playout) \
'     (BOOL)SNDMSG((hwndHD), HDM_LAYOUT, 0, (LPARAM)(HD_LAYOUT *)(playout))
Public Function Header_Layout(hwnd As LongPtr, playout As HDLAYOUT) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_LAYOUT, 0, playout))
  If Ret <> 0 Then
      Header_Layout = True
  Else
      Header_Layout = False
  End If
End Function

' #define Header_GetItemRect(hwnd, iItem, lprc) \
'         (BOOL)SNDMSG((hwnd), HDM_GETITEMRECT, (WPARAM)(iItem), (LPARAM)(lprc))
Public Function Header_GetItemRect(hwnd As LongPtr, iItem As Long, lprc As RECT) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETITEMRECT, iItem, lprc))
  If Ret <> 0 Then
      Header_GetItemRect = True
  Else
      Header_GetItemRect = False
  End If
End Function

' #define Header_SetImageList(hwnd, himl) \
'         (HIMAGELIST)SNDMSG((hwnd), HDM_SETIMAGELIST, HDSIL_NORMAL, (LPARAM)(himl))
Public Function Header_SetImageList(hwnd As LongPtr, himl As Long) As LongPtr
  Header_SetImageList = SendMessage(hwnd, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal himl)
End Function

' #define Header_SetStateImageList(hwnd, himl) \
'         (HIMAGELIST)SNDMSG((hwnd), HDM_SETIMAGELIST, HDSIL_STATE, (LPARAM)(himl))
Public Function Header_SetStateImageList(hwnd As LongPtr, himl As LongPtr) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_SETIMAGELIST, HDSIL_STATE, ByVal himl))
  If Ret <> 0 Then
      Header_SetStateImageList = True
  Else
      Header_SetStateImageList = False
  End If
End Function

' #define Header_GetImageList(hwnd) \
'         (HIMAGELIST)SNDMSG((hwnd), HDM_GETIMAGELIST, HDSIL_NORMAL, 0)
Public Function Header_GetImageList(hwnd As LongPtr) As LongPtr
  Header_GetImageList = SendMessage(hwnd, HDM_GETIMAGELIST, HDSIL_NORMAL, ByVal 0&)
End Function

' #define Header_GetStateImageList(hwnd) \
'         (HIMAGELIST)SNDMSG((hwnd), HDM_GETIMAGELIST, HDSIL_STATE, 0)
Public Function Header_GetStateImageList(hwnd As LongPtr) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETIMAGELIST, HDSIL_STATE, ByVal 0&))
  If Ret <> 0 Then
      Header_GetStateImageList = True
  Else
      Header_GetStateImageList = False
  End If
End Function

' #define Header_OrderToIndex(hwnd, i) \
'         (int)SNDMSG((hwnd), HDM_ORDERTOINDEX, (WPARAM)(i), 0)
Public Function Header_OrderToIndex(hwnd As LongPtr, i As Long) As Long
  Header_OrderToIndex = CLng(SendMessage(hwnd, HDM_ORDERTOINDEX, i, ByVal 0&))
End Function

' #define Header_CreateDragImage(hwnd, i) \
'         (HIMAGELIST)SNDMSG((hwnd), HDM_CREATEDRAGIMAGE, (WPARAM)(i), 0)
Public Function Header_CreateDragImage(hwnd As LongPtr, i As Long) As LongPtr
  Header_CreateDragImage = SendMessage(hwnd, HDM_CREATEDRAGIMAGE, i, ByVal 0&)
End Function

' #define Header_GetOrderArray(hwnd, iCount, lpi) \
'         (BOOL)SNDMSG((hwnd), HDM_GETORDERARRAY, (WPARAM)(iCount), (LPARAM)(lpi))
Public Function Header_GetOrderArray(hwnd As LongPtr, iCount As Long, lpi As LongPtr) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETORDERARRAY, iCount, lpi))
  If Ret <> 0 Then
      Header_GetOrderArray = True
  Else
      Header_GetOrderArray = False
  End If
End Function

' #define Header_SetOrderArray(hwnd, iCount, lpi) \
'         (BOOL)SNDMSG((hwnd), HDM_SETORDERARRAY, (WPARAM)(iCount), (LPARAM)(lpi))
' // lparam = int array of size HDM_GETITEMCOUNT
' // the array specifies the order that all items should be displayed.
' // e.g.  { 2, 0, 1}
' // says the index 2 item should be shown in the 0ths position
' //      index 0 should be shown in the 1st position
' //      index 1 should be shown in the 2nd position
Public Function Header_SetOrderArray(hwnd As LongPtr, iCount As Long, lpi As LongPtr) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_SETORDERARRAY, iCount, ByVal lpi))
  If Ret <> 0 Then
      Header_SetOrderArray = True
  Else
      Header_SetOrderArray = False
  End If
End Function

' #define Header_SetHotDivider(hwnd, fPos, dw) \
'         (int)SNDMSG((hwnd), HDM_SETHOTDIVIDER, (WPARAM)(fPos), (LPARAM)(dw))
' // convenience message for external dragdrop
' // wParam = BOOL  specifying whether the lParam is a dwPos of the cursor
' //              position or the index of which divider to hotlight
' // lParam = depends on wParam  (-1 and wParm = FALSE turns off hotlight)
Public Function Header_SetHotDivider(hwnd As LongPtr, fPos As Long, dw As Long) As Long
  Header_SetHotDivider = CLng(SendMessage(hwnd, HDM_SETHOTDIVIDER, fPos, ByVal dw))
End Function

' #define Header_SetBitmapMargin(hwnd, iWidth) \
'         (int)SNDMSG((hwnd), HDM_SETBITMAPMARGIN, (WPARAM)(iWidth), 0)
Public Function Header_SetBitmapMargin(hwnd As LongPtr, iWidth As Long) As Long
  Header_SetBitmapMargin = CLng(SendMessage(hwnd, HDM_SETBITMAPMARGIN, iWidth, ByVal 0&))
End Function

' #define Header_GetBitmapMargin(hwnd) \
'         (int)SNDMSG((hwnd), HDM_GETBITMAPMARGIN, 0, 0)
Public Function Header_GetBitmapMargin(hwnd As LongPtr) As Long
  Header_GetBitmapMargin = CLng(SendMessage(hwnd, HDM_GETBITMAPMARGIN, 0, ByVal 0&))
End Function

' #define Header_SetUnicodeFormat(hwnd, fUnicode)  \
'     (BOOL)SNDMSG((hwnd), HDM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)
Public Function Header_SetUnicodeFormat(hwnd As LongPtr, fUnicode As Long) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_SETUNICODEFORMAT, fUnicode, ByVal 0&))
  If Ret <> 0 Then
      Header_SetUnicodeFormat = True
  Else
      Header_SetUnicodeFormat = False
  End If
End Function

' #define Header_GetUnicodeFormat(hwnd)  \
'     (BOOL)SNDMSG((hwnd), HDM_GETUNICODEFORMAT, 0, 0)
Public Function Header_GetUnicodeFormat(hwnd As LongPtr) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETUNICODEFORMAT, 0, ByVal 0&))
  If Ret <> 0 Then
      Header_GetUnicodeFormat = True
  Else
      Header_GetUnicodeFormat = False
  End If
End Function

' #define Header_SetFilterChangeTimeout(hwnd, i) \
'         (int)SNDMSG((hwnd), HDM_SETFILTERCHANGETIMEOUT, 0, (LPARAM)(i))
Public Function Header_SetFilterChangeTimeout(hwnd As LongPtr, i As Long) As Long
  Header_SetFilterChangeTimeout = CLng(SendMessage(hwnd, HDM_SETFILTERCHANGETIMEOUT, 0, ByVal i))
End Function

' #define Header_EditFilter(hwnd, i, fDiscardChanges) \
'         (int)SNDMSG((hwnd), HDM_EDITFILTER, (WPARAM)(i), MAKELPARAM(fDiscardChanges, 0))
Public Function Header_EditFilter(hwnd As LongPtr, i As Long, fDiscardChanges As Long) As Long
  Header_EditFilter = CLng(SendMessage(hwnd, HDM_EDITFILTER, i, ByVal fDiscardChanges))
End Function

' // Clear filter takes -1 as a column value to indicate that all
' // the filter should be cleared.  When this happens you will
' // only receive a single filter changed notification.
' #define Header_ClearFilter(hwnd, i) \
'         (int)SNDMSG((hwnd), HDM_CLEARFILTER, (WPARAM)(i), 0)
Public Function Header_ClearFilter(hwnd As LongPtr, i As Long) As Long
  Header_ClearFilter = CLng(SendMessage(hwnd, HDM_CLEARFILTER, i, ByVal 0&))
End Function

' #define Header_ClearAllFilters(hwnd) \
'         (int)SNDMSG((hwnd), HDM_CLEARFILTER, (WPARAM)-1, 0)
Public Function Header_ClearAllFilters(hwnd As LongPtr) As Long
  Header_ClearAllFilters = CLng(SendMessage(hwnd, HDM_CLEARFILTER, -1, ByVal 0&))
End Function

' #define Header_GetItemDropDownRect(hwnd, iItem, lprc) \
'         (BOOL)SNDMSG((hwnd), HDM_GETITEMDROPDOWNRECT, (WPARAM)(iItem), (LPARAM)(lprc))
Public Function Header_GetItemDropDownRect(hwnd As LongPtr, iItem As Long, lprc As RECT) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETITEMDROPDOWNRECT, iItem, lprc))
  If Ret <> 0 Then
      Header_GetItemDropDownRect = True
  Else
      Header_GetItemDropDownRect = False
  End If
End Function

' #define Header_GetOverflowRect(hwnd, lprc) \
'         (BOOL)SNDMSG((hwnd), HDM_GETOVERFLOWRECT, 0, (LPARAM)(lprc))
Public Function Header_GetOverflowRect(hwnd As LongPtr, lprc As RECT) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_GETOVERFLOWRECT, 0, lprc))
  If Ret <> 0 Then
      Header_GetOverflowRect = True
  Else
      Header_GetOverflowRect = False
  End If
End Function

' #define Header_GetFocusedItem(hwnd) \
'         (int)SNDMSG((hwnd), HDM_GETFOCUSEDITEM, (WPARAM)(0), (LPARAM)(0))
Public Function Header_GetFocusedItem(hwnd As LongPtr) As Long
  Header_GetFocusedItem = CLng(SendMessage(hwnd, HDM_GETFOCUSEDITEM, 0, ByVal 0&))
End Function

' #define Header_SetFocusedItem(hwnd, iItem) \
'         (BOOL)SNDMSG((hwnd), HDM_SETFOCUSEDITEM, (WPARAM)(0), (LPARAM)(iItem))
Public Function Header_SetFocusedItem(hwnd As LongPtr, iItem As Long) As Boolean
  Dim Ret As Long
  Ret = CLng(SendMessage(hwnd, HDM_SETFOCUSEDITEM, 0, ByVal iItem))
  If Ret <> 0 Then
      Header_SetFocusedItem = True
  Else
      Header_SetFocusedItem = False
  End If
End Function

Public Function GetHDItemlParam(hwnd As LongPtr, i As Long) As LongPtr
  Dim tHDI As HDITEMA
  tHDI.mask = HDI_LPARAM
  If Header_GetItem(hwnd, i, tHDI) Then
      GetHDItemlParam = tHDI.lParam
  End If
End Function
