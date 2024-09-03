Attribute VB_Name = "HeaderControlModule"
Option Explicit

'Contains all definitions and macros for the header control
'as of common controls v6.0
'Macros are useful because even in cases where they're doing
'something like setting struct members, in many cases it's
'completely random whether a parameter should be passed as
'a wParam or an lParam


'header class id's
Public Const HEADER32_CLASS   As String = "SysHeader32"
Public Const HEADER_CLASS     As String = "SysHeader"

'header info

Public Enum HDMASK
    HDI_WIDTH = &H1
    HDI_HEIGHT = HDI_WIDTH
    HDI_TEXT = &H2
    HDI_FORMAT = &H4
    HDI_LPARAM = &H8
    HDI_BITMAP = &H10
    HDI_IMAGE = &H20
    HDI_DI_SETITEM = &H40
    HDI_ORDER = &H80
    '5.0
    HDI_FILTER = &H100
    '6.0
    HDI_STATE = &H200
End Enum

Public Enum HeaderStyles
    HDS_HORZ = &H0
    HDS_BUTTONS = &H2
    HDS_HIDDEN = &H8
    HDS_HOTTRACK = &H4 ' v 4.70
    HDS_DRAGDROP = &H40 ' v 4.70
    HDS_FULLDRAG = &H80
    HDS_FILTERBAR = &H100 ' v 5.0
    HDS_FLAT = &H200 ' v 5.1
    HDS_CHECKBOXES = &H400 '6.0
    HDS_NOSIZING = &H800
    HDS_OVERFLOW = &H1000
End Enum
Public Enum HeaderHitTestFlags
    HHT_NOWHERE = &H1
    HHT_ONHEADER = &H2
    HHT_ONDIVIDER = &H4
    HHT_ONDIVOPEN = &H8
'#if (_WIN32_IE >= =&h0500)
    HHT_ONFILTER = &H10
    HHT_ONFILTERBUTTON = &H20
'#End If
    HHT_ABOVE = &H100
    HHT_BELOW = &H200
    HHT_TORIGHT = &H400
    HHT_TOLEFT = &H800
'#if _WIN32_WINNT >= =&h0600
    HHT_ONITEMSTATEICON = &H1000
    HHT_ONDROPDOWN = &H2000
    HHT_ONOVERFLOW = &H4000
End Enum
Public Type HDHITTESTINFO
    pt As POINT
    Flags As HeaderHitTestFlags
    iItem As Long
End Type
Public Enum HeaderImageListFlags
    HDSIL_NORMAL = 0
    HDSIL_STATE = 1
End Enum

Public Const HDN_FIRST As Long = -300&
Public Const HDN_ITEMCLICK = (HDN_FIRST - 2)
Public Const HDN_DIVIDERDBLCLICK = (HDN_FIRST - 5)
Public Const HDN_BEGINTRACK = (HDN_FIRST - 6)
Public Const HDN_ENDTRACK = (HDN_FIRST - 7)
Public Const HDN_TRACK = (HDN_FIRST - 8)
Public Const HDN_GETDISPINFO = (HDN_FIRST - 9)
Public Const HDN_ITEMCHANGING As Long = (HDN_FIRST - 0)
Public Const HDN_ITEMDBLCLICK As Long = (HDN_FIRST - 3)
Public Const HDN_ITEMCHANGINGA = (HDN_FIRST - 0)
Public Const HDN_ITEMCHANGINGW = (HDN_FIRST - 20)
Public Const HDN_ITEMCHANGEDA = (HDN_FIRST - 1)
Public Const HDN_ITEMCHANGEDW = (HDN_FIRST - 21)
Public Const HDN_ITEMCLICKA = (HDN_FIRST - 2)
Public Const HDN_ITEMCLICKW = (HDN_FIRST - 22)
Public Const HDN_ITEMDBLCLICKA = (HDN_FIRST - 3)
Public Const HDN_ITEMDBLCLICKW = (HDN_FIRST - 23)
Public Const HDN_DIVIDERDBLCLICKA = (HDN_FIRST - 5)
Public Const HDN_DIVIDERDBLCLICKW = (HDN_FIRST - 25)
Public Const HDN_BEGINTRACKA = (HDN_FIRST - 6)
Public Const HDN_BEGINTRACKW = (HDN_FIRST - 26)
Public Const HDN_ENDTRACKA = (HDN_FIRST - 7)
Public Const HDN_ENDTRACKW = (HDN_FIRST - 27)
Public Const HDN_TRACKA = (HDN_FIRST - 8)
Public Const HDN_TRACKW = (HDN_FIRST - 28)
Public Const HDN_GETDISPINFOA = (HDN_FIRST - 9)
Public Const HDN_GETDISPINFOW = (HDN_FIRST - 29)
Public Const HDN_BEGINDRAG = (HDN_FIRST - 10)
Public Const HDN_ENDDRAG = (HDN_FIRST - 11)
Public Const HDN_FILTERCHANGE = (HDN_FIRST - 12)
Public Const HDN_FILTERBTNCLICK = (HDN_FIRST - 13)
'#If (WIN32_IE > 600) Then
Public Const HDN_BEGINFILTEREDIT = (HDN_FIRST - 14)
Public Const HDN_ENDFILTEREDIT = (HDN_FIRST - 15)
Public Const HDN_ITEMSTATEICONCLICK = (HDN_FIRST - 16)
Public Const HDN_ITEMKEYDOWN = (HDN_FIRST - 17)
Public Const HDN_DROPDOWN = (HDN_FIRST - 18)
Public Const HDN_OVERFLOWCLICK = (HDN_FIRST - 19)
'#End If

Public Const HDM_FIRST As Long = &H1200
Public Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
Public Const HDM_INSERTITEMA = (HDM_FIRST + 1)
Public Const HDM_DELETEITEM = (HDM_FIRST + 2)
Public Const HDM_GETITEMA = (HDM_FIRST + 3)
Public Const HDM_SETITEMA = (HDM_FIRST + 4)
Public Const HDM_LAYOUT = (HDM_FIRST + 5)
Public Const HDM_HITTEST = (HDM_FIRST + 6)
Public Const HDM_GETITEMRECT = (HDM_FIRST + 7)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Public Const HDM_GETIMAGELIST = (HDM_FIRST + 9)
Public Const HDM_INSERTITEMW = (HDM_FIRST + 10)
Public Const HDM_GETITEMW = (HDM_FIRST + 11)
Public Const HDM_SETITEMW = (HDM_FIRST + 12)

Public Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
Public Const HDM_CREATEDRAGIMAGE = (HDM_FIRST + 16)      '// wparam = which item (by index)
Public Const HDM_GETORDERARRAY = (HDM_FIRST + 17)
Public Const HDM_SETORDERARRAY = (HDM_FIRST + 18)
Public Const HDM_SETHOTDIVIDER = (HDM_FIRST + 19)
Public Const HDM_SETBITMAPMARGIN = (HDM_FIRST + 20)
Public Const HDM_GETBITMAPMARGIN = (HDM_FIRST + 21)
Public Const HDM_SETFILTERCHANGETIMEOUT = (HDM_FIRST + 22)
Public Const HDM_EDITFILTER = (HDM_FIRST + 23)
Public Const HDM_CLEARFILTER = (HDM_FIRST + 24)
Public Const HDM_GETITEMDROPDOWNRECT = (HDM_FIRST + 25) ' // rect of item's drop down button
Public Const HDM_GETOVERFLOWRECT = (HDM_FIRST + 26) '// rect of overflow button
Public Const HDM_GETFOCUSEDITEM = (HDM_FIRST + 27)
Public Const HDM_SETFOCUSEDITEM = (HDM_FIRST + 28)
Public Const HDM_TRANSLATEACCELERATOR = &H461  ' CCM_TRANSLATEACCELERATOR

Public Const HDM_GETITEM = HDM_GETITEMA
Public Const HDM_SETITEM = HDM_SETITEMA
Public Const HDM_INSERTITEM = HDM_INSERTITEMA
Public Const HDM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
Public Const HDM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
'#define Header_GetItemDropDownRect(hwnd, iItem, lprc) \
'        (BOOL)SNDMSG((hwnd), HDM_GETITEMDROPDOWNRECT, (WPARAM)(iItem), (LPARAM)(lprc))

'#define Header_GetOverflowRect(hwnd, lprc) \
'        (BOOL)SNDMSG((hwnd), HDM_GETOVERFLOWRECT, 0, (LPARAM)(lprc))
'
'#define Header_GetFocusedItem(hwnd) \
'        (int)SNDMSG((hwnd), HDM_GETFOCUSEDITEM, (WPARAM)(0), (LPARAM)(0))


'#End if
' HDITEM fmt
Public Enum HDITEM_FMT
    HDF_LEFT = 0
    HDF_RIGHT = 1
    HDF_CENTER = 2
    HDF_JUSTIFYMASK = &H3
    HDF_RTLREADING = 4
    HDF_BITMAP = &H2000
    HDF_STRING = &H4000
    HDF_OWNERDRAW = &H8000
    '3.0
    HDF_IMAGE = &H800
    HDF_BITMAP_ON_RIGHT = &H1000
    '5.0
    HDF_SORTUP = &H400
    HDF_SORTDOWN = &H200
    '6.0
    HDF_CHECKBOX = &H40
    HDF_CHECKED = &H80
    HDF_FIXEDWIDTH = &H100
    HDF_SPLITBUTTON = &H1000000
End Enum
Public Const HDFT_ISSTRING = &H0           '// HD_ITEM.pvFilter points to a HD_TEXTFILTER
Public Const HDFT_ISNUMBER = &H1           '// HD_ITEM.pvFilter points to a INT
Public Const HDFT_ISDATE = &H2
Public Const HDFT_HASNOVALUE = &H8000      '// clear the filter, by setting this bit

Public Const HDIS_FOCUSED = &H1

' Header Item Type

Public Type HDITEM
    mask As HDMASK
    CXY As Long
    pszText As String
    hBm As Long
    cchTextMax As Long
    fmt As HDITEM_FMT
    lParam As Long
    iImage As Long
    iOrder As Long
'#If (WIN32_IE >= &H500) then
    type As Long
    pvFilter As Long
'#If (WIN32_IE >= &H600) then
    State As Long
End Type
Public Type HDITEMW
    mask As HDMASK
    CXY As Long
    pszText As Long
    hBm As Long
    cchTextMax As Long
    fmt As HDITEM_FMT
    lParam As Long
'#If (WIN32_IE >= &H300) then
    iImage As Long
    iOrder As Long
'#If (WIN32_IE >= &H500) then
    type As Long
    pvFilter As Long
'#If (WIN32_IE >= &H600) then
    State As Long
End Type
Public Type HD_TEXTFILTERA
    pszText As String
    cchTextMax As Long
End Type
Public Type HD_TEXTFILTERW
    pszText  As Long
    cchTextMax As Long
End Type
Public Type HDLAYOUT
    prc As RECT
    pwpos As Long 'WINDOWPOS
End Type

Public Type NMHEADERX
     hdr As NMHDR
     iItem As Long
     iButton As Long
     pItem As HDITEMW ' HDITEM FAR* pItem
End Type
Public Type HD_NOTIFY
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pItem As HDITEM
End Type
Public Type NMHDDISPINFOW
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type
Public Type NMHDDISPINFOA
    hdr As NMHDR
    iItem As Long
    mask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type
Public Type NMHDFILTERBTNCLICK
    hdr As NMHDR
    iItem As Long
    rc As RECT
End Type

        



Public Function Header_GetItem(hwndHD As Long, iItem As Long, phdi As HDITEM) As Boolean
  Header_GetItem = SendMessage(hwndHD, HDM_GETITEM, iItem, phdi)
End Function

Public Function Header_SetItem(hwndHD As Long, i As Long, phdi As HDITEM) As Boolean
  Header_SetItem = SendMessage(hwndHD, HDM_SETITEMW, ByVal i, phdi)
End Function
 
Public Function Header_GetItemCount(hWnd As Long) As Long

Header_GetItemCount = SendMessage(hWnd, HDM_GETITEMCOUNT, 0, 0)
End Function

Public Function Header_InsertItem(hWnd As Long, i As Long, phdi As HDITEMW) As Long
'#define Header_InsertItem(hwndHD, i, phdi) \
'    (int)SNDMSG((hwndHD), HDM_INSERTITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM *)(phdi))
Header_InsertItem = SendMessage(hWnd, HDM_INSERTITEM, i, phdi)
End Function
Public Function Header_DeleteItem(hWnd As Long, i As Long) As Long
'#define Header_DeleteItem(hwndHD, i) \
'    (BOOL)SNDMSG((hwndHD), HDM_DELETEITEM, (WPARAM)(int)(i), 0L)
Header_DeleteItem = SendMessage(hWnd, HDM_DELETEITEM, i, ByVal 0&)
End Function
Public Function Header_Layout(hWnd As Long, playout As HDLAYOUT) As Long
'#define Header_Layout(hwndHD, playout) \
'    (BOOL)SNDMSG((hwndHD), HDM_LAYOUT, 0, (LPARAM)(HD_LAYOUT *)(playout))
Header_Layout = SendMessage(hWnd, HDM_LAYOUT, 0, playout)
End Function
Public Function Header_GetItemRect(hWnd As Long, iItem As Long, lprc As RECT) As Long
'#define Header_GetItemRect(hwnd, iItem, lprc) \
'        (BOOL)SNDMSG((hwnd), HDM_GETITEMRECT, (WPARAM)(iItem), (LPARAM)(lprc))
Header_GetItemRect = SendMessage(hWnd, HDM_GETITEMRECT, iItem, lprc)
End Function
Public Function Header_SetImageList(hWnd As Long, himl As Long) As Long
'#define Header_SetImageList(hwnd, himl) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_SETIMAGELIST, HDSIL_NORMAL, (LPARAM)(himl))
Header_SetImageList = SendMessage(hWnd, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal himl)
End Function
Public Function Header_SetStateImageList(hWnd As Long, himl As Long) As Long
'#define Header_SetStateImageList(hwnd, himl) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_SETIMAGELIST, HDSIL_STATE, (LPARAM)(himl))
Header_SetStateImageList = SendMessage(hWnd, HDM_SETIMAGELIST, HDSIL_STATE, ByVal himl)
End Function
Public Function Header_GetImageList(hWnd As Long) As Long
'#define Header_GetImageList(hwnd) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_GETIMAGELIST, HDSIL_NORMAL, 0)
Header_GetImageList = SendMessage(hWnd, HDM_GETIMAGELIST, HDSIL_NORMAL, ByVal 0&)
End Function
Public Function Header_GetStateImageList(hWnd As Long) As Long
'#define Header_GetImageList(hwnd) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_GETIMAGELIST, HDSIL_STATE, 0)
Header_GetStateImageList = SendMessage(hWnd, HDM_GETIMAGELIST, HDSIL_STATE, ByVal 0&)
End Function
Public Function Header_OrderToIndex(hWnd As Long, i As Long) As Long
'#define Header_OrderToIndex(hwnd, i) \
'        (int)SNDMSG((hwnd), HDM_ORDERTOINDEX, (WPARAM)(i), 0)
Header_OrderToIndex = SendMessage(hWnd, HDM_ORDERTOINDEX, i, ByVal 0&)
End Function
Public Function Header_CreateDragImage(hWnd As Long, i As Long) As Long
'#define Header_CreateDragImage(hwnd, i) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_CREATEDRAGIMAGE, (WPARAM)(i), 0)
Header_CreateDragImage = SendMessage(hWnd, HDM_CREATEDRAGIMAGE, i, ByVal 0&)
End Function
Public Function Header_GetOrderArray(hWnd As Long, iCount As Long, lpi As Long) As Long
'#define Header_GetOrderArray(hwnd, iCount, lpi) \
'        (BOOL)SNDMSG((hwnd), HDM_GETORDERARRAY, (WPARAM)(iCount), (LPARAM)(lpi))
Header_GetOrderArray = SendMessage(hWnd, HDM_GETORDERARRAY, iCount, lpi)
End Function
Public Function Header_SetOrderArray(hWnd As Long, iCount As Long, lpi As Long) As Long
'#define Header_SetOrderArray(hwnd, iCount, lpi) \
'        (BOOL)SNDMSG((hwnd), HDM_SETORDERARRAY, (WPARAM)(iCount), (LPARAM)(lpi))
'// lparam = int array of size HDM_GETITEMCOUNT
'// the array specifies the order that all items should be displayed.
'// e.g.  { 2, 0, 1}
'// says the index 2 item should be shown in the 0ths position
'//      index 0 should be shown in the 1st position
'//      index 1 should be shown in the 2nd position
'

Header_SetOrderArray = SendMessage(hWnd, HDM_SETORDERARRAY, iCount, ByVal lpi)
End Function
Public Function Header_SetHotDivider(hWnd As Long, fPos As Long, dw As Long) As Long
'#define Header_SetHotDivider(hwnd, fPos, dw) \
'        (int)SNDMSG((hwnd), HDM_SETHOTDIVIDER, (WPARAM)(fPos), (LPARAM)(dw))
Header_SetHotDivider = SendMessage(hWnd, HDM_SETHOTDIVIDER, fPos, ByVal dw)
End Function
Public Function Header_SetBitmapMargin(hWnd As Long, iWidth As Long) As Long
'#define Header_SetBitmapMargin(hwnd, iWidth) \
'        (int)SNDMSG((hwnd), HDM_SETBITMAPMARGIN, (WPARAM)(iWidth), 0)
Header_SetBitmapMargin = SendMessage(hWnd, HDM_SETBITMAPMARGIN, iWidth, ByVal 0&)
End Function
Public Function Header_GetBitmapMargin(hWnd As Long) As Long
'#define Header_GetBitmapMargin(hwnd) \
'        (int)SNDMSG((hwnd), HDM_GETBITMAPMARGIN, 0, 0)
Header_GetBitmapMargin = SendMessage(hWnd, HDM_GETBITMAPMARGIN, 0, ByVal 0&)
End Function
Public Function Header_SetUnicodeFormat(hWnd As Long, fUnicode As Long) As Long
'#define Header_SetUnicodeFormat(hwnd, fUnicode)  \
'    (BOOL)SNDMSG((hwnd), HDM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)
Header_SetUnicodeFormat = SendMessage(hWnd, HDM_SETUNICODEFORMAT, fUnicode, ByVal 0&)
End Function
Public Function Header_GetUnicodeFormat(hWnd As Long) As Long
'#define Header_GetUnicodeFormat(hwnd)  \
'    (BOOL)SNDMSG((hwnd), HDM_GETUNICODEFORMAT, 0, 0)
Header_GetUnicodeFormat = SendMessage(hWnd, HDM_GETUNICODEFORMAT, 0, ByVal 0&)
End Function
Public Function Header_SetFilterChangeTimeout(hWnd As Long, i As Long) As Long
'#define Header_SetFilterChangeTimeout(hwnd, i) \
'        (int)SNDMSG((hwnd), HDM_SETFILTERCHANGETIMEOUT, 0, (LPARAM)(i))
Header_SetFilterChangeTimeout = SendMessage(hWnd, HDM_SETFILTERCHANGETIMEOUT, 0, ByVal i)
End Function
Public Function Header_EditFilter(hWnd As Long, i As Long, fDiscardChanges As Long) As Long
'#define Header_EditFilter(hwnd, i, fDiscardChanges) \
'        (int)SNDMSG((hwnd), HDM_EDITFILTER, (WPARAM)(i), MAKELPARAM(fDiscardChanges, 0))
Header_EditFilter = SendMessage(hWnd, HDM_EDITFILTER, i, ByVal fDiscardChanges)
End Function
Public Function Header_ClearFilter(hWnd As Long, i As Long) As Long
'#define Header_ClearFilter(hwnd, i) \
'        (int)SNDMSG((hwnd), HDM_CLEARFILTER, (WPARAM)(i), 0)
Header_ClearFilter = SendMessage(hWnd, HDM_CLEARFILTER, i, ByVal 0&)
End Function
Public Function Header_ClearAllFilters(hWnd As Long) As Long
'#define Header_ClearAllFilters(hwnd) \
'        (int)SNDMSG((hwnd), HDM_CLEARFILTER, (WPARAM)-1, 0)
Header_ClearAllFilters = SendMessage(hWnd, HDM_CLEARFILTER, -1, ByVal 0&)
End Function
Public Function Header_GetItemDropDownRect(hWnd As Long, iItem As Long, lprc As RECT) As Long
'#define Header_GetItemDropDownRect(hwnd, iItem, lprc) \
'        (BOOL)SNDMSG((hwnd), HDM_GETITEMDROPDOWNRECT, (WPARAM)(iItem), (LPARAM)(lprc))
Header_GetItemDropDownRect = SendMessage(hWnd, HDM_GETITEMDROPDOWNRECT, iItem, lprc)
End Function
Public Function Header_GetOverflowRect(hWnd As Long, lprc As RECT) As Long
'#define Header_GetOverflowRect(hwnd, lprc) \
'        (BOOL)SNDMSG((hwnd), HDM_GETOVERFLOWRECT, 0, (LPARAM)(lprc))
Header_GetOverflowRect = SendMessage(hWnd, HDM_GETOVERFLOWRECT, 0, lprc)
End Function
Public Function Header_GetFocusedItem(hWnd As Long) As Long
'#define Header_GetFocusedItem(hwnd) \
'        (int)SNDMSG((hwnd), HDM_GETFOCUSEDITEM, (WPARAM)(0), (LPARAM)(0))
Header_GetFocusedItem = SendMessage(hWnd, HDM_GETFOCUSEDITEM, 0, ByVal 0&)
End Function
Public Function Header_SetFocusedItem(hWnd As Long, iItem As Long) As Long
'#define Header_SetFocusedItem(hwnd, iItem) \
'        (BOOL)SNDMSG((hwnd), HDM_SETFOCUSEDITEM, (WPARAM)(0), (LPARAM)(iItem))
Header_SetFocusedItem = SendMessage(hWnd, HDM_SETFOCUSEDITEM, 0, ByVal iItem)
End Function

Public Function GetHDItemlParam(hWnd As Long, i As Long) As Long
Dim tHDI As HDITEM
tHDI.mask = HDI_LPARAM
If Header_GetItem(hWnd, i, tHDI) Then
    GetHDItemlParam = tHDI.lParam
End If

End Function


