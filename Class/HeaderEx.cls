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

'--------------Compatible Properties----------------

'--------------Compatible Methods----------------
' #define Header_GetItemCount(hwndHD) \
'     (int)SNDMSG((hwndHD), HDM_GETITEMCOUNT, 0, 0L)
Public Function Header_GetItemCount(hwnd As LongPtr) As Long
  Header_GetItemCount = CLng(SendMessage(hwnd, HDM_GETITEMCOUNT, 0, 0))
End Function

' #define Header_InsertItem(hwndHD, i, phdi) \
'     (int)SNDMSG((hwndHD), HDM_INSERTITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM *)(phdi))
Public Function Header_InsertItem(hwnd As LongPtr, i As Long, phdi As HDITEMW) As Long
  Header_InsertItem = CLng(SendMessage(hwnd, HDM_INSERTITEM, i, phdi))
End Function

' #define Header_DeleteItem(hwndHD, i) \
'     (BOOL)SNDMSG((hwndHD), HDM_DELETEITEM, (WPARAM)(int)(i), 0L)
Public Function Header_DeleteItem(hwnd As LongPtr, i As Long) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_DELETEITEM, i, ByVal 0&))
  if ret <> 0 Then
      Header_DeleteItem = True
  Else
      Header_DeleteItem = False
  End If
End Function

' #define Header_GetItem(hwndHD, i, phdi) \
'     (BOOL)SNDMSG((hwndHD), HDM_GETITEM, (WPARAM)(int)(i), (LPARAM)(HD_ITEM *)(phdi))
Public Function Header_GetItem(hwnd As LongPtr, iItem As Long, phdi As HDITEM) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETITEM, iItem, phdi))
  If ret <> 0 Then
      Header_GetItem = True
  Else
      Header_GetItem = False
  End If
End Function

' #define Header_SetItem(hwndHD, i, phdi) \
'     (BOOL)SNDMSG((hwndHD), HDM_SETITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM *)(phdi))
Public Function Header_SetItem(hwnd As LongPtr, i As Long, phdi As HDITEM) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_SETITEM, ByVal i, phdi))
  If ret <> 0 Then
      Header_SetItem = True
  Else
      Header_SetItem = False
  End If
End Function

' #define Header_Layout(hwndHD, playout) \
'     (BOOL)SNDMSG((hwndHD), HDM_LAYOUT, 0, (LPARAM)(HD_LAYOUT *)(playout))
Public Function Header_Layout(hwnd As LongPtr, playout As HDLAYOUT) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_LAYOUT, 0, playout))
  If ret <> 0 Then
      Header_Layout = True
  Else
      Header_Layout = False
  End If
End Function

' #define Header_GetItemRect(hwnd, iItem, lprc) \
'         (BOOL)SNDMSG((hwnd), HDM_GETITEMRECT, (WPARAM)(iItem), (LPARAM)(lprc))
Public Function Header_GetItemRect(hwnd As LongPtr, iItem As Long, lprc As RECT) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETITEMRECT, iItem, lprc))
  If ret <> 0 Then
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
Public Function Header_SetStateImageList(hwnd As LongPtr, himl As LongPtr As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_SETIMAGELIST, HDSIL_STATE, ByVal himl))
  If ret <> 0 Then
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
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETIMAGELIST, HDSIL_STATE, ByVal 0&))
  If ret <> 0 Then
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
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETORDERARRAY, iCount, lpi))
  If ret <> 0 Then
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
Public Function Header_SetOrderArray(hwnd As LongPtr, iCount As Long, lpi As LongPtr As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_SETORDERARRAY, iCount, ByVal lpi))
  If ret <> 0 Then
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
  Header_SetHotDivider = Clng(SendMessage(hwnd, HDM_SETHOTDIVIDER, fPos, ByVal dw))
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
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_SETUNICODEFORMAT, fUnicode, ByVal 0&))
  If ret <> 0 Then
      Header_SetUnicodeFormat = True
  Else
      Header_SetUnicodeFormat = False
  End If
End Function

' #define Header_GetUnicodeFormat(hwnd)  \
'     (BOOL)SNDMSG((hwnd), HDM_GETUNICODEFORMAT, 0, 0)
Public Function Header_GetUnicodeFormat(hwnd As LongPtr) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETUNICODEFORMAT, 0, ByVal 0&))
  If ret <> 0 Then
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
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETITEMDROPDOWNRECT, iItem, lprc))
  If ret <> 0 Then
      Header_GetItemDropDownRect = True
  Else
      Header_GetItemDropDownRect = False
  End If
End Function

' #define Header_GetOverflowRect(hwnd, lprc) \
'         (BOOL)SNDMSG((hwnd), HDM_GETOVERFLOWRECT, 0, (LPARAM)(lprc))
Public Function Header_GetOverflowRect(hwnd As LongPtr, lprc As RECT) As Boolean
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_GETOVERFLOWRECT, 0, lprc))
  If ret <> 0 Then
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
  Dim ret As Long
  ret = CLng(SendMessage(hwnd, HDM_SETFOCUSEDITEM, 0, ByVal iItem))
  If ret <> 0 Then
      Header_SetFocusedItem = True
  Else
      Header_SetFocusedItem = False
  End If
End Function

Public Function GetHDItemlParam(hwnd As LongPtr, i As Long) As Long
  Dim tHDI As HDITEM
  tHDI.mask = HDI_LPARAM
  If Header_GetItem(hwnd, i, tHDI) Then
      GetHDItemlParam = tHDI.lParam
  End If
End Function