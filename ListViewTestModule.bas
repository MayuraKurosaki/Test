Attribute VB_Name = "ListViewTestModule"
Option Explicit

Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long
Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

Const GWL_EXSTYLE = (-20&)
Const WS_EX_LAYERED = &H80000
Const LWA_COLORKEY = &H1

' messages
Public Const LVM_FIRST = &H1000
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)


Public Type POINTAPI
    x As Long
    y As Long
End Type

' LVM_HITTEST lParam
Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
    pt As POINTAPI
    Flags As LVHT_FLAGS
    iItem As Long
'#If (WIN32_IE >= &H300) Then
    iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
'#End If
'#If (WIN32_IE >= &H600) then
    iGroup As Long
'#End If
End Type

Public Enum LVHT_FLAGS
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

Public Function MakeTransparentFrame(Frame As Object) ', Optional Color As Long)
    Dim hWndFrame As LongPtr
'    Dim Opacity As Byte
    
    Call WindowFromAccessibleObject(Frame, hWndFrame)
'    Frame.SetFocus
'    hWndFrame = GetFocus
'    hWndFrame = Extraithandle(Frame)
   
'    If IsMissing(Color) Then Color = rgbWhite
'    Opacity = 200
    
    Call SetWindowLongPtr(hWndFrame, GWL_EXSTYLE, GetWindowLongPtr(hWndFrame, GWL_EXSTYLE) Or WS_EX_LAYERED)
    
'    Frame.BackColor = Color
    Call SetLayeredWindowAttributes(hWndFrame, Frame.BackColor, 0, LWA_COLORKEY)
'    Call Frame.ZOrder(0)
'Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

End Function
