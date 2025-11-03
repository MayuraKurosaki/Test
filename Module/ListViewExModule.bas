Attribute VB_Name = "ListViewExModule"
Option Explicit

'--------------Constants----------------
'Class Names
Public Const WC_LISTVIEW = "SysListView32"

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

Public Const LVM_SETUNICODEFORMAT           As Long = CCM_SETUNICODEFORMAT
Public Const LVM_GETUNICODEFORMAT           As Long = CCM_GETUNICODEFORMAT


'Other Constants

Public Const WHEEL_DELTA                    As Long = 120


'Private Const WM_SETFOCUS = 7&
'Private Const WM_NOTIFY = &H4E&
'Private Const WM_MOUSEACTIVATE = &H21&
'Private Const WM_NCDESTROY = &H82&
'Private Const WM_SETFONT = &H30
'Private Const WM_VSCROLL = &H115

'Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
'Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
'Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Public Sub ListViewEditTest()
    ListViewEditView.Show
End Sub

Public Function Redirect(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
    Redirect = lv.WndProc(hwnd, uMsg, wParam, lParam)
End Function

'Public Function RedirectFrm(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
'                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
'    RedirectFrm = lv.WndProcFrm(hWnd, uMsg, wParam, lParam)
'End Function
'
'Public Function RedirectLV(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
'                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
'    RedirectLV = lv.WndProcLV(hWnd, uMsg, wParam, lParam)
'End Function

Public Sub ShowMessage(prm As LongPtr)
    ListViewEditView.Label2.Caption = prm
End Sub

Public Sub ShowMessage2(prm As LongPtr)
    ListViewEditView.Label3.Caption = prm
End Sub

Public Sub ShowMessage3(prm As Long)
    ListViewEditView.Label4.Caption = prm
End Sub

'Public Function WndProcLV(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
''    If hWnd = This.hwndLV Then
'        Select Case uMsg
'
''            Case WM_MOUSEACTIVATE
'''                If GetFocus() <> hWnd Then
'''                    acc.accSelect 1&
'''                    WndProc = MA_NOACTIVATE
'''                    Exit Function
'''                End If
'''
''            Case WM_NCDESTROY
''                RemoveSubClass
''                Exit Function
'''           '----------------------------
''            Case WM_KEYDOWN
'''                Select Case wParam
'''                    Case VK_RETURN
'''                        Exit Function
'''                    Case VK_UP, VK_DOWN
'''                        WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'''                        Exit Function
'''                    Case Else
'''                End Select
'''                Exit Function
''       '-----------------------------
'            Case WM_VSCROLL
'                Debug.Print "WM_VSCROLL"
'            Case Else
'        End Select
''    End If
'    WndProcLV = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'
'End Function
'
'Public Function WndProcFrm(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
''    If hWnd = This.hWndFrm Then
''            Call ShowMessage(hWnd)
''        Call ShowMessage2(This.hWndFrm)
'        Call ShowMessage3(uMsg)
'        Debug.Print "wndproc"
'        Select Case uMsg
'
'            Case WM_SETFOCUS
''                SetFocus This.hwndLV
'                Call ShowMessage(lParam)
'                Exit Function
'
'            Case WM_VSCROLL
'                Debug.Print "WM_VSCROLL"
'
'            Case WM_NCDESTROY
'                Debug.Print "WM_NCDESTROY"
'            Case WM_NOTIFY
''    ''            MoveMemory iCode, ByVal lParam + 8, 4
''    '
''    ''            Select Case iCode
'                Debug.Print "WM_NOTIFY"
'                Call ShowMessage(lParam)
'                Select Case lParam
''
'''                    Case LVN_ITEMCHANGED 'ItemSelected
'''                        MoveMemory nlv, ByVal lParam, LenB(nlv)
'''                        RaiseEvent ItemSelected(nlv.iItem, nlv.iSubItem)
'''                        Exit Function
'''
'''                    Case NM_CLICK 'ItemClick
'''                        MoveMemory nia, ByVal lParam, LenB(nia)
'''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'''                        Exit Function
'''
'''                    Case NM_DBLCLK
'''                        MoveMemory nia, ByVal lParam, LenB(nia)
'''                        RaiseEvent ItemClick(nia.iItem, nia.iSubItem)
'''                        Exit Function
''
''                    'Case ‚¢‚ë‚¢‚ë...
''                    '      :
''                    '      :
'                    Case Else
'
'                End Select
''
'        End Select
''
''    End If
'    WndProcFrm = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'
'End Function

'Private Function GetPtr(ByVal ptr As LongPtr) As LongPtr
'    GetPtr = ptr
'End Function

'Private Sub RemoveSubClass()
'    With This
'        If .hwndLV Then
'            RemoveWindowSubclass .hwndLV, .pfnLV, .hwndLV
'            .hwndLV = 0
'        End If
'
'        If .hWndFrm Then
'            RemoveWindowSubclass .hWndFrm, .pfnFrm, .hWndFrm
'            .hWndFrm = 0
'        End If
'    End With
'End Sub
