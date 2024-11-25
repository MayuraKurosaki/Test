VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlTest 
   Caption         =   "UserForm1"
   ClientHeight    =   6800
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10540
   OleObjectBlob   =   "ControlTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ControlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

Private OnMouseOver As Boolean
Private OnFocus As Boolean

Private TextEx As TextBoxEx

Implements IControlEvent

Private Type Field
    Controls As ControlEvents
    PrevControl As ControlEvent
End Type

Private This As Field

Function ConvertedUTF8(ByRef Data() As Byte) As String
    Dim TotalBuffer() As Byte, Converted() As Byte, I As Long
    
    
    I = I + UBound(Data) + 1
    ReDim Preserve TotalBuffer(I - 1)
    RtlMoveMemory TotalBuffer(I - UBound(Data) - 1), Data(0), UBound(Data) + 1
    
    Dim lSize As Long
    lSize = MultiByteToWideChar(CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, ByVal 0&, 0&)
    
    ReDim Converted(lSize * 2 - 1)
    MultiByteToWideChar CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, Converted(0), lSize
    ConvertedUTF8 = Converted
End Function

Private Function CreateListView(hWndParent As Long, iid As Long, dwStyle As Long, dwExStyle As Long) As Long
    Dim rc As RECT
    Dim hwndLV As Long
    
    Call GetClientRect(hWndParent, rc)
    hwndLV = CreateWindowEx(dwExStyle, WC_LISTVIEW, "", _
                                                  dwStyle, 218, 2, 650, rc.Bottom - 30, _
                                                  hWndParent, iid, App.hInstance, 0)
     ListView_SetItemCount hwndLV, UBound(VLItems) + 1
    CreateListView = hwndLV
End Function

Private Sub InitListView()
    Dim dwStyle As Long, dwStyle2 As Long
    Dim lvcol As LVCOLUMNW
    Dim I As Long
    Dim rc As RECT
    
    hLVVG = CreateListView(Me.hwnd, IDD_LISTVIEW, _
                      LVS_AUTOARRANGE Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_ALIGNTOP Or LVS_OWNERDATA Or _
                      WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, WS_EX_CLIENTEDGE)

    Call GetClientRect(Me.hwnd, rc)
    SetWindowPos hLVVG, 0, 200, 0, rc.Right - 200, rc.Bottom, 0
      
    Dim lvsex As LVStylesEx
    lvsex = LVS_EX_DOUBLEBUFFER Or LVS_EX_FULLROWSELECT
    
    Call ListView_SetExtendedStyle(hLVVG, lvsex)
    Dim swt1 As String
    Dim swt2 As String
    swt1 = "explorer"
    swt2 = ""
    Call SetWindowTheme(hLVVG, StrPtr(swt1), 0&)
    
    Dim iCurViewMode As Long
    iCurViewMode = LV_VIEW_DETAILS
    Call SendMessage(hLVVG, LVM_SETVIEW, iCurViewMode, ByVal 0&)
    
    ReDim sColText(1)
    sColText(0) = "Index"
    sColText(1) = "Name"
    
    lvcol.mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    lvcol.fmt = LVCFMT_CENTER
    lvcol.cchTextMax = Len(sColText(0))
    lvcol.pszText = StrPtr(sColText(0))
    lvcol.cx = 70
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 1, lvcol)

    lvcol.cchTextMax = Len(sColText(1))
    lvcol.pszText = StrPtr(sColText(1))
    lvcol.cx = 140
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 2, lvcol)
End Sub

Private Property Get IControlEvent_Base() As MSForms.UserForm
    Set IControlEvent_Base = Me
End Property

Private Property Get IControlEvent_ControlEvents() As ControlEvents
    Set IControlEvent_ControlEvents = This.Controls
End Property

Private Property Get IControlEvent_PrevControl() As ControlEvent
    Set IControlEvent_PrevControl = This.PrevControl
End Property

Private Property Let IControlEvent_PrevControl(RHS As ControlEvent)
    Set This.PrevControl = RHS
End Property

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)

End Sub

Private Sub ListView1_AfterUpdate()

End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListView1_Click()

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

End Sub

Private Sub ListView1_DblClick()

End Sub

Private Sub ListView1_Enter()

End Sub

Private Sub ListView1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, ByVal Shift As Integer)

End Sub

Private Sub ListView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub ListView1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub ListView1_OLECompleteDrag(Effect As Long)

End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub ListView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

End Sub

Private Sub ListView1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

End Sub

Private Sub ListView1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)

End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

End Sub

Private Sub UserForm_Initialize()
    OnMouseOver = False
    Set This.Controls = New ControlEvents
    
    With This.Controls
        .ParentForm = Me
        .Init
    End With
'    Call FlatButtonInitialize
    Set TextEx = New TextBoxEx
    TextEx.Init "TestTextBox", Me, Me, 10, 10, 20, 80
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Set This.PrevControl = Nothing
End Sub

Private Sub UserForm_Terminate()
End Sub

'Private Sub FlatButtonInitialize()
'    Dim Ctrl As MSForms.Control
'    For Each Ctrl In Me.Controls
'        If InStr(1, Ctrl.Name, "FlatButton") > 0 Then
'            Call This.Controls.RegisterControl(Ctrl, "FlatButton;Hover") ', BaseStyle, HighlightStyle, ClickStyle)
'        End If
'    Next Ctrl
'End Sub
'

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "TextBoxName"
            searchCriteriaName = CtrlEvt.Control.Text
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxRegNoFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxRegNoTo"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxBtNoFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxBtNoTo"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxPathNumFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxPathNumTo"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxOperationDayFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxOperationDayTo"
            Call Filter3
    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case CtrlEvt.Control.Name = "TextBoxOperationDayEdit"
            If VBA.IsDate(CtrlEvt.Control.value) Then
'                searchCriteriaDate = CtrlEvt.Control.value
                CtrlEvt.Control.Text = Format(searchCriteriaDate, "YYYY/MM/DD")
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxRegNoFrom"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                 If Me.OptionButtonRegNoSingle Then
                    searchCriteriaRegNo = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaRegNo = CtrlEvt.Control.value & "," & TextBoxRegNoTo.value
                End If
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxRegNoTo"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                searchCriteriaRegNo = TextBoxRegNoFrom.value & "," & CtrlEvt.Control.value
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxNoByTargetFrom"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                 If Me.OptionButtonNoByTargetSingle Then
                    searchCriteriaNoByTarget = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaNoByTarget = CtrlEvt.Control.value & "," & TextBoxNoByTargetTo.value
                End If
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxNoByTargetTo"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                searchCriteriaNoByTarget = TextBoxNoByTargetFrom.value & "," & CtrlEvt.Control.value
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxPathNumFrom"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                 If Me.OptionButtonPathNumSingle Then
                    searchCriteriaPathNum = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaPathNum = CtrlEvt.Control.value & "," & TextBoxPathNumTo.value
                End If
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxPathNumTo"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                searchCriteriaPathNum = TextBoxPathNumFrom.value & "," & CtrlEvt.Control.value
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxOperationDayFrom"
            If VBA.IsDate(CtrlEvt.Control.value) Then
                If Me.OptionButtonSingleDay Then
                    searchCriteriaDate = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaDate = CtrlEvt.Control.value & "," & TextBoxOperationDayTo.value
                End If
                CtrlEvt.Control.Text = Format(CtrlEvt.Control.value, "YYYY/MM/DD")
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxOperationDayTo"
            If VBA.IsDate(CtrlEvt.Control.value) Then
                searchCriteriaDate = TextBoxOperationDayFrom.value & "," & CtrlEvt.Control.value
                CtrlEvt.Control.Text = Format(CtrlEvt.Control.value, "YYYY/MM/DD")
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)
    Dim I As Long
    Select Case True
'        Case CtrlEvt.Control.Name = "ComboBoxAddress"
'            Debug.Print CtrlEvt.Control.Name & " Change:" & CtrlEvt.Control.Text
'            searchCriteriaAddress = CtrlEvt.Control.Text
'            Call Filter3
'        Case Left(CtrlEvt.Control.Name, 14) = "CheckNoByTarget"
'            For i = 1 To 4
'                With Me.Controls("CheckNoByTarget" & i)
'                    If .value Then
'                        searchCriteriaNoByTarget = searchCriteriaNoByTarget & "," & .Caption
'                    End If
'                End With
'            Next i
'            searchCriteriaNoByTarget = Right(searchCriteriaNoByTarget, Len(searchCriteriaNoByTarget) - 1)

        Case Else
    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " OnClick"
    Select Case True
'        Case CtrlEvt.Attributes.Exists("DatePicker")
'            Debug.Print "OpenDatePicker"
'            Call OpenDatePicker(CtrlEvt)
'        Case CtrlEvt.Control.Tag = "SideBar"
'            Call OpenSideBar(CtrlEvt)
        Case Else
            Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Call IControlEvent_OnClick(CtrlEvt)
    DoEvents
    Cancel = True
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "ComboBoxAddress"
            Debug.Print onFocusComboBox
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            HookControl CtrlEvt

    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(CtrlEvt As ControlEvent)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(CtrlEvt As ControlEvent, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
    Select Case True
        Case CtrlEvt.Control.Name = "TextBoxOperationDayEdit"
            If KeyCode = 187 And Shift = 2 Then CtrlEvt.Control.value = Format(Now, "YYYY/MM/DD") ' Ctrl + 「;」

    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnKeyPress(CtrlEvt As ControlEvent, _
                                   ByVal KeyAscii As MSForms.IReturnInteger)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(CtrlEvt As ControlEvent, _
                                ByVal KeyCode As MSForms.IReturnInteger, _
                                ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(CtrlEvt As ControlEvent)
    Select Case True
        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonEditNoByTarget") > 0
            searchCriteriaNoByTarget = Replace(CtrlEvt.Control.Name, "OptionButtonEditNoByTarget", "")
            Call Filter3
        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonMode") > 0
            If CtrlEvt.Control.Name = "OptionButtonModeRegistorItem" Then
                Me.MultiPageSwitchMode.value = 0
            Else
                Me.MultiPageSwitchMode.value = 1
            End If
        Case Else
    End Select
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
'    Select Case True
'        Case CtrlEvt.Control.Name = "ListBoxEdit"
''            If Util.GetTimer > Delay + toolTipDelayTime Then
'                Call ShowListToolTip(CtrlEvt, X, Y)
''            End If
'    End Select
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseMove:(" & X & "," & Y & ") / Button:" & Button & " / Shift:" & Shift
End Sub

Private Sub IControlEvent_OnMouseUp(CtrlEvt As ControlEvent, _
                                  ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal x As Single, _
                                  ByVal y As Single)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseUp:"
End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal x As Single, _
                               ByVal y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(CtrlEvt As ControlEvent, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, _
                   ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Error:"
End Sub

Private Sub IControlEvent_OnAddControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " AddControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnLayout(CtrlEvt As ControlEvent)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Layout"
End Sub

Private Sub IControlEvent_OnRemoveControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " RemoveControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent)
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Scroll"
'End Sub

Private Sub IControlEvent_OnZoom(CtrlEvt As ControlEvent, Percent As Integer)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(CtrlEvt As ControlEvent)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(CtrlEvt As ControlEvent)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " SpinUp"
End Sub

Private Sub IControlEvent_OnMouseOver(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOver / Button(" & Button & ") / Shift(" & Shift & ")"
    Call MouseOver(CtrlEvt, Button, Shift, x, y)
    Set This.PrevControl = CtrlEvt
End Sub

Private Sub IControlEvent_OnMouseOut(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOut / Button(" & Button & ") / Shift(" & Shift & ")"
    Call MouseOut(CtrlEvt, Button, Shift, x, y)
End Sub

'-------------------------------------------------------------------------------
Private Sub MouseOver(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Static Delay As Double
'    Select Case True
'        Case CtrlEvt.Control.Name = "ComboBoxAddress"
'            If onFocusComboBox Then Exit Sub
'            onFocusComboBox = True
'            HookControl CtrlEvt '.Control
'        Case CtrlEvt.Control.Name = "ListBoxEdit"
''            Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseMove:TopIndex:" & CtrlEvt.Control.TopIndex & " / MousePointer:(" & X & "," & Y & ") / Button:" & Button & " / Shift:" & Shift
'            If onFocusListBox Then
'            Else
'                onFocusListBox = True
'                Delay = Util.GetTimer
'                HookControl CtrlEvt '.Control
'            End If
''            If Util.GetTimer > Delay + toolTipDelayTime Then
'                Call ShowListToolTip(CtrlEvt, X, Y)
''            End If
'
'        Case TypeName(CtrlEvt.Control) = "Frame"
'            If Left(CtrlEvt.Control.Tag, 13) = "SelectionField" Then
'                If onFocusFrame Then Exit Sub
'                If Me.FrameFilter.ScrollBars = fmScrollBarsNone Then UnHook: Exit Sub
'                onFocusFrame = True
'                HookControl CtrlEvt '.Control
'            End If
'        Case TypeName(CtrlEvt.Control) = "Label"
'            Select Case True
'                Case CtrlEvt.Control.Tag = "Button"
'                    CtrlEvt.Control.Object.BackStyle = fmBackStyleTransparent
'                Case CtrlEvt.Control.Tag = "SideBar"
'                    CtrlEvt.Control.Object.BackColor = MouseOverColor
'                Case Else
'            End Select
'        Case Else
'            UnHook
'            onFocusListBox = False
'            onFocusComboBox = False
'            onFocusFrame = False
'            Call CloseListToolTip
'            Delay = 0
'    End Select
    If OnMouseOver Then Exit Sub
    
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOver"
'    With Me.Frame1
'        .BorderColor = &HFFFFC0
'        .ForeColor = &HFFFFC0
'    End With
    If CtrlEvt.Attributes.Exists("TextBoxEx") Then
        If TypeName(CtrlEvt.Control) = "TextBox" Then
            Call TextEx.MouseOver(CtrlEvt, Button, Shift, x, y)
        End If
    End If
    OnMouseOver = True
End Sub

Private Sub MouseOut(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Select Case True
'        Case This.PrevControl.Control.Tag = "Button"
'            This.PrevControl.Control.Object.BackStyle = fmBackStyleOpaque
'        Case This.PrevControl.Control.Tag = "SideBar"
'            This.PrevControl.Control.Object.BackColor = FrameBaseColor
'        Case This.PrevControl.Control.Name = "ListBoxEdit"
'            UnHook
'            onFocusListBox = False
'            Call CloseListToolTip
'    End Select
    
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOut"
'    With Me.Image5
'        .BackColor = &HE0E0E0
''        .ForeColor = &HFFFFC0
'    End With

    If CtrlEvt.Attributes.Exists("TextBoxEx") Then
        If TypeName(CtrlEvt.Control) = "TextBox" Then
            Call TextEx.MouseOut(CtrlEvt, Button, Shift, x, y)
        End If
    End If
    OnMouseOver = False
End Sub
