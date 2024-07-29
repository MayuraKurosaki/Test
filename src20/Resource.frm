VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Resource 
   Caption         =   "Resource"
   ClientHeight    =   7880
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   12600
   OleObjectBlob   =   "Resource.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IControlEvent

Private OnMouseOver As Boolean
Private OnFocus As Boolean
'Private FrameTop As Single

Private TextEx As TextBoxEx

Private Type Field
    Controls As ControlEvents
    PrevControl As ControlEvent
End Type

Private This As Field

Public Sub AddControl(bstrProgID As String, Optional Name As String = "", Optional Visible As Boolean = True)
    Call Me.Controls.Add(bstrProgID, Name, Visible)
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

Private Sub UserForm_Initialize()
'    FrameTop = Me.Frame1.Top
    OnMouseOver = False
    Set This.Controls = New ControlEvents
'    This.Controls.ParentForm = Me
    
'    Set TextEx = New TextBoxEx
'    TextEx.Init "TestTextBox", Me, Me, 300, 300, 30, 60
    
    
    With This.Controls
        .ParentForm = Me
        .Init
    End With
    Call FlatButtonInitialize
'    Call Util.MakeTransparentFrame(Frame1)
    Set TextEx = New TextBoxEx
    TextEx.Init "TestTextBox", Me, Me, 300, 300, 30, 60
    
    DatePicker.Init
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    If Not This.PrevControl Is Nothing Then
''        This.PrevControl.Object.BorderStyle = fmBorderStyleNone
'        Select Case This.PrevControl.Control.Tag
'            Case "Button"
'                Debug.Print This.PrevControl.Control.Name & "MouseOut"
'                This.PrevControl.Control.Object.BackStyle = fmBackStyleOpaque
'                This.PrevControl.Control.Object.BorderStyle = fmBorderStyleNone
'            Case Else
'        End Select
'    End If
'    Me.LabelEditDatePicker.BackStyle = fmBackStyleTransparent
'    Me.LabelNewDatePicker.BackStyle = fmBackStyleTransparent
'    UnHook
'    onFocusListBox = False
'    onFocusComboBox = False
'    onFocusFrame = False
    Set This.PrevControl = Nothing
End Sub

Private Sub UserForm_Terminate()
    Unload DatePicker
End Sub

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
    Dim i As Long
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
        Case CtrlEvt.Attributes.Exists("DatePicker")
            Debug.Print "OpenDatePicker"
            Call OpenDatePicker(CtrlEvt)
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
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
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
                                  ByVal X As Single, _
                                  ByVal Y As Single)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseUp:"
End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
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
                     ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOver / Button(" & Button & ") / Shift(" & Shift & ")"
    Call MouseOver(CtrlEvt, Button, Shift, X, Y)
    Set This.PrevControl = CtrlEvt
End Sub

Private Sub IControlEvent_OnMouseOut(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Debug.Print TypeName(Me) & ":" & CtrlEvt.Control.Name & " MouseOut / Button(" & Button & ") / Shift(" & Shift & ")"
    Call MouseOut(CtrlEvt, Button, Shift, X, Y)
End Sub

'-------------------------------------------------------------------------------
Private Sub MouseOver(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
    OnMouseOver = True
End Sub

Private Sub MouseOut(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
    OnMouseOver = False
End Sub

Private Sub FlatButtonInitialize()
    Dim Ctrl As MSForms.Control
    For Each Ctrl In Me.Controls
        If InStr(1, Ctrl.Name, "FlatButton") > 0 Then
            Call This.Controls.RegisterControl(Ctrl, "FlatButton;Hover") ', BaseStyle, HighlightStyle, ClickStyle)
            If InStr(1, Ctrl.Name, "DatePicker") > 0 Then
                This.Controls(Ctrl.Name).AttributeItems = "DatePicker"
            End If
        End If
    Next Ctrl
End Sub

Private Sub OpenDatePicker(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "FlatButtonOperationDatePickerEdit"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.ShowPicker Me.TextBoxOperationDayEdit
'            DatePicker.Init Me.TextBoxOperationDayEdit
        Case CtrlEvt.Control.Name = "FlatButtonCreationDatePickerEdit"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.ShowPicker Me.TextBoxCreationDayEdit
'            DatePicker.Init Me.TextBoxCreationDayEdit
        Case CtrlEvt.Control.Name = "LabelDatePickerFrom"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.ShowPicker Me.TextBoxOperationDayEdit
        Case CtrlEvt.Control.Name = "LabelDatePickerTo"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.ShowPicker Me.TextBoxOperationDayEdit
        Case CtrlEvt.Control.Name = "CommandButton1"
            DatePicker.ShowPicker Me.TextBox1
            
        Case Else
    End Select
End Sub





Private Sub CommandButton1_Click()
'    DatePicker.Init Me.TextBox1
    DatePicker.ShowPicker Me.TextBox1
End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub Image5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image_MouseOver
End Sub

Private Sub Image_MouseOver()
    Debug.Print OnMouseOver
    If OnMouseOver Then Exit Sub
    
    Debug.Print "Image_MouseOver"
    With Me.Image5
        .BackColor = &H808080
'        .ForeColor = &HFFFFC0
    End With
    OnMouseOver = True
End Sub

Private Sub Image_MouseOut()
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    Debug.Print "Image_MouseOut"
    With Me.Image5
        .BackColor = &HE0E0E0
'        .ForeColor = &HFFFFC0
    End With
    OnMouseOver = False
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub TextBox2_Enter()
    ExTextBox_GotFocus
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus
End Sub

Private Sub TextBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub ExTextBox_GotFocus()
    If OnFocus Then Exit Sub
    
    With Me.Frame1
        .Caption = Me.Label2.Caption
        .Top = FrameTop - 6
        FrameTop = .Top
        .Height = 30
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
    Me.Label2.Visible = False
    OnFocus = True
End Sub

Private Sub ExTextBox_LostFocus()
    If Not OnFocus Then Exit Sub
    
    With Me.Frame1
        .Caption = ""
        .Top = FrameTop + 6
        FrameTop = .Top
        .Height = 24
        .BorderColor = &H808080
    End With
    Me.Label2.Visible = True
    OnFocus = False
End Sub

Private Sub ExTextBox_MouseOver()
    If OnMouseOver Then Exit Sub
    
    With Me.Frame1
'        .Caption = Me.Label2.Caption
'        .Top = FrameTop - 6
'        FrameTop = .Top
'        .Height = 36
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
'    Me.Label2.Visible = False
    OnMouseOver = True
End Sub

Private Sub ExTextBox_MouseOut()
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    With Me.Frame1
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 30
        .BorderColor = &H808080
    End With
    Me.Label2.Visible = True
    OnMouseOver = False
End Sub

Private Sub TextBoxBody_Enter()
    ExTextBox_GotFocus2
End Sub

Private Sub TextBoxBody_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus2
End Sub

Private Sub TextBoxBody_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver2
End Sub

Private Sub LabelTextBoxCaption_Click()
    ExTextBox_GotFocus2
End Sub

Private Sub LabelTextBoxCaption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver2
End Sub

Private Sub ExTextBox_GotFocus2()
    If OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
    Me.LabelTextBoxCaption.Top = Me.LabelTextBoxFrame.Top - 8
    Me.LabelTextBoxCaption.Left = Me.LabelTextBoxFrame.Left + 6
    Me.LabelTextBoxCaption.FontSize = 8
    Me.LabelTextBoxCaption.ForeColor = &HFFFFC0
    Me.LabelTextBoxCaption.BackStyle = fmBackStyleTransparent
    Me.LabelTextBoxMask.Left = Me.LabelTextBoxCaption.Left
    With Me.LabelTextBoxCaption
        Me.LabelTextBoxMask.Width = MeasureTextSize(.Caption, .FontName, .FontSize).X * 1.3
    End With
'    Me.LabelTextBoxMask.Width = Me.LabelTextBoxCaption.Width
    Me.LabelTextBoxMask.Visible = True
    OnFocus = True
End Sub

Private Sub ExTextBox_LostFocus2()
    If Not OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 26
        .BorderColor = &H808080
        .ForeColor = &HC0C0C0
    End With
'    Me.LabelTextBoxCaption.Visible = True
    Me.LabelTextBoxCaption.Top = Me.LabelTextBoxFrame.Top + 3
    Me.LabelTextBoxCaption.Left = Me.LabelTextBoxFrame.Left + 1
    Me.LabelTextBoxCaption.FontSize = 10
    Me.LabelTextBoxCaption.ForeColor = &HC0C0C0
    Me.LabelTextBoxCaption.BackStyle = fmBackStyleTransparent
    Me.LabelTextBoxMask.Visible = False
    OnFocus = False
End Sub

Private Sub ExTextBox_MouseOver2()
    If OnMouseOver Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = Me.Label2.Caption
'        .Top = FrameTop - 6
'        FrameTop = .Top
'        .Height = 36
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
'    Me.LabelTextBoxCaption.Visible = False
    OnMouseOver = True
End Sub

Private Sub ExTextBox_MouseOut2()
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 30
        .BorderColor = &H808080
        .ForeColor = &HFFFFC0
    End With
'    Me.LabelTextBoxCaption.Visible = True
    OnMouseOver = False
End Sub

