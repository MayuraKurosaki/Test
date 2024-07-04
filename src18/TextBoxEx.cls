VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IControlEvent

Public WithEvents Form As MSForms.UserForm
Attribute Form.VB_VarHelpID = -1
Public WithEvents Frame As MSForms.Frame
Attribute Frame.VB_VarHelpID = -1
Public WithEvents MultiPage As MSForms.MultiPage
Attribute MultiPage.VB_VarHelpID = -1
Public WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1
Public WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents Label As MSForms.Label
Attribute Label.VB_VarHelpID = -1

Private Type Field
    Controls As ControlEvents
'    PrevControl As MSForms.IControl
    Parent As MSForms.Control
    PrevControl As ControlEvent
End Type

Private This As Field

Private Sub Class_Initialize()
    Set This.Controls = New ControlEvents
End Sub

Private Sub Class_Terminate()
    Set This.Controls = Nothing
End Sub

Private Property Get IControlEvent_PrevControl() As ControlEvent
    Set IControlEvent_PrevControl = This.PrevControl
End Property

Private Property Let IControlEvent_PrevControl(RHS As ControlEvent)
    Set This.PrevControl = RHS
End Property

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, _
                                         ByVal Cancel As MSForms.IReturnBoolean)
'    Debug.Print CtrlEvt.Control.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)
    Call ClickProcedure(CtrlEvt)
'    Debug.Print CtrlEvt.Control.Name & " Click"
End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, _
                                     ByVal Cancel As MSForms.IReturnBoolean)
    Select Case CtrlEvt.Control.Name
        Case "FlatButtonPagePrev", "FlatButtonPageNext"
            Call IControlEvent_OnClick(CtrlEvt)
            DoEvents
            Cancel = True
    End Select
'    Debug.Print CtrlEvt.Control.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(CtrlEvt As ControlEvent, _
                                 ByVal Cancel As MSForms.IReturnBoolean)
'    Debug.Print CtrlEvt.Control.Name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(CtrlEvt As ControlEvent, _
                                    ByVal KeyCode As MSForms.IReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call AllMouseOut
        Me.Hide
    End If
'    Debug.Print CtrlEvt.Control.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnKeyPress(CtrlEvt As ControlEvent, _
                                     ByVal KeyAscii As MSForms.IReturnInteger)
'    Debug.Print CtrlEvt.Control.Name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(CtrlEvt As ControlEvent, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
'    Debug.Print CtrlEvt.Control.Name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseDown:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseMove:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnMouseUp(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseUp:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, _
                                           ByVal Cancel As MSForms.ReturnBoolean, _
                                           ByVal Data As MSForms.DataObject, _
                                           ByVal X As Single, _
                                           ByVal Y As Single, _
                                           ByVal DragState As MSForms.fmDragState, _
                                           ByVal Effect As MSForms.ReturnEffect, _
                                           ByVal Shift As Integer)
'    Debug.Print CtrlEvt.Control.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                                              ByVal Cancel As MSForms.ReturnBoolean, _
                                              ByVal Action As MSForms.fmAction, _
                                              ByVal Data As MSForms.DataObject, _
                                              ByVal X As Single, _
                                              ByVal Y As Single, _
                                              ByVal Effect As MSForms.ReturnEffect, _
                                              ByVal Shift As Integer)
'    Debug.Print CtrlEvt.Control.Name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(CtrlEvt As ControlEvent, _
                                  ByVal Number As Integer, _
                                  ByVal Description As MSForms.ReturnString, _
                                  ByVal SCode As Long, _
                                  ByVal Source As String, _
                                  ByVal HelpFile As String, _
                                  ByVal HelpContext As Long, _
                                  ByVal CancelDisplay As MSForms.ReturnBoolean)
'    Debug.Print CtrlEvt.Control.Name & " Error:"
End Sub

Private Sub IControlEvent_OnAddControl(CtrlEvt As ControlEvent, _
                                       ByVal Control As MSForms.Control)
'    Debug.Print CtrlEvt.Control.Name & " AddControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnLayout(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Layout"
End Sub

Private Sub IControlEvent_OnRemoveControl(CtrlEvt As ControlEvent, _
                                          ByVal Control As MSForms.Control)
'    Debug.Print CtrlEvt.Control.Name & " RemoveControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent, _
                                   ByVal ActionX As MSForms.fmScrollAction, _
                                   ByVal ActionY As MSForms.fmScrollAction, _
                                   ByVal RequestDx As Single, _
                                   ByVal RequestDy As Single, _
                                   ByVal ActualDx As MSForms.ReturnSingle, _
                                   ByVal ActualDy As MSForms.ReturnSingle)
'    Debug.Print CtrlEvt.Control.Name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Scroll"
'End Sub

Private Sub IControlEvent_OnZoom(CtrlEvt As ControlEvent, _
                                 Percent As Integer)
'    Debug.Print CtrlEvt.Control.Name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " SpinUp"
End Sub

Private Sub IControlEvent_OnMouseOver(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseOver:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
    Call MouseOver(CtrlEvt, Button, Shift, X, Y)
    Set This.PrevControl = CtrlEvt
End Sub

Private Sub IControlEvent_OnMouseOut(CtrlEvt As ControlEvent, _
                                     ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseOut:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
    Call MouseOut(CtrlEvt, Button, Shift, X, Y)
End Sub

'-------------------------------------------------------------
Public Sub Init(Parent As Object, ByVal Top As Single, ByVal Left As Single, ByVal Height As Single, ByVal Width As Single, Optional ByVal Combo As Boolean = False)
    Select Case TypeName(Parent)
        Case "UserForm"
            This.Controls.ParentForm = Parent
            Set This.Parent = Parent
        Case "Frame"
        Case "MultiPage"
        Case "TabStrip"
    End Select
    
    This.Controls.ParentForm = Form
    Set This.Parent = Parent
    
    
    
'    Call RegisterControls
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        Select Case True
            Case VBA.Left$(ctrl.Name, 10) = "FlatButton"
                Call This.Controls.RegisterControl(ctrl, "FlatButton")
            Case VBA.Left$(ctrl.Name, 6) = "Picker"
                Call This.Controls.RegisterControl(ctrl, "Picker")
            Case Else
                Call This.Controls.RegisterControl(ctrl)
        End Select
    Next ctrl
    
    Me.Height = 190
    Me.Width = 212
    
    This.Today = VBA.Fix(Now)
    
    This.YearMin = YearMin
    This.YearMax = YearMax
    
    Call MakeHolidayDictionary(YearMin, YearMax, SheetList.ListObjects("T_月日固定休日"), SheetList.ListObjects("T_月週曜日固定休日"))
End Sub

Private Sub MouseOver(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With CtrlEvt
        Select Case True
            Case .Attributes.Exists("Picker")
                Me.Controls(VBA.Replace$(.Control.Name, "Picker", "") & "BG").BackColor = &H808080
            Case .Control.Name = "FlatButtonPagePrev"
                .Control.Object.Picture = ResourceFlatButtonPagePrevHover.Picture
            Case .Control.Name = "FlatButtonPageNext"
                .Control.Object.Picture = ResourceFlatButtonPageNextHover.Picture
            Case .Control.Name = "FlatButtonClose"
                .Control.Object.Picture = ResourceFlatButtonCloseHover.Picture
            Case .Control.Name = "FlatButtonSelectToday"
                .Control.Object.Picture = ResourceFlatButtonSelectTodayHover.Picture
            Case .Control.Name = "FlatButtonBackFromYear"
                .Control.Object.Picture = ResourceFlatButtonBackHover.Picture
            Case .Control.Name = "FlatButtonBackFromMonth"
                .Control.Object.Picture = ResourceFlatButtonBackHover.Picture
    '        Case .Control.Name = "LabelPeriod"
    '            Me.Controls("LabelPeriodBG").Picture = ResourceLabelPeriodHover.Picture
            Case .Control.Name = "FlatButtonSelectYear"
                Me.Controls("SelectYearBG").Picture = ResourceFlatButtonSelectYearHover.Picture
            Case .Control.Name = "FlatButtonSelectMonth"
                Me.Controls("SelectMonthBG").Picture = ResourceFlatButtonSelectMonthHover.Picture
            Case Else
        End Select
    End With
End Sub

Private Sub MouseOut(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With This.PrevControl
        Select Case True
            Case .Attributes.Exists("Picker")
                Me.Controls(VBA.Replace$(.Control.Name, "Picker", "") & "BG").BackColor = &HFFFFFF
            Case .Control.Name = "FlatButtonPagePrev"
                .Control.Object.Picture = ResourceFlatButtonPagePrev.Picture
            Case .Control.Name = "FlatButtonPageNext"
                .Control.Object.Picture = ResourceFlatButtonPageNext.Picture
            Case .Control.Name = "FlatButtonClose"
                .Control.Object.Picture = ResourceFlatButtonClose.Picture
            Case .Control.Name = "FlatButtonSelectToday"
                .Control.Object.Picture = ResourceFlatButtonSelectToday.Picture
            Case .Control.Name = "FlatButtonBackFromYear"
                .Control.Object.Picture = ResourceFlatButtonBack.Picture
            Case .Control.Name = "FlatButtonBackFromMonth"
                .Control.Object.Picture = ResourceFlatButtonBack.Picture
    '        Case .Control.Name = "LabelPeriod"
    '            Me.Controls("LabelPeriodBG").Picture = ResourceLabelPeriod.Picture
            Case .Control.Name = "FlatButtonSelectYear"
                Me.Controls("SelectYearBG").Picture = ResourceFlatButtonSelectYear.Picture
            Case .Control.Name = "FlatButtonSelectMonth"
                Me.Controls("SelectMonthBG").Picture = ResourceFlatButtonSelectMonth.Picture
            Case Else
        End Select
    End With
End Sub

Private Sub AllMouseOut()
    If Not This.PrevControl Is Nothing Then
        Call IControlEvent_OnMouseOut(This.PrevControl, 0, 0, 0, 0)
        Set This.PrevControl = Nothing
    End If
End Sub

Private Sub ClickProcedure(CtrlEvt As ControlEvent)
    Select Case This.Mode
        Case dpNormal
            Select Case True
                Case CtrlEvt.Control.Name = "FlatButtonPagePrev"
                    If This.Month = 1 Then
                        If This.Year - 1 < This.YearMin Then
                            Exit Sub
                        Else
                            This.Year = This.Year - 1
                            This.PeriodStart = (This.Year \ 20) * 20
                            This.Month = 12
                        End If
                    Else
                        This.Month = This.Month - 1
                    End If
                    PopulateDatePicker This.Year, This.Month
                Case CtrlEvt.Control.Name = "FlatButtonPageNext"
                    If This.Month = 12 Then
                        If This.Year + 1 > This.YearMax Then
                            Exit Sub
                        Else
                            This.Year = This.Year + 1
                            This.PeriodStart = (This.Year \ 20) * 20
                            This.Month = 1
                        End If
                    Else
                        This.Month = This.Month + 1
                    End If
                    PopulateDatePicker This.Year, This.Month
                Case CtrlEvt.Attributes.Exists("Picker")
                    With This
                        If VBA.Replace(CtrlEvt.Control.Name, "PickerDay", "") < .StartIndex Then
                            If .Month = 1 Then
                                .Year = .Year - 1
                                .PeriodStart = (.Year \ 20) * 20
                                .Month = 12
                            Else
                                .Month = .Month - 1
                            End If
                        ElseIf VBA.Replace(CtrlEvt.Control.Name, "PickerDay", "") > .EndIndex Then
                            If .Month = 12 Then
                                .Year = .Year + 1
                                .PeriodStart = (.Year \ 20) * 20
                                .Month = 1
                            Else
                                .Month = .Month + 1
                            End If
                        End If
                        .Day = VBA.CInt(CtrlEvt.Control.Caption)
                        .CurrentDate = VBA.DateSerial(.Year, .Month, .Day)
                    End With
                    Call SetDateToTextBox
                    Call AllMouseOut
                    Me.Hide
                Case CtrlEvt.Control.Name = "FlatButtonSelectToday"
                    This.CurrentDate = This.Today
                    Call SetDateToTextBox
                    Call AllMouseOut
                    Me.Hide
            End Select
        Case dpYear
            Select Case True
                Case CtrlEvt.Control.Name = "FlatButtonPagePrev"
                    If This.PeriodStart - 1 < This.YearMin Then Exit Sub
                    This.PeriodStart = This.PeriodStart - 20
                    PopulateYearPicker This.PeriodStart
                Case CtrlEvt.Control.Name = "FlatButtonPageNext"
                    If This.PeriodStart + 1 > This.YearMax Then Exit Sub
                    This.PeriodStart = This.PeriodStart + 20
                    PopulateYearPicker This.PeriodStart
                Case CtrlEvt.Attributes.Exists("Picker")
                    With This
                        .Year = VBA.CInt(CtrlEvt.Control.Caption)
                        .PeriodStart = (.Year \ 20) * 20
                        PopulateDatePicker .Year, .Month
                        SetPickerMode dpNormal
                    End With
                Case CtrlEvt.Control.Name = "FlatButtonBackFromYear"
                    SetPickerMode dpNormal
            End Select
        Case dpMonth
            Select Case True
                Case CtrlEvt.Control.Name = "FlatButtonPagePrev"
                    If This.Year - 1 < This.YearMin Then Exit Sub
                    This.Year = This.Year - 1
                    This.PeriodStart = (This.Year \ 20) * 20
                    PopulateMonthPicker This.Year
                Case CtrlEvt.Control.Name = "FlatButtonPageNext"
                    If This.Year + 1 > This.YearMax Then Exit Sub
                    This.Year = This.Year + 1
                    This.PeriodStart = (This.Year \ 20) * 20
                    PopulateMonthPicker This.Year
                Case CtrlEvt.Attributes.Exists("Picker")
                    With This
                        .Month = VBA.CInt(CtrlEvt.Control.Caption)
                        PopulateDatePicker .Year, .Month
                        SetPickerMode dpNormal
                    End With
                Case CtrlEvt.Control.Name = "FlatButtonSelectMonth"
                    SetPickerMode dpMonth
            End Select
    End Select
    
    Select Case True
        Case CtrlEvt.Control.Name = "FlatButtonSelectYear"
            SetPickerMode dpYear
        Case CtrlEvt.Control.Name = "FlatButtonSelectMonth"
            SetPickerMode dpMonth
        Case CtrlEvt.Control.Name = "FlatButtonClose"
            Call AllMouseOut
            Me.Hide
        Case Else
    End Select
End Sub

Private Sub RegisterControls()
'    Dim ctrl As MSForms.Control
'    For Each ctrl In Me.Controls
'        Select Case True
'            Case VBA.Left$(ctrl.Name, 10) = "FlatButton"
'                Call This.Controls.RegisterControl(ctrl, "FlatButton")
'            Case VBA.Left$(ctrl.Name, 6) = "Picker"
'                Call This.Controls.RegisterControl(ctrl, "Picker")
'            Case Else
'                Call This.Controls.RegisterControl(ctrl)
'        End Select
'    Next ctrl
End Sub
