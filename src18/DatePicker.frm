VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "UserForm1"
   ClientHeight    =   8240.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4260
   OleObjectBlob   =   "DatePicker.frx":0000
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IControlEvent

Private Enum PickerMode
    dpNormal = 0
    dpYear = 1
    dpMonth = 2
End Enum

Private Type Field
    Controls As ControlEvents
    PrevControl As ControlEvent
    Mode As PickerMode '0=normal, 1=months, 2=years
    Today As Date
    Year As Integer
    Month As Integer
    Day As Integer
    YearMin As Long
    YearMax As Long
    PeriodStart As Integer
    StartOfMonthDay As Integer
    StartIndex As Long
    EndIndex As Long
    CurrentDate As Date
    LinkTextBox As MSForms.TextBox
End Type

Private This As Field

Private Property Get IControlEvent_PrevControl() As ControlEvent
    Set IControlEvent_PrevControl = This.PrevControl
End Property

Private Property Let IControlEvent_PrevControl(RHS As ControlEvent)
    Set This.PrevControl = RHS
End Property

Private Sub UserForm_Initialize()
    FormNonCaption Me, True
    Set This.Controls = New ControlEvents
    This.Controls.ParentForm = Me

    Call RegisterControls
    Me.Height = 190
    Me.Width = 212
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "UserForm_KeyDown"
    If KeyCode = vbKeyEscape Then
        Call AllMouseOut
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
'    Select Case True
'        Case TypeName(This.LinkTextBox) = "TextBox"
'            This.LinkTextBox.value = VBA.Fix(This.CurrentDate)
'    End Select
'    Set dicHoliday_ = Nothing
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call AllMouseOut
End Sub

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
Public Sub Init(Optional YearMin As Long = 2000, Optional YearMax As Long = 2050)
    This.Today = VBA.Fix(Now)
    
    This.YearMin = YearMin
    This.YearMax = YearMax
    
    Call MakeHolidayDictionary(YearMin, YearMax, SheetList.ListObjects("T_月日固定休日"), SheetList.ListObjects("T_月週曜日固定休日"))
End Sub

Public Sub ShowPicker(TextBox As MSForms.TextBox)
    Set This.LinkTextBox = TextBox
    This.Today = VBA.Fix(Now)
    Select Case True
        Case TypeName(This.LinkTextBox) = "TextBox"
            This.CurrentDate = CDate(IIf(This.LinkTextBox.Text = "", This.Today, This.LinkTextBox.Text))
    End Select
    
    Dim DispSizePixel As POINTAPI
    DispSizePixel = GetDisplaySize
    Dim DispSizePoint As POINTF
    DispSizePoint = PixcelToPoint(DispSizePixel)
    
    Dim pos As POINTF
    pos = GetControlPosition(TextBox, BottomLeft)
    If pos.X + Me.Width > DispSizePoint.X Then pos.X = pos.X - (Me.Width - TextBox.Width)
    If pos.X < 0 Then pos.X = 0
    If pos.Y + Me.Height > DispSizePoint.Y Then pos.Y = pos.Y - Me.Height - TextBox.Height
    
    Me.Top = pos.Y
    Me.Left = pos.X

    Call PopulateDatePicker(VBA.Year(This.CurrentDate), VBA.Month(This.CurrentDate))
    SetPickerMode dpNormal

    Me.Show
End Sub

Private Sub SetDateToTextBox()
    Select Case True
        Case TypeName(This.LinkTextBox) = "TextBox"
            This.LinkTextBox.value = VBA.Fix(This.CurrentDate)
    End Select
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

Private Sub PopulateDatePicker(Optional Year As Integer = 0, Optional Month As Integer = 0)
    If Year = 0 Or Month = 0 Then
        Year = VBA.Year(This.Today)
        Month = VBA.Month(This.Today)
    End If
    
    With This
        .Year = Year
        .Month = Month
    End With
    
    Me.Controls("LabelPeriod").Visible = False
    Me.Controls("LabelPeriodBG").Visible = False
    Me.Controls("FlatButtonSelectYear").Caption = Year & "年"
    Me.Controls("FlatButtonSelectYear").Visible = True
    Me.Controls("SelectYearBG").Visible = True
    Me.Controls("FlatButtonSelectMonth").Caption = VBA.MonthName(Month, False)
    Me.Controls("FlatButtonSelectMonth").Visible = True
    Me.Controls("SelectMonthBG").Visible = True
    
    Dim startOfMonth As Date
    Dim trackingDate As Date
    startOfMonth = VBA.DateSerial(Year, Month, 1)
    This.StartOfMonthDay = VBA.Weekday(startOfMonth, vbSunday)
    trackingDate = DateAdd("d", -This.StartOfMonthDay + 1, startOfMonth)
    
    Dim captionDay As Integer: captionDay = 0
    Dim labelDay As Control
    Dim i As Long
    Dim HolidayName As String
    For i = 1 To 42
        Set labelDay = Me.Controls("PickerDay" & i)
        captionDay = VBA.Day(trackingDate)
        labelDay.Caption = captionDay
        If This.StartIndex = 0 And captionDay = 1 Then This.StartIndex = i
        If This.EndIndex = 0 And This.StartIndex <> 0 And VBA.Month(trackingDate) <> This.Month Then This.EndIndex = i
        labelDay.Enabled = True
        labelDay.ControlTipText = ""
        Select Case Weekday(trackingDate, vbSunday)
            Case vbSaturday: labelDay.ForeColor = rgbRoyalBlue
            Case vbSunday: labelDay.ForeColor = rgbLightCoral
            Case Else: labelDay.ForeColor = ColorConstants.vbBlack
        End Select
        
        If trackingDate = This.Today Then
            labelDay.BackColor = &HFFFFC0
        Else
            labelDay.BackColor = &HFFFFFF
        End If

        If VBA.Year(trackingDate) < This.YearMin Or VBA.Year(trackingDate) > This.YearMax Then
            Debug.Print trackingDate
            labelDay.Enabled = False
        Else
            If IsHoliday(trackingDate, HolidayName) Then
                labelDay.ForeColor = rgbLightCoral
                labelDay.ControlTipText = HolidayName
            End If
        End If
        
        If VBA.Month(trackingDate) <> Month Then
            labelDay.ForeColor = rgbGray
        End If
        
        trackingDate = VBA.DateAdd("d", 1, trackingDate)
    Next i
End Sub

Private Sub PopulateMonthPicker(Optional Year As Integer = 0)
    If Year = 0 Then Year = VBA.Year(This.Today)
    
    Me.Controls("LabelPeriod").Visible = False
    Me.Controls("LabelPeriodBG").Visible = False
    Me.Controls("FlatButtonSelectYear").Caption = Year & "年"
    Me.Controls("FlatButtonSelectYear").Visible = True
    Me.Controls("FlatButtonSelectMonth").Visible = False
    Me.Controls("SelectYearBG").Visible = True
    Me.Controls("SelectMonthBG").Visible = False
    
    Dim labelMonth As Control
    Dim i As Long
    For i = 1 To 12
        Set labelMonth = Me.Controls("PickerMonth" & i)
        If This.Year = VBA.Year(This.Today) And i = VBA.Month(This.Today) Then
            labelMonth.BackColor = &HFFFFC0
        Else
            labelMonth.BackColor = &HFFFFFF
        End If
    Next i
End Sub

Private Sub PopulateYearPicker(Optional PeriodStartYear As Integer = 0)
    If PeriodStartYear = 0 Then PeriodStartYear = (This.Year \ 20) * 20
    
    Me.Controls("FlatButtonSelectYear").Visible = False
    Me.Controls("FlatButtonSelectMonth").Visible = False
    Me.Controls("SelectYearBG").Visible = False
    Me.Controls("SelectMonthBG").Visible = False
    
    Dim labelYear As Control
    Dim loopStart As Integer
    
    loopStart = PeriodStartYear
    Dim captionStart As Integer, captionEnd As Integer
    If PeriodStartYear < This.YearMin Then
        captionStart = This.YearMin
    Else
        captionStart = PeriodStartYear
    End If
    If PeriodStartYear + 19 > This.YearMax Then
        captionEnd = This.YearMax
    Else
        captionEnd = PeriodStartYear + 19
    End If
    Me.Controls("LabelPeriod").Caption = captionStart & "年-" & captionEnd & "年"
    Me.Controls("LabelPeriod").Visible = True
    Me.Controls("LabelPeriodBG").Visible = True
    
    Dim i As Long
    For i = 1 To 20
        Set labelYear = Me.Controls("PickerYear" & i)
        
        labelYear.Caption = loopStart
               
        If loopStart < This.YearMin Or loopStart > This.YearMax Then
            labelYear.Enabled = False
        Else
            labelYear.Enabled = True
            If loopStart = VBA.Year(This.Today) Then
                labelYear.BackColor = &HFFFFC0
            Else
                labelYear.BackColor = &HFFFFFF
            End If
        End If
        
        loopStart = loopStart + 1
    Next i
End Sub

Private Sub SetPickerMode(Mode As PickerMode)
    Select Case Mode
        Case 0
            PopulateDatePicker This.Year, This.Month
        Case 1
            PopulateYearPicker This.PeriodStart
        Case 2
            PopulateMonthPicker This.Year
    End Select
    This.Mode = Mode
    Me.MultiPageSelectPicker.value = Mode
End Sub

Private Sub RegisterControls()
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
End Sub
