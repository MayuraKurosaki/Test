VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePickerForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4160
   OleObjectBlob   =   "DatePickerForm.frx":0000
End
Attribute VB_Name = "DatePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IControlEvent

Private Enum PickerMode
    dpNormal = 0
    dpMonth = 1
    dpYear = 2
End Enum

Private Type TState
    Control As ControlEvents
    PrevControl As MSForms.IControl
    Mode As PickerMode '0=normal, 1=months, 2=years
    Year As Integer
    Month As Integer
    Day As Integer
    CurrentDate As Date
End Type

Private this As TState

Private Sub FrameDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub FrameMonth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub FrameYear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub MultiPageSelectPicker_Change()
    If Me.MultiPageSelectPicker.value = 1 Then populateYearPicker VBA.Year(this.CurrentDate)
End Sub

Private Sub MultiPageSelectPicker_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Private Property Let Mode(ByVal RHS As Integer)
'    This.Mode = RHS
'End Property
'
'Private Property Get Mode() As Integer
'    Mode = This.Mode
'End Property
'
'Private Property Let Year(ByVal RHS As Integer)
'    This.Year = RHS
'End Property
'
'Private Property Get Year() As Integer
'    Year = This.Year
'End Property
'
'Private Property Let Month(ByVal RHS As Integer)
'    This.Month = RHS
'End Property
'
'Private Property Get Month() As Integer
'    Month = This.Month
'End Property
'
'Private Property Let Day(ByVal RHS As Integer)
'    This.Day = RHS
'End Property
'
'Private Property Get Day() As Integer
'    Day = This.Day
'End Property
'
'Private Property Let CurrentDate(ByVal RHS As Date)
'    This.CurrentDate = RHS
'End Property
'
'Private Property Get CurrentDate() As Date
'    CurrentDate = This.CurrentDate
'End Property

Private Sub UserForm_Initialize()
    Call populateDatePicker
    setPickerMode dpNormal
    Set this.Control = New ControlEvents
    With this.Control
        .parent = Me
        .Init
    End With
    Me.Height = 220
    FormNonCaption Me, True
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
End Sub

Private Sub UserForm_Terminate()
End Sub

'populates the date picker for the month / year
Private Sub populateDatePicker(Optional Year As Integer, Optional Month As Integer)
    'if the year or month isn't passed in, then assume current year and month
    If Year = 0 Or Month = 0 Then
        Year = VBA.Year(Now())
        Month = VBA.Month(Now())
    End If
    
    With this
        .Year = Year
        .Month = Month
    End With
    
    'set the month and year in the calendar
    Me.Controls("LabelSelectYear").Caption = Year
    Me.Controls("LabelSelectMonth").Caption = VBA.MonthName(Month, True)
    
    Dim startOfMonth As Date
    Dim trackingDate As Date
    Dim startOfMonthDay As Integer
    'calcuate the start of the month and the start of the calendar (top left day)
    startOfMonth = VBA.DateSerial(Year, Month, 1)
    this.CurrentDate = startOfMonth
    startOfMonthDay = VBA.Weekday(startOfMonth, vbSunday)
    trackingDate = DateAdd("d", -startOfMonthDay + 1, startOfMonth)
    
    Dim labelDay As Control
    Dim i As Long
    Dim Holiday As Holiday
    Set Holiday = New Holiday
    Dim holidayName As String
    'loop through all the day controls
    For i = 1 To 42
        'get and set the day controls
        Set labelDay = Me.Controls("LabelDay" & i)
        labelDay.Caption = VBA.Day(trackingDate)
        labelDay.Tag = trackingDate
        Select Case Weekday(trackingDate, vbSunday)
            Case vbSaturday: labelDay.ForeColor = rgbLightSkyBlue
            Case vbSunday: labelDay.ForeColor = rgbLightPink
            Case Else: labelDay.ForeColor = ColorConstants.vbBlack
        End Select
        
        If Holiday.isCompanyHoliday2(trackingDate, holidayName) Then
            labelDay.ForeColor = rgbLightPink
            labelDay.ControlTipText = holidayName
        End If
        
        'make the days not in the current month gray
        If VBA.Month(trackingDate) <> Month Then
            labelDay.ForeColor = rgbGray
        End If
        
'        specialHighlight labelDay
        'move to the next day
        trackingDate = VBA.DateAdd("d", 1, trackingDate)
    Next i
End Sub

Private Sub populateYearPicker(Optional Year As Integer)
    Dim i As Long

    'ref to the control to update
    Dim myControl As Control
    Dim loopStart As Integer
    
    loopStart = Year - 6
    
    'loop through and populate the months
    For i = 1 To 12
        'get a ref to the control to update
        Set myControl = Me.Controls("LabelYear" & i)
        
        'set the string
        myControl.Caption = loopStart
        
        'set the tag to the value we'll act on later
        myControl.Tag = loopStart
        
        'clear any highlight
'        myBGControl.BackColor = 16777215
        
        'see if we should highlight
'        specialHighlight myBGControl, True
        
        'inc the year
        loopStart = loopStart + 1
    Next i
End Sub

Private Sub SwitchToDatePicker()
    Me.MultiPageSelectPicker.value = 0
End Sub

Private Sub SwitchToYearPicker()
    Me.MultiPageSelectPicker.value = 1
End Sub

Private Sub SwitchToMonthPicker()
    Me.MultiPageSelectPicker.value = 2
End Sub

''shows the Year Picker
'Private Sub showYearPicker()
''    Me.FrameYear.Top = 24
''    Me.FrameYear.Visible = True
'End Sub
'
''hides the Year Picker
'Private Sub hideYearPicker()
''    Me.FrameYear.Top = 200
''    Me.FrameYear.Visible = False
'End Sub
'
''shows the Month Picker
'Private Sub showMonthPicker()
''    Me.FrameMonth.Top = 24
''    Me.FrameMonth.Visible = True
'End Sub
'
''hides the Month Picker
'Private Sub hideMonthPicker()
''    Me.FrameMonth.Top = 360
''    Me.FrameMonth.Visible = False
'End Sub

''hides/shows the buttons when they don't apply (for month picker)
'Sub hidePrevNext()
'    Me.LabelPrev.Visible = False
'    Me.LabelNext.Visible = False
'End Sub
'
'Sub showPrevNext()
'    Me.LabelPrev.Visible = True
'    Me.LabelNext.Visible = True
'End Sub

'set the mode for the picker, and updates the UI
Private Sub setPickerMode(Mode As PickerMode)
    this.Mode = Mode
    Select Case Mode
        Case dpNormal
            Call SwitchToDatePicker
'            hideYearPicker
'            hideMonthPicker
'            showPrevNext
        Case dpMonth
            Call SwitchToMonthPicker
'            hideYearPicker
'            hidePrevNext
'            showMonthPicker
        Case dpYear
            Call SwitchToYearPicker
'            hideMonthPicker
'            showPrevNext
'            showYearPicker
    End Select
End Sub

''highlights days in the calendar
'Private Sub specialHighlight(ctrl As Control, Optional picker As Boolean = False)
'
'    'check to see if its already flagged and remove it
'    '(bug with changing months)
'    If ctrl.BackColor = 12632319 Then
'        ctrl.BackColor = 16777215
'    End If
'
'    'see if we should highlight it
'    If picker Then
'        'use highlight picker for the picker control
'        If highlightPicker = ctrl.Tag Then
'            ctrl.BackColor = 12632319
'        End If
'    Else
'        'use highlight date for calendar
'        If highlightDate = ctrl.Tag Then
'            ctrl.BackColor = 12632319
'        End If
'    End If
'
'End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not this.PrevControl Is Nothing Then
        this.PrevControl.Object.BorderStyle = fmBorderStyleNone
        Set this.PrevControl = Nothing
    End If
End Sub

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(Cont As MSForms.IControl)
    Debug.Print Cont.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(Cont As MSForms.IControl, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(Cont As MSForms.IControl)
    Debug.Print Cont.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(Cont As MSForms.IControl)
    Select Case True
        Case VBA.Left$(Cont.Name, 8) = "LabelDay"
            With this
                .Day = VBA.CInt(Cont.Caption)
                DatePicker.SelectionDate = VBA.DateSerial(.Year, .Month, .Day)
            End With
            Unload Me
        Case VBA.Left$(Cont.Name, 9) = "LabelYear"
            With this
                populateDatePicker VBA.CInt(Cont.Caption), VBA.Month(this.CurrentDate)
                setPickerMode dpNormal
            End With
        Case VBA.Left$(Cont.Name, 10) = "LabelMonth"
            With this
                populateDatePicker VBA.Year(this.CurrentDate), VBA.CInt(Cont.Caption)
                setPickerMode dpNormal
            End With
        Case Cont.Name = "LabelSelectYear"
            'set the special highlight
'            highlightPicker = This.Year
            
            'reset the current highlight
'            c1Highlight_Picker = 0
            setPickerMode dpYear
        Case Cont.Name = "LabelSelectMonth"
            'set the special highlight
'            highlightPicker = This.Month
            
            'reset the current highlight
'            c1Highlight_Picker = 0
            setPickerMode dpMonth
        Case Cont.Name = "LabelPrev"
            Select Case this.Mode
                Case dpNormal
                    this.CurrentDate = DateAdd("m", -1, this.CurrentDate)
                    populateDatePicker VBA.Year(this.CurrentDate), VBA.Month(this.CurrentDate)
                Case dpYear
                    this.Year = this.Year - 1
                    populateYearPicker this.Year
                Case dpMonth
                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption - 1
                    this.Year = this.Year - 1
'                    populateYearPicker This.Year
            End Select
        Case Cont.Name = "LabelNext"
            Select Case this.Mode
                Case dpNormal
                    this.CurrentDate = DateAdd("m", 1, this.CurrentDate)
                    populateDatePicker VBA.Year(this.CurrentDate), VBA.Month(this.CurrentDate)
                Case dpYear
                    this.Year = this.Year + 1
                    populateYearPicker this.Year
                Case dpMonth
                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption + 1
                    this.Year = this.Year + 1
            End Select
        Case Cont.Name = "LabelSelectToday"
            DatePicker.SelectionDate = VBA.Now
            Unload Me
        Case Cont.Name = "LabelBackFromYear"
            setPickerMode dpNormal
        Case Cont.Name = "LabelBackFromMonth"
            setPickerMode dpNormal
        Case Else
            Debug.Print Cont.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(Cont As MSForms.IControl)
    Debug.Print Cont.Name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(Cont As MSForms.IControl)
    Debug.Print Cont.Name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(Cont As MSForms.IControl, _
                               ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.Name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(Cont As MSForms.IControl, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
    Debug.Print Cont.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnKeyPress(Cont As MSForms.IControl, _
                                   ByVal KeyAscii As MSForms.IReturnInteger)
    Debug.Print Cont.Name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(Cont As MSForms.IControl, _
                                ByVal KeyCode As MSForms.IReturnInteger, _
                                ByVal Shift As Integer)
    Debug.Print Cont.Name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(Cont As MSForms.IControl)
    Debug.Print Cont.Name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(Cont As MSForms.IControl, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
    Debug.Print Cont.Name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(Cont As MSForms.IControl, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
    Select Case True
        Case VBA.Left$(Cont.Name, 8) = "LabelDay"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
'                    This.PrevControl.BackColor = vbWhite
                End If
            End If
            Set this.PrevControl = Cont
        Case VBA.Left$(Cont.Name, 9) = "LabelYear"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case VBA.Left$(Cont.Name, 10) = "LabelMonth"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelSelectYear"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelSelectMonth"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelPrev"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelNext"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelSelectToday"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelBackFromYear"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Cont.Name = "LabelBackFromMonth"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not this.PrevControl Is Nothing Then
                If Not this.PrevControl Is Cont Then
                    this.PrevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set this.PrevControl = Cont
        Case Else
            If Not this.PrevControl Is Nothing Then
                this.PrevControl.Object.BorderStyle = fmBorderStyleNone
            End If
            Set this.PrevControl = Cont
    End Select
End Sub

Private Sub IControlEvent_OnMouseUp(Cont As MSForms.IControl, _
                                  ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal x As Single, _
                                  ByVal y As Single)
    Debug.Print Cont.Name & " MouseUp:"
End Sub

Private Sub IControlEvent_BeforeDragOver(Cont As MSForms.Control, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print Cont.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(Cont As MSForms.Control, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal x As Single, _
                               ByVal y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    Debug.Print Cont.Name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(Cont As MSForms.Control, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, _
                   ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    Debug.Print Cont.Name & " Error:"
End Sub

Private Sub IControlEvent_AddControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
    Debug.Print Cont.Name & " AddControl:" & Control.Name
End Sub

Private Sub IControlEvent_Layout(Cont As MSForms.Control)
    Debug.Print Cont.Name & " Layout"
End Sub

Private Sub IControlEvent_RemoveControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
    Debug.Print Cont.Name & " RemoveControl:" & Control.Name
End Sub

Private Sub IControlEvent_Scroll(Cont As MSForms.Control, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    Debug.Print Cont.Name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(Cont As MSForms.Control)
'    Debug.Print Cont.Name & " Scroll"
'End Sub

Private Sub IControlEvent_Zoom(Cont As MSForms.Control, Percent As Integer)
    Debug.Print Cont.Name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(Cont As MSForms.Control)
    Debug.Print Cont.Name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(Cont As MSForms.Control)
    Debug.Print Cont.Name & " SpinUp"
End Sub


