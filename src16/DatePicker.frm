VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "UserForm1"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4160
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
    Control As ControlEvents
    PrevControl As MSForms.IControl
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

'月日固定の祝日情報
Private Type HolidayInfoMonthDay
    MonthDay As String
    BeginYear As Long
    EndYear As Long
    Name As String
End Type

'月週曜日固定の祝日情報
Private Type HolidayInfoDayOfWeek
    Month      As Long
    NthWeek    As Long
    DayOfWeek  As Long
    BeginYear  As Long
    EndYear    As Long
    Name       As String
End Type

'「国民の祝日に関する法律」施行年月日
Private Const BEGIN_DATE As Date = #7/20/1948#

'「振替休日」施行年月日
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE As Date = #1/1/2007#

'「国民の休日」施行年月日
Private Const NATIONAL_HOLIDAY_BEGIN_DATE As Date = #12/27/1985#

'エラーコード（パラメータ異常）
Private Const ERROR_INVALID_PARAMETER As Long = &H57

'国民の祝日格納用ディクショナリ
'キー：年月日（DateTime型）
'値　：祝日名
Private dicHoliday_ As Dictionary

Private This As Field

'Private Sub MultiPageSelectPicker_Change()
'    If Me.MultiPageSelectPicker.value = 0 Then
'        PopulateDatePicker This.Year, This.Month
'    ElseIf Me.MultiPageSelectPicker.value = 1 Then
'        PopulateYearPicker This.PeriodStart
'    ElseIf Me.MultiPageSelectPicker.value = 2 Then
'        PopulateMonthPicker This.Year
'    End If
'End Sub

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
    Me.Height = 220
    With Me.LabelBackFromMonth
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelBackFromYear
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelNext
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelPrev
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelSelectMonth
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelSelectToday
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    With Me.LabelSelectYear
        .BackColor = rgbGray
        .BorderColor = rgbSilver
    End With
    FormNonCaption Me, True
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "UserForm_KeyDown"
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub UserForm_Terminate()
'    Select Case True
'        Case TypeName(This.LinkTextBox) = "TextBox"
'            This.LinkTextBox.value = VBA.Fix(This.CurrentDate)
'    End Select
    Set dicHoliday_ = Nothing
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not This.PrevControl Is Nothing Then
        This.PrevControl.Object.BorderStyle = fmBorderStyleNone
        Select Case This.PrevControl.Tag
            Case "Button"
                This.PrevControl.Object.BackStyle = fmBackStyleTransparent
            Case Else
        End Select
        Set This.PrevControl = Nothing
    End If
End Sub

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print CtrlEvt.Control.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)
    If CtrlEvt.Control.Name = "MultiPageSelectPicker" Then
        If Me.MultiPageSelectPicker.value = 0 Then
            PopulateDatePicker This.Year, This.Month
        ElseIf Me.MultiPageSelectPicker.value = 1 Then
            PopulateYearPicker This.PeriodStart
        ElseIf Me.MultiPageSelectPicker.value = 2 Then
            PopulateMonthPicker This.Year
        End If
    End If
    Debug.Print CtrlEvt.Control.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)
    Select Case True
        Case VBA.Left$(CtrlEvt.Control.Name, 8) = "LabelDay"
            With This
'                If VBA.Replace(CtrlEvt.Control.Name, "LabelDay", "") < VBA.CInt(CtrlEvt.Control.Caption) Then
'                    If .Month = 1 Then
'                        .Year = .Year - 1
'                        .Month = 12
'                    Else
'                        .Month = .Month - 1
'                    End If
'                ElseIf VBA.Replace(CtrlEvt.Control.Name, "LabelDay", "") > 28 And VBA.Replace(CtrlEvt.Control.Name, "LabelDay", "") > VBA.CInt(CtrlEvt.Control.Caption) Then
'                    If .Month = 12 Then
'                        .Year = .Year + 1
'                        .Month = 1
'                    Else
'                        .Month = .Month + 1
'                    End If
'                End If
                If VBA.Replace(CtrlEvt.Control.Name, "LabelDay", "") < .StartIndex Then
                    If .Month = 1 Then
                        .Year = .Year - 1
                        .PeriodStart = (.Year \ 20) * 20
                        .Month = 12
                    Else
                        .Month = .Month - 1
                    End If
                ElseIf VBA.Replace(CtrlEvt.Control.Name, "LabelDay", "") > .EndIndex Then
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
            Unload Me
        Case VBA.Left$(CtrlEvt.Control.Name, 9) = "LabelYear"
            With This
                .Year = VBA.CInt(CtrlEvt.Control.Caption)
                .PeriodStart = (.Year \ 20) * 20
                PopulateDatePicker .Year, .Month
                SetPickerMode dpNormal
            End With
        Case VBA.Left$(CtrlEvt.Control.Name, 10) = "LabelMonth"
            With This
                .Month = VBA.CInt(CtrlEvt.Control.Caption)
                PopulateDatePicker .Year, .Month
                SetPickerMode dpNormal
            End With
        Case CtrlEvt.Control.Name = "LabelSelectYear"
            SetPickerMode dpYear
        Case CtrlEvt.Control.Name = "LabelSelectMonth"
            SetPickerMode dpMonth
        Case CtrlEvt.Control.Name = "LabelPrev"
            Select Case This.Mode
                Case dpNormal
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
'                    This.CurrentDate = DateAdd("m", -1, This.CurrentDate)
                    PopulateDatePicker This.Year, This.Month
                Case dpYear
                    If This.PeriodStart - 1 < This.YearMin Then Exit Sub
                    This.PeriodStart = This.PeriodStart - 20
'                    This.PeriodStart = (This.Year \ 20) * 20
                    PopulateYearPicker This.PeriodStart
                Case dpMonth
                    If This.Year - 1 < This.YearMin Then Exit Sub
'                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption - 1
                    This.Year = This.Year - 1
                    This.PeriodStart = (This.Year \ 20) * 20
                    PopulateMonthPicker This.Year
            End Select
        Case CtrlEvt.Control.Name = "LabelNext"
            Select Case This.Mode
                Case dpNormal
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
'                    This.CurrentDate = DateAdd("m", 1, This.CurrentDate)
                    PopulateDatePicker This.Year, This.Month
                Case dpYear
                    If This.PeriodStart + 1 > This.YearMax Then Exit Sub
                    This.PeriodStart = This.PeriodStart + 20
'                    This.PeriodStart = (This.Year \ 20) * 20
                    PopulateYearPicker This.PeriodStart
                Case dpMonth
                    If This.Year + 1 > This.YearMax Then Exit Sub
                    This.Year = This.Year + 1
                    This.PeriodStart = (This.Year \ 20) * 20
'                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption + 1
                    PopulateMonthPicker This.Year
            End Select
        Case CtrlEvt.Control.Name = "LabelSelectToday"
            This.CurrentDate = This.Today
            Call SetDateToTextBox
            Unload Me
        Case CtrlEvt.Control.Name = "LabelBackFromYear"
            SetPickerMode dpNormal
        Case CtrlEvt.Control.Name = "LabelBackFromMonth"
            SetPickerMode dpNormal
        Case CtrlEvt.Control.Name = "LabelClose"
            Unload Me
        Case Else
            Debug.Print CtrlEvt.Control.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case CtrlEvt.Control.Name = "LabelPrev"
            Call IControlEvent_OnClick(CtrlEvt)
            DoEvents
            Cancel = True
        Case CtrlEvt.Control.Name = "LabelNext"
            Call IControlEvent_OnClick(CtrlEvt)
            DoEvents
            Cancel = True
    End Select
    Debug.Print CtrlEvt.Control.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print CtrlEvt.Control.Name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(CtrlEvt As ControlEvent, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
    Debug.Print CtrlEvt.Control.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub IControlEvent_OnKeyPress(CtrlEvt As ControlEvent, _
                                   ByVal KeyAscii As MSForms.IReturnInteger)
    Debug.Print CtrlEvt.Control.Name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(CtrlEvt As ControlEvent, _
                                ByVal KeyCode As MSForms.IReturnInteger, _
                                ByVal Shift As Integer)
    Debug.Print CtrlEvt.Control.Name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Call Hover(CtrlEvt)
End Sub

Private Sub IControlEvent_OnMouseUp(CtrlEvt As ControlEvent, _
                                  ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseUp:"
End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print CtrlEvt.Control.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    Debug.Print CtrlEvt.Control.Name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(CtrlEvt As ControlEvent, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, _
                   ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    Debug.Print CtrlEvt.Control.Name & " Error:"
End Sub

Private Sub IControlEvent_OnAddControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
    Debug.Print CtrlEvt.Control.Name & " AddControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnLayout(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " Layout"
End Sub

Private Sub IControlEvent_OnRemoveControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
    Debug.Print CtrlEvt.Control.Name & " RemoveControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    Debug.Print CtrlEvt.Control.Name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Scroll"
'End Sub

Private Sub IControlEvent_OnZoom(CtrlEvt As ControlEvent, Percent As Integer)
    Debug.Print CtrlEvt.Control.Name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " SpinUp"
End Sub

'-------------------------------------------------------------
Public Sub Init(TextBox As MSForms.TextBox, Optional YearMin As Long = 2000, Optional YearMax As Long = 2050)
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
'    If pos.y - TextBox.Height < 0 Then pos.y = DispSizePoint.x - Me.Width
    
    Me.Top = pos.Y
    Me.Left = pos.X
    
'    Debug.Print MainForm.Top * LogicalPixcel.y / 72 & ":" & MainForm.Left * LogicalPixcel.x / 72
'    Debug.Print TextBox.Top * LogicalPixcel.y / 72 & ":" & TextBox.Left * LogicalPixcel.x / 72
'    Debug.Print LogicalPixcel.y / 72 & ":" & LogicalPixcel.x / 72
'    Me.Top = MainForm.Top * LogicalPixcel.y / 72 + (120 + TextBox.Top + TextBox.Height) / 1.2 ' * LogicalPixcel.y / 72
'    Me.Left = (MainForm.Left + 40 + TextBox.Left) ' * LogicalPixcel.x / 72
    
    Debug.Print "Pos:" & Me.Top & ":" & Me.Left
    
'    Me.Top = TextBox.Top + TextBox.Height
'    Me.Left = TextBox.Left
    
    This.YearMin = YearMin
    This.YearMax = YearMax
    
    Call MakeHolidayDictionary
    
    Call PopulateDatePicker(VBA.Year(This.CurrentDate), VBA.Month(This.CurrentDate))
    SetPickerMode dpNormal
    Set This.Control = New ControlEvents
    With This.Control
        .Parent = Me
        .Init
    End With
    
    Me.Show
End Sub

Private Sub SetDateToTextBox()
    Select Case True
        Case TypeName(This.LinkTextBox) = "TextBox"
            This.LinkTextBox.value = VBA.Fix(This.CurrentDate)
    End Select
End Sub

'各コントロールのTagプロパティに応じてMouseHover時の処理を規定する
'処理対象とするコントロールにはコード内またはFormデザイン時にTagプロパティを設定しておくこと
Private Sub Hover(CtrlEvt As ControlEvent)
    'MouseOver
    Select Case CtrlEvt.Control.Tag
        Case "SelectDay", "SelectYear", "SelectMonth"
            CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
        Case "Button"
            CtrlEvt.Control.Object.BackStyle = fmBackStyleOpaque
            CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
        Case Else
    End Select
    
    'MouseOut
    If Not This.PrevControl Is Nothing Then
        If Not This.PrevControl Is CtrlEvt.Control Then
            This.PrevControl.Object.BorderStyle = fmBorderStyleNone
            Select Case This.PrevControl.Tag
                Case "Button"
                    This.PrevControl.Object.BackStyle = fmBackStyleTransparent
            End Select
        End If
    End If
    
    Set This.PrevControl = CtrlEvt.Control
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
    Me.Controls("LabelSelectYear").Caption = Year & "年"
    Me.Controls("LabelSelectYear").Visible = True
    Me.Controls("LabelSelectMonth").Caption = VBA.MonthName(Month, False)
    Me.Controls("LabelSelectMonth").Visible = True
    
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
        Set labelDay = Me.Controls("LabelDay" & i)
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
            labelDay.BackStyle = fmBackStyleOpaque
        Else
            labelDay.BackStyle = fmBackStyleTransparent
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
    Me.Controls("LabelSelectYear").Caption = Year & "年"
    Me.Controls("LabelSelectYear").Visible = True
    Me.Controls("LabelSelectMonth").Visible = False
    
    Dim labelMonth As Control
    Dim i As Long
    For i = 1 To 12
        Set labelMonth = Me.Controls("LabelMonth" & i)
        If This.Year = VBA.Year(This.Today) And i = VBA.Month(This.Today) Then
            labelMonth.BackStyle = fmBackStyleOpaque
        Else
            labelMonth.BackStyle = fmBackStyleTransparent
        End If
    Next i
End Sub

Private Sub PopulateYearPicker(Optional PeriodStartYear As Integer = 0)
    If PeriodStartYear = 0 Then PeriodStartYear = (This.Year \ 20) * 20
    
    Me.Controls("LabelSelectYear").Visible = False
    Me.Controls("LabelSelectMonth").Visible = False
    
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
    
    Dim i As Long
    For i = 1 To 20
        Set labelYear = Me.Controls("LabelYear" & i)
        
        labelYear.Caption = loopStart
               
        If loopStart < This.YearMin Or loopStart > This.YearMax Then
            labelYear.Enabled = False
        Else
            labelYear.Enabled = True
            If loopStart = VBA.Year(This.Today) Then
                labelYear.BackStyle = fmBackStyleOpaque
            Else
                labelYear.BackStyle = fmBackStyleTransparent
            End If
        End If
        
        loopStart = loopStart + 1
    Next i
End Sub

'Private Sub SwitchToDatePicker()
'    Me.MultiPageSelectPicker.value = 0
'End Sub
'
'Private Sub SwitchToYearPicker()
'    Me.MultiPageSelectPicker.value = 1
'End Sub
'
'Private Sub SwitchToMonthPicker()
'    Me.MultiPageSelectPicker.value = 2
'End Sub

Private Sub SetPickerMode(Mode As PickerMode)
    This.Mode = Mode
    Me.MultiPageSelectPicker.value = Mode
'    Select Case Mode
'        Case dpNormal
'            Call SwitchToDatePicker
'        Case dpMonth
'            Call SwitchToMonthPicker
'        Case dpYear
'            Call SwitchToYearPicker
'    End Select
End Sub

'指定日が会社休日か？
Public Function IsHoliday(ByVal dtDate As Date, ByRef HolidayName As String) As Boolean
    HolidayName = ""
    Dim dtDateW As Date

    '時分秒データを切り捨てる
    dtDateW = VBA.Fix(dtDate)

    If dtDateW < BEGIN_DATE Then
        ERR.Raise ERROR_INVALID_PARAMETER, "IsHoliday", Format$(dtDateW, "yyyy/mm/dd") & "は、適用範囲外です。"
        Exit Function
    ElseIf VBA.Year(dtDateW) > This.YearMax Then
        ERR.Raise ERROR_INVALID_PARAMETER, "IsHoliday", This.YearMax + 1 & "年以降は、適用範囲外です。"
        Exit Function
    End If

    IsHoliday = dicHoliday_.Exists(dtDateW)
    If IsHoliday Then HolidayName = dicHoliday_.Item(dtDateW)
End Function

'Dictionaryへ祝日情報を格納
Private Sub MakeHolidayDictionary()
    Dim HolidayInfoMD() As HolidayInfoMonthDay
    Dim HolidayInfoDOW() As HolidayInfoDayOfWeek

    Set dicHoliday_ = New Dictionary
    
    '月日固定の祝日情報
    Call getNationalHolidayInfoMD(HolidayInfoMD, SheetList.ListObjects("T_月日固定休日"))

    '月週曜日固定の祝日情報
    Call getNationalHolidayInfoWN(HolidayInfoDOW, SheetList.ListObjects("T_月週曜日固定休日"))
    
    'Dictionaryへ追加
    Call AddToDictionary(HolidayInfoMD, HolidayInfoDOW)
End Sub

'月日固定の祝日情報生成
Private Sub getNationalHolidayInfoMD(ByRef HolidayInfo() As HolidayInfoMonthDay, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
    
        Dim i As Long
        For i = 1 To .ListRows.Count
            HolidayInfo(i).MonthDay = .ListColumns("月日").DataBodyRange(i)
            HolidayInfo(i).BeginYear = CLng(.ListColumns("適用開始年").DataBodyRange(i))
            HolidayInfo(i).EndYear = CLng(.ListColumns("適用終了年").DataBodyRange(i))
            HolidayInfo(i).Name = .ListColumns("名前").DataBodyRange(i)
        Next i
    End With
End Sub

'月週曜日固定の祝日情報生成
Private Sub getNationalHolidayInfoWN(ByRef HolidayInfo() As HolidayInfoDayOfWeek, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
        
        Dim i As Long
        For i = 1 To .ListRows.Count
            HolidayInfo(i).Month = CLng(.ListColumns("月").DataBodyRange(i))
            HolidayInfo(i).NthWeek = CLng(.ListColumns("週").DataBodyRange(i))
            Select Case .ListColumns("曜日").DataBodyRange(i)
                Case "日": HolidayInfo(i).DayOfWeek = 1
                Case "月": HolidayInfo(i).DayOfWeek = 2
                Case "火": HolidayInfo(i).DayOfWeek = 3
                Case "水": HolidayInfo(i).DayOfWeek = 4
                Case "木": HolidayInfo(i).DayOfWeek = 5
                Case "金": HolidayInfo(i).DayOfWeek = 6
                Case "土": HolidayInfo(i).DayOfWeek = 7
            End Select
            HolidayInfo(i).BeginYear = CLng(.ListColumns("適用開始年").DataBodyRange(i))
            HolidayInfo(i).EndYear = CLng(.ListColumns("適用終了年").DataBodyRange(i))
            HolidayInfo(i).Name = .ListColumns("名前").DataBodyRange(i)
        Next i
    End With
End Sub

'祝日情報をDictionaryへ格納
Private Sub AddToDictionary(ByRef HolidayInfoMD() As HolidayInfoMonthDay, ByRef HolidayInfoDOW() As HolidayInfoDayOfWeek)
    Dim dtHoliday           As Date
    Dim lAddedDays          As Long
    Dim dtAdded()           As Date
    Dim existsHoliday       As Boolean
    Dim lYear               As Long
    Dim i                   As Long

    For lYear = This.YearMin To This.YearMax
        '年間の祝日格納用配列クリア
        lAddedDays = 0
        ReDim dtAdded(lAddedDays)

        '月日固定の祝日
        For i = 0 To UBound(HolidayInfoMD)
            '適用期間のみを対象とする
            If HolidayInfoMD(i).BeginYear <= lYear And HolidayInfoMD(i).EndYear >= lYear Then
                dtHoliday = CDate(CStr(lYear) & "/" & HolidayInfoMD(i).MonthDay)

                dicHoliday_.Add dtHoliday, HolidayInfoMD(i).Name

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '月週曜日固定の祝日
        For i = 0 To UBound(HolidayInfoDOW)
            '適用期間のみを対象とする
            If HolidayInfoDOW(i).BeginYear <= lYear And HolidayInfoDOW(i).EndYear >= lYear Then
                dtHoliday = getNthWeeksDayOfWeek(lYear, HolidayInfoDOW(i).Month, HolidayInfoDOW(i).NthWeek, HolidayInfoDOW(i).DayOfWeek)

                dicHoliday_.Add dtHoliday, HolidayInfoDOW(i).Name

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '春分の日
        dtHoliday = getVernalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "春分の日"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '秋分の日
        dtHoliday = getAutumnalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "秋分の日"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '振替休日
        For i = 0 To lAddedDays - 1
            existsHoliday = existsSubstituteHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "振替休日"
            End If
        Next i

        '国民の休日
        For i = 0 To lAddedDays - 1
            existsHoliday = existsNationalHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "国民の休日"
            End If
        Next i

        Erase dtAdded
    Next lYear
End Sub

'振替休日の有無
'　祝日（dtDate）に対する振替休日の有無（ある場合は、dtSubstituteHolidayに代入される）
Private Function existsSubstituteHoliday(ByVal dtDate As Date, ByRef dtSubstituteHoliday As Date) As Boolean
    Dim dtNextDay   As Date

    existsSubstituteHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDateが祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If dtDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And dtDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(dtDate) = vbSunday Then
            '祝日が日曜日であれば、翌日（月曜日）が振替休日
            dtSubstituteHoliday = DateAdd("d", 1, dtDate)

            existsSubstituteHoliday = True
        End If
    ElseIf dtDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '「国民の祝日」が日曜日に当たるときは、その日後においてその日に最も近い「国民の祝日」でない日を休日とする
        If Weekday(dtDate) = vbSunday Then
            dtNextDay = DateAdd("d", 1, dtDate)

            '直近の祝日でない日を取得
            Do Until dicHoliday_.Exists(dtNextDay) = False
                dtNextDay = DateAdd("d", 1, dtNextDay)
            Loop

            dtSubstituteHoliday = dtNextDay

            existsSubstituteHoliday = True
        End If
    End If
End Function

'国民の休日の有無
'　祝日（dtDate）に対す国民の休日の有無（ある場合は、dtNationalHolidayに代入される）
Private Function existsNationalHoliday(ByVal dtDate As Date, ByRef dtNationalHoliday As Date) As Boolean
    Dim dtBaseDay   As Date
    Dim dtNextDay   As Date

    existsNationalHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDateが祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If dtDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        dtBaseDay = DateAdd("d", 1, dtDate)

        '直近の祝日でない日を取得
        Do Until dicHoliday_.Exists(dtBaseDay) = False
            dtBaseDay = DateAdd("d", 1, dtBaseDay)
        Loop

        '日曜日であれば対象外
        If Weekday(dtBaseDay) <> vbSunday Then
            dtNextDay = DateAdd("d", 1, dtBaseDay)

            '翌日が祝日であれば対象
            If dicHoliday_.Exists(dtNextDay) = True Then
                existsNationalHoliday = True

                dtNationalHoliday = dtBaseDay
            End If
        End If
    End If
End Function

'月の第N W曜日の日時を取得
Private Function getNthWeeksDayOfWeek(ByVal lYear As Long, _
                                      ByVal lMonth As Long, _
                                      ByVal lNth As Long, _
                                      ByVal lDayOfWeek As VbDayOfWeek) As Date
    Dim dt1stDate       As Date
    Dim lDayOfWeek1st   As Long
    Dim lOffset         As Long

    '指定年月の１日を取得
    dt1stDate = DateSerial(lYear, lMonth, 1)

    '１日の曜日を取得
    lDayOfWeek1st = Weekday(dt1stDate)

    '指定日へのオフセットを取得
    lOffset = lDayOfWeek - lDayOfWeek1st

    If lDayOfWeek1st > lDayOfWeek Then
        lOffset = lOffset + 7
    End If

    lOffset = lOffset + 7 * (lNth - 1)

    getNthWeeksDayOfWeek = DateAdd("d", lOffset, dt1stDate)
End Function

'春分の日
Private Function getVernalEquinoxDay(ByVal lYear As Long) As Date
    Dim lDay    As Long

    lDay = Int(20.8431 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getVernalEquinoxDay = DateSerial(lYear, 3, lDay)
End Function

'秋分の日
Private Function getAutumnalEquinoxDay(ByVal lYear As Long) As Date
    Dim lDay    As Long

    lDay = Int(23.2488 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getAutumnalEquinoxDay = DateSerial(lYear, 9, lDay)
End Function
