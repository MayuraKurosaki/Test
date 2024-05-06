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
    dpMonth = 1
    dpYear = 2
End Enum

Private Type TState
    Control As ControlEvents
    PrevControl As MSForms.IControl
    Mode As PickerMode '0=normal, 1=months, 2=years
    Today As Date
    Year As Integer
    Month As Integer
    Day As Integer
    YearMin As Long
    YearMax As Long
    StartOfMonthDay As Integer
    StartIndex As Long
    EndIndex As Long
    CurrentDate As Date
    LinkTextBox As MSForms.TextBox
End Type

'�����Œ�̏j�����
Private Type FixMD
    MD         As String
    BeginYear  As Long
    EndYear    As Long
    Name       As String
End Type

'���T�j���Œ�̏j�����
Private Type FixWN
    Month      As Long
    NthWeek    As Long
    DayOfWeek  As Long
    BeginYear  As Long
    EndYear    As Long
    Name       As String
End Type

'�u�����̏j���Ɋւ���@���v�{�s�N����
Private Const BEGIN_DATE As Date = #7/20/1948#

'�u�U�֋x���v�{�s�N����
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE As Date = #1/1/2007#

'�u�����̋x���v�{�s�N����
Private Const NATIONAL_HOLIDAY_BEGIN_DATE As Date = #12/27/1985#

''�N����
'Private Const YEAR_MIN As Long = 2000
'
''�N���
'Private Const YEAR_MAX As Long = 2050

'�G���[�R�[�h�i�p�����[�^�ُ�j
Private Const ERROR_INVALID_PARAMETER As Long = &H57

'�����̏j���i�[�p�f�B�N�V���i��
'�L�[�F�N�����iDateTime�^�j
'�l�@�F�j����
Private dicHoliday_ As Dictionary

'Private lInitializedLastYear_ As Long

Private This As TState

Private Sub MultiPageSelectPicker_Change()
    If Me.MultiPageSelectPicker.value = 0 Then
        PopulateDatePicker This.Year, This.Month
    ElseIf Me.MultiPageSelectPicker.value = 1 Then
        PopulateYearPicker This.Year
    ElseIf Me.MultiPageSelectPicker.value = 2 Then
        PopulateMonthPicker This.Year
    End If
End Sub

Private Sub MultiPageSelectPicker_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub MultiPageSelectPicker_Layout(ByVal Index As Long)

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
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

'--------------------�C���^�[�t�F�C�X����R�[���o�b�N����郁���o�֐�
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
            With This
'                If VBA.Replace(Cont.Name, "LabelDay", "") < VBA.CInt(Cont.Caption) Then
'                    If .Month = 1 Then
'                        .Year = .Year - 1
'                        .Month = 12
'                    Else
'                        .Month = .Month - 1
'                    End If
'                ElseIf VBA.Replace(Cont.Name, "LabelDay", "") > 28 And VBA.Replace(Cont.Name, "LabelDay", "") > VBA.CInt(Cont.Caption) Then
'                    If .Month = 12 Then
'                        .Year = .Year + 1
'                        .Month = 1
'                    Else
'                        .Month = .Month + 1
'                    End If
'                End If
                If VBA.Replace(Cont.Name, "LabelDay", "") < This.StartIndex Then
                    If .Month = 1 Then
                        .Year = .Year - 1
                        .Month = 12
                    Else
                        .Month = .Month - 1
                    End If
                ElseIf VBA.Replace(Cont.Name, "LabelDay", "") > This.EndIndex Then
                    If .Month = 12 Then
                        .Year = .Year + 1
                        .Month = 1
                    Else
                        .Month = .Month + 1
                    End If
                End If
                .Day = VBA.CInt(Cont.Caption)
                .CurrentDate = VBA.DateSerial(.Year, .Month, .Day)
            End With
            Call SetDateToTextBox
            Unload Me
        Case VBA.Left$(Cont.Name, 9) = "LabelYear"
            With This
                .Year = VBA.CInt(Cont.Caption)
                PopulateDatePicker .Year, .Month
                SetPickerMode dpNormal
            End With
        Case VBA.Left$(Cont.Name, 10) = "LabelMonth"
            With This
                .Month = VBA.CInt(Cont.Caption)
                PopulateDatePicker .Year, .Month
                SetPickerMode dpNormal
            End With
        Case Cont.Name = "LabelSelectYear"
            SetPickerMode dpYear
        Case Cont.Name = "LabelSelectMonth"
            SetPickerMode dpMonth
        Case Cont.Name = "LabelPrev"
            Select Case This.Mode
                Case dpNormal
                    If This.Month = 1 Then
                        If This.Year - 1 < This.YearMin Then
                            Exit Sub
                        Else
                            This.Year = This.Year - 1
                            This.Month = 12
                        End If
                    Else
                        This.Month = This.Month - 1
                    End If
'                    This.CurrentDate = DateAdd("m", -1, This.CurrentDate)
                    PopulateDatePicker This.Year, This.Month
                Case dpYear
                    If This.Year - 1 < This.YearMin + 6 Then Exit Sub
                    This.Year = This.Year - 1
                    PopulateYearPicker This.Year
                Case dpMonth
                    If This.Year - 1 < This.YearMin Then Exit Sub
                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption - 1
                    This.Year = This.Year - 1
                    PopulateMonthPicker
            End Select
        Case Cont.Name = "LabelNext"
            Select Case This.Mode
                Case dpNormal
                    If This.Month = 12 Then
                        If This.Year + 1 > This.YearMax Then
                            Exit Sub
                        Else
                            This.Year = This.Year + 1
                            This.Month = 1
                        End If
                    Else
                        This.Month = This.Month + 1
                    End If
'                    This.CurrentDate = DateAdd("m", 1, This.CurrentDate)
                    PopulateDatePicker This.Year, This.Month
                Case dpYear
                    If This.Year + 1 > This.YearMax - 5 Then Exit Sub
                    This.Year = This.Year + 1
                    PopulateYearPicker This.Year
                Case dpMonth
                    If This.Year + 1 > This.YearMax Then Exit Sub
                    This.Year = This.Year + 1
                    Me.LabelSelectYear.Caption = Me.LabelSelectYear.Caption + 1
                    PopulateMonthPicker
            End Select
        Case Cont.Name = "LabelSelectToday"
            This.CurrentDate = This.Today
            Call SetDateToTextBox
            Unload Me
        Case Cont.Name = "LabelBackFromYear"
            SetPickerMode dpNormal
        Case Cont.Name = "LabelBackFromMonth"
            SetPickerMode dpNormal
        Case Else
            Debug.Print Cont.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case Cont.Name = "LabelPrev"
            Call IControlEvent_OnClick(Cont)
            DoEvents
            Cancel = True
        Case Cont.Name = "LabelNext"
            Call IControlEvent_OnClick(Cont)
            DoEvents
            Cancel = True
    End Select
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
    If KeyCode = vbKeyEscape Then Unload Me
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
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Debug.Print Cont.Name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(Cont As MSForms.IControl, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Call Hover(Cont)
End Sub

Private Sub IControlEvent_OnMouseUp(Cont As MSForms.IControl, _
                                  ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)
    Debug.Print Cont.Name & " MouseUp:"
End Sub

Private Sub IControlEvent_BeforeDragOver(Cont As MSForms.Control, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print Cont.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(Cont As MSForms.Control, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
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
        .parent = Me
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

'�e�R���g���[����Tag�v���p�e�B�ɉ�����MouseHover���̏������K�肷��
'�����ΏۂƂ���R���g���[���ɂ̓R�[�h���܂���Form�f�U�C������Tag�v���p�e�B��ݒ肵�Ă�������
Private Sub Hover(Cont As MSForms.IControl)
    'MouseOver
    Select Case Cont.Tag
        Case "SelectDay", "SelectYear", "SelectMonth"
            Cont.Object.BorderStyle = fmBorderStyleSingle
        Case "Button"
            Cont.Object.BackStyle = fmBackStyleOpaque
            Cont.Object.BorderStyle = fmBorderStyleSingle
        Case Else
    End Select
    
    'MouseOut
    If Not This.PrevControl Is Nothing Then
        If Not This.PrevControl Is Cont Then
            This.PrevControl.Object.BorderStyle = fmBorderStyleNone
            Select Case This.PrevControl.Tag
                Case "Button"
                    This.PrevControl.Object.BackStyle = fmBackStyleTransparent
            End Select
        End If
    End If
    
    Set This.PrevControl = Cont
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
    
    Me.Controls("LabelSelectYear").Caption = Year
    Me.Controls("LabelSelectMonth").Caption = VBA.MonthName(Month, True)
'    Me.Controls("LabelSelectToday").Caption = "�����F" & This.Today
    
    Dim startOfMonth As Date
'    Dim StartOfMonthDay As Integer
    Dim trackingDate As Date
    startOfMonth = VBA.DateSerial(Year, Month, 1)
    This.StartOfMonthDay = VBA.Weekday(startOfMonth, vbSunday)
    trackingDate = DateAdd("d", -This.StartOfMonthDay + 1, startOfMonth)
    
    Dim captionDay As Integer: captionDay = 0
    Dim labelDay As Control
    Dim i As Long
'    Dim Holiday As Holiday
'    Set Holiday = New Holiday
    Dim holidayName As String
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
'        If Holiday.isCompanyHoliday2(trackingDate, holidayName) Then

        If VBA.Year(trackingDate) < This.YearMin Or VBA.Year(trackingDate) > This.YearMax Then
            Debug.Print trackingDate
            labelDay.Enabled = False
        Else
            If isCompanyHoliday(trackingDate, holidayName) Then
                labelDay.ForeColor = rgbLightCoral
                labelDay.ControlTipText = holidayName
            End If
        End If
        
        If VBA.Month(trackingDate) <> Month Then
            labelDay.ForeColor = rgbGray
        End If
        
'        If VBA.Year(trackingDate) < YEAR_MIN Or VBA.Year(trackingDate) > YEAR_MAX Then
'            Debug.Print trackingDate
'            labelDay.Enabled = False
'        End If
'        If VBA.Year(trackingDate) > YEAR_MAX Then
'            labelDay.Enabled = False
'        End If
        
        trackingDate = VBA.DateAdd("d", 1, trackingDate)
    Next i
End Sub

Private Sub PopulateMonthPicker(Optional Year As Integer = 0)
    If Year = 0 Then Year = VBA.Year(This.Today)
    
    Dim myControl As Control
    Dim i As Long
    For i = 1 To 12
        Set myControl = Me.Controls("LabelMonth" & i)
        If This.Year = VBA.Year(This.Today) And i = VBA.Month(This.Today) Then
            myControl.BackStyle = fmBackStyleOpaque
        Else
            myControl.BackStyle = fmBackStyleTransparent
        End If
    Next i
End Sub

Private Sub PopulateYearPicker(Optional Year As Integer = 0)
    If Year = 0 Then Year = VBA.Year(This.Today)
    If Year - 6 < This.YearMin Then Year = This.YearMin + 6
    If Year + 5 > This.YearMax Then Year = This.YearMax - 5
    
    Dim myControl As Control
    Dim loopStart As Integer
    
    loopStart = Year - 6
    
    Dim i As Long
    For i = 1 To 12
        Set myControl = Me.Controls("LabelYear" & i)
        
        myControl.Caption = loopStart
               
        If loopStart = VBA.Year(This.Today) Then
            myControl.BackStyle = fmBackStyleOpaque
        Else
            myControl.BackStyle = fmBackStyleTransparent
        End If
        
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

Private Sub SetPickerMode(Mode As PickerMode)
    This.Mode = Mode
    Select Case Mode
        Case dpNormal
            Call SwitchToDatePicker
        Case dpMonth
            Call SwitchToMonthPicker
        Case dpYear
            Call SwitchToYearPicker
    End Select
End Sub

'�w�������Ћx�����H
Public Function isCompanyHoliday(ByVal dtDate As Date, ByRef sHolidayName As String) As Boolean
    sHolidayName = ""
    Dim dtDateW As Date

    '�����b�f�[�^��؂�̂Ă�
'    dtDateW = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate))
    dtDateW = VBA.Fix(dtDate)

    If dtDateW < BEGIN_DATE Then
        ERR.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", Format$(dtDateW, "yyyy/mm/dd") & "�́A�K�p�͈͊O�ł��B"
        Exit Function
    ElseIf Year(dtDateW) > This.YearMax Then
        ERR.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", This.YearMax + 1 & "�N�ȍ~�́A�K�p�͈͊O�ł��B"
        Exit Function
'    ElseIf Year(dtDateW) > InitializedLastYear Then
'        ERR.Raise ERROR_INVALID_PARAMETER, "isCompanyHoliday", Format$(dtDateW, "yyyy�N") & "�́A�f�[�^����������Ă��Ȃ����߁A����ł��܂���B" _
'                            & vbCrLf & "reInitialize���\�b�h�őΏ۔N��ݒ��A�ēx�m�F���Ă݂ĉ������B"
'        Exit Function
    End If

    isCompanyHoliday = dicHoliday_.Exists(dtDateW)
    If isCompanyHoliday Then sHolidayName = dicHoliday_.Item(dtDateW)
End Function

'Dictionary�֏j�������i�[
Private Sub MakeHolidayDictionary()
    Dim uFixMD()    As FixMD
    Dim uFixWN()    As FixWN

    Set dicHoliday_ = New Dictionary
    
    '�����Œ�̏j�����
    Call getNationalHolidayInfoMD(uFixMD)

    '���T�j���Œ�̏j�����
    Call getNationalHolidayInfoWN(uFixWN)

    'Dictionary�֒ǉ�
    Call add2Dictionary(This.YearMax, uFixMD, uFixWN)
End Sub

'�����Œ�̏j����񐶐�
Private Sub getNationalHolidayInfoMD(ByRef uFixMD() As FixMD)
    Dim sFixMD(29)  As String   '�j���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��

    '//////////////////////////////////////////////////
    '               �����Œ�̏j��
    '//////////////////////////////////////////////////
    '�K�p�J�n�N�ɂ���
    '�@���U�i1/1�j
    '�@���l�̓��i1/15�j
    '�@�V�c�a�����i4/29�j
    '�@���@�L�O���i5/3�j
    '�@���ǂ��̓��i5/5�j
    '�̂T�́A�u�����̏j���Ɋւ���@���v�{�s�N�i1948�N�j�ɐ��肳��Ă��邪
    '���@�̎{�s��7/20�ł���A����ȑO�ƂȂ邽�߁A�K�p�J�n�N�𗂔N�i1949�N�j�ɕ␳���Ă���B
    '
    '����,�K�p�J�n�N,�K�p�I���N,���O
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixMD(0) = "01/01,1949,9999,����"          '�K�p�J�n�N�␳�ς�
    sFixMD(1) = "01/15,1949,1999,���l�̓�"      '�K�p�J�n�N�␳�ς�
    sFixMD(2) = "02/11,1967,9999,�����L�O�̓�"
    sFixMD(3) = "02/23,2020,9999,�V�c�a����"    '�K�p�J�n�N�␳�ς�
    sFixMD(4) = "02/24,1989,1989,���a�V�c�̑�r�̗�"
    sFixMD(5) = "04/10,1959,1959,�c���q���m�e���̌����̋V"
    sFixMD(6) = "04/29,1949,1988,�V�c�a����"    '�K�p�J�n�N�␳�ς�
    sFixMD(7) = "04/29,1989,2006,�݂ǂ�̓�"
    sFixMD(8) = "04/29,2007,9999,���a�̓�"
    sFixMD(9) = "05/01,2019,2019,�V�c�̑���"
    sFixMD(10) = "05/03,1949,9999,���@�L�O��"    '�K�p�J�n�N�␳�ς�
    sFixMD(11) = "05/04,2007,9999,�݂ǂ�̓�"
    sFixMD(12) = "05/05,1949,9999,���ǂ��̓�"    '�K�p�J�n�N�␳�ς�
    sFixMD(13) = "06/09,1993,1993,�c���q���m�e���̌����̋V"
    sFixMD(14) = "07/20,1996,2002,�C�̓�"
    sFixMD(15) = "07/22,2021,2021,�C�̓�"
    sFixMD(16) = "07/23,2020,2020,�C�̓�"
    sFixMD(17) = "07/23,2021,2021,�X�|�[�c�̓�"
    sFixMD(18) = "07/24,2020,2020,�X�|�[�c�̓�"
    sFixMD(19) = "08/08,2021,2021,�R�̓�"
    sFixMD(20) = "08/10,2020,2020,�R�̓�"
    sFixMD(21) = "08/11,2016,2019,�R�̓�"
    sFixMD(22) = "08/11,2022,9999,�R�̓�"
    sFixMD(23) = "09/15,1966,2002,�h�V�̓�"
    sFixMD(24) = "10/10,1966,1999,�̈�̓�"
    sFixMD(25) = "10/22,2019,2019,���ʗ琳�a�̋V"
    sFixMD(26) = "11/03,1948,9999,�����̓�"
    sFixMD(27) = "11/12,1990,1990,���ʗ琳�a�̋V"
    sFixMD(28) = "11/23,1948,9999,�ΘJ���ӂ̓�"
    sFixMD(29) = "12/23,1989,2018,�V�c�a����"

    ReDim uFixMD(UBound(sFixMD))

    Dim sResult() As String
    Dim i As Long
    For i = 0 To UBound(sFixMD)
        sResult = Split(sFixMD(i), ",")

        uFixMD(i).MD = sResult(0)
        uFixMD(i).BeginYear = CLng(sResult(1))
        uFixMD(i).EndYear = CLng(sResult(2))
        uFixMD(i).Name = sResult(3)
    Next i
    Erase sFixMD
End Sub

'���T�j���Œ�̏j����񐶐�
Private Sub getNationalHolidayInfoWN(ByRef uFixWN() As FixWN)
    Dim sFixWN(5)   As String   '�j���f�[�^��ǉ��폜�����ꍇ�A���̔z��v�f����ύX���邱��

    '//////////////////////////////////////////////////
    '               ���T�j���Œ�̏j��
    '//////////////////////////////////////////////////
    '��,�T,�j��,�K�p�J�n�N,�K�p�I���N,���O
    '�j���F�� 1
    '�@�@�@�� 2
    '�@�@�@�� 3
    '�@�@�@�� 4
    '�@�@�@�� 5
    '�@�@�@�� 6
    '�@�@�@�y 7
    '�K�p�I���N�G9999�́A���݂��K�p��
    sFixWN(0) = "01,2,2,2000,9999,���l�̓�"
    sFixWN(1) = "07,3,2,2003,2019,�C�̓�"
    sFixWN(2) = "07,3,2,2022,9999,�C�̓�"
    sFixWN(3) = "09,3,2,2003,9999,�h�V�̓�"
    sFixWN(4) = "10,2,2,2000,2019,�̈�̓�"
    sFixWN(5) = "10,2,2,2022,9999,�X�|�[�c�̓�"

    ReDim uFixWN(UBound(sFixWN))

    Dim sResult() As String
    Dim i As Long
    For i = 0 To UBound(sFixWN)
        sResult = Split(sFixWN(i), ",")

        uFixWN(i).Month = CLng(sResult(0))
        uFixWN(i).NthWeek = CLng(sResult(1))
        uFixWN(i).DayOfWeek = CLng(sResult(2))
        uFixWN(i).BeginYear = CLng(sResult(3))
        uFixWN(i).EndYear = CLng(sResult(4))
        uFixWN(i).Name = sResult(5)
    Next i
    Erase sFixWN
End Sub

'�j������Dictionary�֊i�[
Private Sub add2Dictionary(ByVal lLastYear As Long, ByRef uFixMD() As FixMD, ByRef uFixWN() As FixWN)
    Dim lInitializedLastYear    As Long
'    Dim lBeginYear          As Long
    Dim lEndYear            As Long
    Dim dtHoliday           As Date
    Dim lAddedDays          As Long
    Dim dtAdded()           As Date
    Dim existsHoliday       As Boolean
    Dim lYear               As Long
    Dim i                   As Long

'    '�������ς݂̍ŏI�N���擾
'    lInitializedLastYear = InitializedLastYear
'
'    If lInitializedLastYear < Year(BEGIN_DATE) Then
'        '�{�H�N���O�Ȃ�΁A�{�H�N���J�n�N�Ƃ���
'        lBeginYear = Year(BEGIN_DATE)
'    Else
'        '�{�H�N�Ȍ�Ȃ�A�������ς݂̗��N���J�n�N�Ƃ���
'        lBeginYear = lInitializedLastYear + 1
'    End If
'    lBeginYear = YEAR_MIN
'    lEndYear = lLastYear

    For lYear = This.YearMin To This.YearMax
        '�N�Ԃ̏j���i�[�p�z��N���A
        lAddedDays = 0
        ReDim dtAdded(lAddedDays)

        '�����Œ�̏j��
        For i = 0 To UBound(uFixMD)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If uFixMD(i).BeginYear <= lYear And uFixMD(i).EndYear >= lYear Then
                dtHoliday = CDate(CStr(lYear) & "/" & uFixMD(i).MD)

                dicHoliday_.Add dtHoliday, uFixMD(i).Name

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '���T�j���Œ�̏j��
        For i = 0 To UBound(uFixWN)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If uFixWN(i).BeginYear <= lYear And uFixWN(i).EndYear >= lYear Then
                dtHoliday = getNthWeeksDayOfWeek(lYear, uFixWN(i).Month, uFixWN(i).NthWeek, uFixWN(i).DayOfWeek)

                dicHoliday_.Add dtHoliday, uFixWN(i).Name

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '�t���̓�
        dtHoliday = getVernalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "�t���̓�"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '�H���̓�
        dtHoliday = getAutumnalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "�H���̓�"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '�U�֋x��
        For i = 0 To lAddedDays - 1
            existsHoliday = existsSubstituteHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "�U�֋x��"
            End If
        Next i

        '�����̋x��
        For i = 0 To lAddedDays - 1
            existsHoliday = existsNationalHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "�����̋x��"
            End If
        Next i

        Erase dtAdded
    Next lYear

End Sub

'�U�֋x���̗L��
'�@�j���idtDate�j�ɑ΂���U�֋x���̗L���i����ꍇ�́AdtSubstituteHoliday�ɑ�������j
Private Function existsSubstituteHoliday(ByVal dtDate As Date, ByRef dtSubstituteHoliday As Date) As Boolean

    Dim dtNextDay   As Date

    existsSubstituteHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDate���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If dtDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And dtDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(dtDate) = vbSunday Then
            '�j�������j���ł���΁A�����i���j���j���U�֋x��
            dtSubstituteHoliday = DateAdd("d", 1, dtDate)

            existsSubstituteHoliday = True
        End If
    ElseIf dtDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '�u�����̏j���v�����j���ɓ�����Ƃ��́A���̓���ɂ����Ă��̓��ɍł��߂��u�����̏j���v�łȂ������x���Ƃ���
        If Weekday(dtDate) = vbSunday Then
            dtNextDay = DateAdd("d", 1, dtDate)

            '���߂̏j���łȂ������擾
            Do Until dicHoliday_.Exists(dtNextDay) = False
                dtNextDay = DateAdd("d", 1, dtNextDay)
            Loop

            dtSubstituteHoliday = dtNextDay

            existsSubstituteHoliday = True
        End If
    End If

End Function

'�����̋x���̗L��
'�@�j���idtDate�j�ɑ΂������̋x���̗L���i����ꍇ�́AdtNationalHoliday�ɑ�������j
Private Function existsNationalHoliday(ByVal dtDate As Date, ByRef dtNationalHoliday As Date) As Boolean

    Dim dtBaseDay   As Date
    Dim dtNextDay   As Date

    existsNationalHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDate���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If dtDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        dtBaseDay = DateAdd("d", 1, dtDate)

        '���߂̏j���łȂ������擾
        Do Until dicHoliday_.Exists(dtBaseDay) = False
            dtBaseDay = DateAdd("d", 1, dtBaseDay)
        Loop

        '���j���ł���ΑΏۊO
        If Weekday(dtBaseDay) <> vbSunday Then
            dtNextDay = DateAdd("d", 1, dtBaseDay)

            '�������j���ł���ΑΏ�
            If dicHoliday_.Exists(dtNextDay) = True Then
                existsNationalHoliday = True

                dtNationalHoliday = dtBaseDay
            End If
        End If
    End If

End Function

'//////////////////////////////////////////////////
'���̑�N W�j���̓������擾
'//////////////////////////////////////////////////
Private Function getNthWeeksDayOfWeek(ByVal lYear As Long, _
                                      ByVal lMonth As Long, _
                                      ByVal lNth As Long, _
                                      ByVal lDayOfWeek As VbDayOfWeek) As Date

    Dim dt1stDate       As Date
    Dim lDayOfWeek1st   As Long
    Dim lOffset         As Long

    '�w��N���̂P�����擾
    dt1stDate = DateSerial(lYear, lMonth, 1)

    '�P���̗j�����擾
    lDayOfWeek1st = Weekday(dt1stDate)

    '�w����ւ̃I�t�Z�b�g���擾
    lOffset = lDayOfWeek - lDayOfWeek1st

    If lDayOfWeek1st > lDayOfWeek Then
        lOffset = lOffset + 7
    End If

    lOffset = lOffset + 7 * (lNth - 1)

    getNthWeeksDayOfWeek = DateAdd("d", lOffset, dt1stDate)

End Function

'//////////////////////////////////////////////////
'�t���̓����擾
'//////////////////////////////////////////////////
Private Function getVernalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(20.8431 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getVernalEquinoxDay = DateSerial(lYear, 3, lDay)

End Function

'//////////////////////////////////////////////////
'�H���̓����擾
'//////////////////////////////////////////////////
Private Function getAutumnalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(23.2488 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getAutumnalEquinoxDay = DateSerial(lYear, 9, lDay)

End Function

Private Sub qSort(ByRef dtHolidays() As Date, ByVal lLeft As Long, ByVal lRight As Long)

    Dim dtCenter    As Date
    Dim dtTemp      As Date
    Dim i           As Long
    Dim j           As Long

    If lLeft < lRight Then
        dtCenter = dtHolidays((lLeft + lRight) \ 2)

        i = lLeft - 1
        j = lRight + 1

        Do While (True)
            i = i + 1
            Do While (dtHolidays(i) < dtCenter)
                i = i + 1
            Loop

            j = j - 1
            Do While (dtHolidays(j) > dtCenter)
                j = j - 1
            Loop

            If i >= j Then
                Exit Do
            End If

            dtTemp = dtHolidays(i)
            dtHolidays(i) = dtHolidays(j)
            dtHolidays(j) = dtTemp
        Loop

        Call qSort(dtHolidays, lLeft, i - 1)
        Call qSort(dtHolidays, j + 1, lRight)
    End If

End Sub

