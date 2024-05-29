VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   12790
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   19780
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlEvent

Private searchCriteriaName As String
Private searchCriteriaAge As Long
Private searchCriteriaAddress As String
Private searchCriteriaSex As String
Private searchCriteriaBloodType As String
Private searchCriteriaDate As Date
Private searchCriteriaDateLevel As Long

Private originalTable As ListObject
Private workTable As ListObject
Private resultSheet As Worksheet
Private criteriaSheet As Worksheet

Private onFocusListBox As Boolean
Private onFocusComboBox As Boolean
Private onFocusFrame As Boolean

Private FrameTargetFullHeight As Single
Private FrameStationFullHeight As Single
Private FrameTargetOpen As Boolean
Private FrameStationOpen As Boolean

Private ListBoxFullWidth As Single
Private ListBoxFullHeight As Single

Private startTime As Single

Private Enum FormMode
    fmNewItem = 0
    fmEdit = 1
End Enum

Private Type TState
    Control As ControlEvents
    PrevControl As MSForms.IControl
    Mode As FormMode
End Type

Private This As TState

'���F��:approver�@���F����:Approve�@���F:Approval
'����:signature
'����:constraint
'OperationProcedure
'Reason for operation
'Operation results
'TimeUnit
'�F��:authentication

Private Sub CreateCheckBox()
    Dim TargetList() As Variant
    TargetList = Sheet3.ListObjects("T_�i��").DataBodyRange
       
    Dim MarginY As Long: MarginY = 2
    Dim posX As Long, posY As Long: posX = 6: posY = 30
    Dim group As Long: group = 1
    Dim i As Long
    
    With Me.FrameTarget
        For i = 1 To UBound(TargetList, 1)
            If group <> TargetList(i, 2) Then
                group = TargetList(i, 2)
                posY = posY + MarginY
                With .Controls.Add("Forms.Label.1", "LabelTargetSeparator" & group - 1)
                    .SpecialEffect = fmButtonEffectFlat
                    .BorderStyle = fmBorderStyleSingle
                    .Top = posY
                    .Left = posX
                    .Height = 1
                    .Width = Me.FrameTarget.InsideWidth - posX * 2
                End With
                posY = posY + MarginY
            End If
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxTarget" & i)
                .Caption = TargetList(i, 1)
                .GroupName = group
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = rgbWhite
                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .Height + MarginY
            End With
        Next i
'        .Width = FrameWidth
'        .Height = FrameHeight
        FrameTargetFullHeight = posY + 12
    End With
    
    TargetList = Sheet3.ListObjects("T_��").DataBodyRange
    group = 1
    posY = 30
    With Me.FrameStation
        For i = 1 To UBound(TargetList, 1)
            If group <> TargetList(i, 2) Then
                group = TargetList(i, 2)
                posY = posY + MarginY
                With .Controls.Add("Forms.Label.1", "LabelStationSeparator" & group - 1)
                    .SpecialEffect = fmButtonEffectFlat
                    .BorderStyle = fmBorderStyleSingle
                    .Top = posY
                    .Left = posX
                    .Height = 1
                    .Width = Me.FrameStation.InsideWidth - posX * 2
                End With
                posY = posY + MarginY
            End If
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxStation" & i)
                .Caption = TargetList(i, 1)
                .GroupName = group
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = rgbWhite
                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .Height + MarginY
            End With
        Next i
'        .Left = Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewTarget.Left + Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewTarget.Width + 4
'        .Top = Me.FrameTarget.Top + Me.FrameTarget.Height + 4
'        .Width = FrameWidth
        FrameStationFullHeight = posY + 12
    End With
End Sub

Private Sub UserForm_Initialize()
    startTime = Timer
    searchCriteriaAge = -1
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_�s���{��").ListColumns("�s���{����").DataBodyRange
    Dim i As Long
    With ComboBoxEditAddress
        For i = 1 To listRange.Rows.count
            .AddItem listRange(i)
        Next
    End With
    Call CreateCheckBox
    
    Me.Top = 298
    Me.Left = 430.5
    
    With Me.FrameFilter
        .ScrollBars = fmScrollBarsNone
    End With
    
    Call AddTemporarySheet
    Call AutoFitListbox
    Debug.Print ListBoxFullWidth

    Set This.Control = New ControlEvents
    With This.Control
        .parent = Me
        .Init
    End With
    Debug.Print Timer - startTime
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not This.PrevControl Is Nothing Then
'        This.PrevControl.Object.BorderStyle = fmBorderStyleNone
        Select Case This.PrevControl.Tag
            Case "Button"
                This.PrevControl.Object.BackStyle = fmBackStyleTransparent
            Case Else
        End Select
    End If
'    Me.LabelEditDatePicker.BackStyle = fmBackStyleTransparent
'    Me.LabelNewDatePicker.BackStyle = fmBackStyleTransparent
    UnHook
    onFocusListBox = False
    onFocusComboBox = False
    onFocusFrame = False
    Set This.PrevControl = Nothing
End Sub

Private Sub UserForm_Terminate()
    Set originalTable = Nothing
'    Set workTable = Nothing
'
'    With ThisWorkbook.Worksheets("FilterResult")
'        .Visible = True
'    End With
'
'    Application.DisplayAlerts = False
'    resultSheet.Delete
    Set resultSheet = Nothing
    Set criteriaSheet = Nothing
'    Application.DisplayAlerts = True
    UnHook
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

'--------------------�C���^�[�t�F�C�X����R�[���o�b�N����郁���o�֐�
Private Sub IControlEvent_OnAfterUpdate(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "TextBoxEditName"
            searchCriteriaName = Cont.Text
            Call Filter2
        Case Cont.Name = "TextBoxEditAge"
            If Cont.Text = "" Then
            searchCriteriaAge = -1
            Else
                searchCriteriaAge = Cont.Text
            End If
            Call Filter2
        Case Cont.Name = "TextBoxNewBirthDay"
            Call Filter2
        Case Cont.Name = "TextBoxEditBirthDay"
            Call Filter2
'        Case Cont.Name = "ComboBoxEditAddress"
'            Debug.Print Cont.Name & " AfterUpdate:" & Cont.Text
'            searchCriteriaAddress = Cont.Text
'            Call Filter2
    End Select
    Debug.Print Cont.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(Cont As MSForms.IControl, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case Cont.Name = "TextBoxNewBirthDay"
            If VBA.IsDate(Cont.value) Then
                searchCriteriaDate = Cont.value
                Cont.Text = Format(searchCriteriaDate, "YYYY/MM/DD")
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxEditBirthDay"
            If VBA.IsDate(Cont.value) Then
                searchCriteriaDate = Cont.value
                Cont.Text = Format(searchCriteriaDate, "YYYY/MM/DD")
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
    End Select
    Debug.Print Cont.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(Cont As MSForms.IControl)
    Select Case True
'        Case Cont.Name = "CheckBoxEditFemale"
'            If Cont.value Then searchCriteriaSex = "��"
'            Call Filter2
'        Case Cont.Name = "CheckBoxEditMale"
'            If Cont.value Then searchCriteriaSex = "�j"
'            Call Filter2
        Case Cont.Name = "ComboBoxEditAddress"
            Debug.Print Cont.Name & " Change:" & Cont.Text
            searchCriteriaAddress = Cont.Text
            Call Filter2

'        Case InStr(1, Cont.Name, "OptionButtonEditBloodType") > 0
'            searchCriteriaBloodType = Replace(Cont.Name, "OptionButtonEditBloodType", "")
'            Call Filter2
'        Case Cont.Name = "OptionButtonEditFemale"
'            If Cont.value Then searchCriteriaSex = "��"
'            Call Filter2
'        Case Cont.Name = "OptionButtonEditMale"
'            If Cont.value Then searchCriteriaSex = "�j"
'            Call Filter2
'        Case InStr(1, Cont.Name, "OptionButtonMode") > 0
'            If Cont.Name = "OptionButtonModeRegistorItem" Then
'                Me.MultiPageSwitchMode.value = 0
'            Else
'                Me.MultiPageSwitchMode.value = 1
'            End If
        Case Else
    End Select
    Debug.Print Cont.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(Cont As MSForms.IControl)
    Dim pos As POINTAPI
'    Dim dt As DatePicker
'    Set dt = New DatePicker
    Dim totalHeight As Long
    Select Case True
        Case Cont.Name = "LabelEditDatePicker"
            Cont.BackStyle = fmBackStyleOpaque
'            pos = GetControlPosition(Cont, BottomLeft)
'            Debug.Print pos.x
'            Call DatePicker.Init(pos.y, pos.x)
'            Me.TextBoxEditBirthDay.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
            DatePicker.Init Me.TextBoxEditBirthDay
        Case Cont.Name = "LabelNewDatePicker"
            Cont.BackStyle = fmBackStyleOpaque
'            pos = GetControlPosition(Cont, BottomLeft)
'            Debug.Print pos.x
'            Call DatePicker.Init
'            Me.TextBoxNewBirthDay.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
            DatePicker.Init Me.TextBoxNewBirthDay
        Case Cont.Name = "LabelSelectTarget"
            FrameTargetOpen = Not FrameTargetOpen
            If FrameTargetOpen Then
                Me.FrameTarget.Height = FrameTargetFullHeight
            Else
                Me.FrameTarget.Height = 18
            End If
            Me.FrameStation.Top = Me.FrameTarget.Top + Me.FrameTarget.Height + 6
            totalHeight = Me.FrameTarget.Height + Me.FrameStation.Height + 6 + 12
            If Me.FrameFilter.Height < totalHeight Then
                With Me.FrameFilter
                    .ScrollBars = fmScrollBarsVertical
                    .ScrollHeight = totalHeight
                    .ScrollTop = 0
                End With
                onFocusFrame = True
                HookFrame Me.FrameFilter
            Else
                With Me.FrameFilter
                    .ScrollTop = 0
                    .ScrollBars = fmScrollBarsNone
                End With
                UnHook
            End If

        Case Cont.Name = "LabelSelectStation"
            FrameStationOpen = Not FrameStationOpen
            If FrameStationOpen Then
                Me.FrameStation.Height = FrameStationFullHeight
            Else
                Me.FrameStation.Height = 18
            End If
            totalHeight = Me.FrameTarget.Height + Me.FrameStation.Height + 6 + 12
            If Me.FrameFilter.Height < totalHeight Then
                With Me.FrameFilter
                    .ScrollBars = fmScrollBarsVertical
                    .ScrollHeight = totalHeight
                    If FrameStationOpen Then
'                        .ScrollTop = Me.FrameStation.Top
                    Else
                        .ScrollTop = 0
                    End If
                End With
                onFocusFrame = True
                HookFrame Me.FrameFilter
            Else
                With Me.FrameFilter
                    .ScrollTop = 0
                    .ScrollBars = fmScrollBarsNone
                End With
                UnHook
            End If
        Case Else
            Debug.Print Cont.Name & " Click"
    End Select
'    Set dt = Nothing
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "ComboBoxEditAddress"
            Debug.Print onFocusComboBox
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            HookComboBox Cont

    End Select
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
    Select Case True
        Case Cont.Name = "TextBoxEditBirthDay"
            If KeyCode = 187 And Shift = 2 Then Cont.value = Format(Now, "YYYY/MM/DD") ' Ctrl + �u;�v

    End Select
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
    Select Case True
        Case InStr(1, Cont.Name, "OptionButtonEditBloodType") > 0
            searchCriteriaBloodType = Replace(Cont.Name, "OptionButtonEditBloodType", "")
            Call Filter2
        Case Cont.Name = "OptionButtonEditFemale"
            If Cont.value Then searchCriteriaSex = "��"
            Call Filter2
        Case Cont.Name = "OptionButtonEditMale"
            If Cont.value Then searchCriteriaSex = "�j"
            Call Filter2
        Case InStr(1, Cont.Name, "OptionButtonMode") > 0
            If Cont.Name = "OptionButtonModeRegistorItem" Then
                Me.MultiPageSwitchMode.value = 0
            Else
                Me.MultiPageSwitchMode.value = 1
            End If
        Case Else
    End Select
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
'    Select Case True
'        Case Cont.Name = "ComboBoxEditAddress"
'            If onFocusComboBox Then Exit Sub
'            onFocusComboBox = True
'            HookComboBox Cont
'        Case Cont.Name = "ListBoxEdit"
'            If onFocusListBox Then Exit Sub
'            onFocusListBox = True
'            HookListBox Cont
'        Case Cont.Name = "FrameFilter"
'            If onFocusFrame Then Exit Sub
'            If Cont.ScrollBars = fmScrollBarsNone Then UnHook: Exit Sub
'            onFocusFrame = True
'            HookFrame Cont
'        Case TypeName(Cont) = "Label"
            Call Hover(Cont)

'        Case Cont.Name = "LabelEditDatePicker"
'            Cont.BackStyle = fmBackStyleOpaque
'        Case Cont.Name = "LabelNewDatePicker"
'            Cont.BackStyle = fmBackStyleOpaque
'        Case Else
''            Me.LabelEditDatePicker.BackStyle = fmBackStyleTransparent
''            Me.LabelNewDatePicker.BackStyle = fmBackStyleTransparent
'            UnHook
'            onFocusListBox = False
'            onFocusComboBox = False
'            onFocusFrame = False
'    End Select
'    Debug.Print Cont.Name & " MouseMove:"
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
'    If Cont.Name = "FrameEditData" Then
'        Dim s As String
'        s = "ScrollLeft:" & Cont.ScrollLeft & "   "
'        s = s & "ScrollTop:" & Cont.ScrollTop
'        Me.Caption = s
'    End If
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

'-------------------------------------------------------------------------------
'�e�R���g���[����Tag�v���p�e�B�ɉ�����MouseHover���̏������K�肷��
'�����ΏۂƂ���R���g���[���ɂ̓R�[�h���܂���Form�f�U�C������Tag�v���p�e�B��ݒ肵�Ă�������
Private Sub Hover(Cont As MSForms.IControl)
    'MouseOver
    Select Case True
        Case Cont.Name = "ComboBoxEditAddress"
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            HookComboBox Cont
        Case Cont.Name = "ListBoxEdit"
            If onFocusListBox Then Exit Sub
            onFocusListBox = True
            HookListBox Cont
        Case Cont.Name = "FrameFilter"
            If onFocusFrame Then Exit Sub
            If Cont.ScrollBars = fmScrollBarsNone Then UnHook: Exit Sub
            onFocusFrame = True
            HookFrame Cont
        Case Else
            UnHook
            onFocusListBox = False
            onFocusComboBox = False
            onFocusFrame = False
    End Select
    Select Case Cont.Tag
'        Case "SelectDay", "SelectYear", "SelectMonth"
'            Cont.Object.BorderStyle = fmBorderStyleSingle
        Case "Button"
            Cont.Object.BackStyle = fmBackStyleOpaque
            Cont.Object.BorderStyle = fmBorderStyleSingle
        Case Else
    End Select
    
    'MouseOut
    If Not This.PrevControl Is Nothing Then
        If Not This.PrevControl Is Cont Then
'            This.PrevControl.Object.BorderStyle = fmBorderStyleNone
            Select Case This.PrevControl.Tag
                Case "Button"
                    This.PrevControl.Object.BackStyle = fmBackStyleTransparent
            End Select
        End If
    End If
    
    Set This.PrevControl = Cont
End Sub

Private Sub AddTemporarySheet()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Dummy")
        Set originalTable = .ListObjects("T_Dummy")
        originalTable.ShowAutoFilter = False
        .Activate
        If Util.ExistsWorksheet("FilterResult") Then ' ThisWorkbook.Worksheets("FilterResult").Delete
            Set resultSheet = ThisWorkbook.Worksheets("FilterResult")
            resultSheet.Cells.Clear
        Else
            Set resultSheet = Sheets.Add
            resultSheet.Name = "FilterResult"
        End If
        If Util.ExistsWorksheet("Criteria") Then ' ThisWorkbook.Worksheets("Criteria").Delete
            Set criteriaSheet = ThisWorkbook.Worksheets("Criteria")
            criteriaSheet.Cells.Clear
        Else
            Set criteriaSheet = Sheets.Add
            criteriaSheet.Name = "Criteria"
        End If
        With resultSheet
            originalTable.HeaderRowRange.Copy .Range("A1")
'            originalTable.Range.Copy .Range("A1")
            With .Range("A1").CurrentRegion.Font
                .Name = Me.ListBoxEdit.Font.Name
                .size = Me.ListBoxEdit.Font.size
            End With
'            .Visible = False
        End With
        With criteriaSheet
            originalTable.HeaderRowRange.Copy .Range("A1")
'            .Visible = False
        End With
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub Filter()
    Application.ScreenUpdating = False
'    ThisWorkbook.Worksheets("Dummy").Activate
'    workTable.DataBodyRange.Delete
    With originalTable
        If Me.TextBoxName.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxAddress.value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.CheckBoxFemale.value Or Me.CheckBoxMale Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxDate.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY�NMM��DD��"), VisibleDropDown:=False
    
        Dim CellsCnt As Long    '���i�荞���ް��ٌ̾�
        Dim ColCnt As Long      '��ð��ق̗�
        Dim buf1 As Variant    '���e�[�u���S�̂̃f�[�^
'        buf1 = .Range.SpecialCells(xlCellTypeVisible)
        buf1 = .Range
'        CellsCnt = .DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).count
        ColCnt = UBound(buf1, 2)
'
        Dim buf2 As Variant    '���߂�l�Ƃ���ꎞ�I�Ȕz��
        ReDim buf2(1 To (CellsCnt / ColCnt) - 1, 1 To ColCnt)

        Dim i As Long            '�������ϐ��i�z��̍s�ʒu�j
        Dim j As Long            '�������ϐ��i�z��̗�ʒu�j
        Dim k As Long            '�e�[�u���̃f�[�^�s�{�^�C�g���s�̍s��
        For k = 2 To UBound(buf1, 1)
          If .Range.Rows(k).Hidden = False Then
            i = i + 1
            For j = 1 To ColCnt
              buf2(i, j) = buf1(k, j)
            Next j
          End If
        Next k
               
        '�I�[�g�t�B���^������
        .Range.AutoFilter
        .ShowAutoFilter = False
    End With
    With workTable
        .DataBodyRange.Delete
        .Range(2, 1).Resize(i, j) = buf2
'        .Range(2, 1).Resize(UBound(buf2, 1), UBound(buf2, 2)) = buf2
    End With
    Erase buf1
    Erase buf2
    ThisWorkbook.Worksheets("FilterResult").Activate
    Me.ListBoxEdit.RowSource = workTable.DataBodyRange.Address
    Application.ScreenUpdating = True
End Sub

Private Sub Filter2()
    startTime = Timer
    Application.ScreenUpdating = False
'    ThisWorkbook.Worksheets("Dummy").Activate
'    workTable.DataBodyRange.Delete
    With originalTable
        .ShowAutoFilter = False
        If Me.TextBoxEditName.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxEditAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxEditAddress.value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.OptionButtonEditFemale.value Or Me.OptionButtonEditMale Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonEditBloodTypeA Or Me.OptionButtonEditBloodTypeB Or Me.OptionButtonEditBloodTypeAB Or Me.OptionButtonEditBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxEditBirthDay.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY/MM/DD"), VisibleDropDown:=False
        
        resultSheet.Cells.Clear
        Dim CellsCnt As Long
        
        If .ListColumns(1).Range.SpecialCells(xlCellTypeVisible).count = 1 Then
            CellsCnt = 1
        Else
            CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
        End If
        .HeaderRowRange.Copy resultSheet.Range("A1")
'        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy resultSheet.Range("A3")
        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy resultSheet.Range("A2")
        .ShowAutoFilter = False
        With resultSheet.Range("A1").CurrentRegion.Font
            .Name = Me.ListBoxEdit.Font.Name
            .size = Me.ListBoxEdit.Font.size
        End With
'        With resultSheet.Range("A3").CurrentRegion.Font
'            .Name = Me.ListBoxEdit.Font.Name
'            .size = Me.ListBoxEdit.Font.size
'        End With
        Call AutoFitListbox

'        Me.ListBoxEdit.RowSource = resultSheet.Name & "!" & resultSheet.Range("A3").Resize(CellsCnt, .ListColumns.count).Address
    
    End With
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

Private Sub Filter3()
    startTime = Timer
    Application.ScreenUpdating = False
'    ThisWorkbook.Worksheets("Dummy").Activate
'    workTable.DataBodyRange.Delete
    With originalTable
        .ShowAutoFilter = False
        If Me.TextBoxEditName.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxEditAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxEditAddress.value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.OptionButtonEditFemale.value Or Me.OptionButtonEditMale Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonEditBloodTypeA Or Me.OptionButtonEditBloodTypeB Or Me.OptionButtonEditBloodTypeAB Or Me.OptionButtonEditBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxEditBirthDay.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY/MM/DD"), VisibleDropDown:=False
        
        resultSheet.Cells.Clear
        Dim CellsCnt As Long
        
        If .ListColumns(1).Range.SpecialCells(xlCellTypeVisible).count = 1 Then
            CellsCnt = 1
        Else
            CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
        End If
        .Range.SpecialCells(xlCellTypeVisible).Copy resultSheet.Range("A1")
        .ShowAutoFilter = False
        With resultSheet.Range("A1").CurrentRegion.Font
            .Name = Me.ListBoxEdit.Font.Name
            .size = Me.ListBoxEdit.Font.size
        End With
        Call AutoFitListbox

        Me.ListBoxEdit.RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(CellsCnt, .ListColumns.count).Address
    
    End With
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

Private Sub AutoFitListbox()
    resultSheet.Cells.EntireColumn.AutoFit
    ListBoxFullWidth = 0
    On Error GoTo ERROR_HANDLER:
    With Me.ListBoxEdit
        .ColumnHeads = True
        Dim maxColumn As Long
        maxColumn = .ColumnCount
        Dim cellWidth As Long
        .ColumnWidths = ""
        Dim i As Long
        For i = 1 To maxColumn - 1
            cellWidth = resultSheet.Cells(1, i).Width + 6
            .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & cellWidth
            Debug.Print cellWidth
            ListBoxFullWidth = ListBoxFullWidth + cellWidth
        Next i
        .RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(resultSheet.Range("A2").CurrentRegion.Rows.count, originalTable.ListColumns.count).Address
    End With
    
    Exit Sub
ERROR_HANDLER:
    MsgBox "�񕝒����G���[:" & ERR.Number & " : " & ERR.Description
End Sub
