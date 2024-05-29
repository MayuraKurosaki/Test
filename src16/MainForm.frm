VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   13440
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   21380
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

Private searchCriteriaRegNo As String
Private searchCriteriaBtNo As String
Private searchCriteriaName As String
Private searchCriteriaAge As String
Private searchCriteriaDate As String
Private searchCriteriaSex As String
Private searchCriteriaBloodType As String
Private searchCriteriaAddress As String

Private ListBoxHeaderText() As String

Private originalTable As ListObject
Private workTable As ListObject
Private resultSheet As Worksheet
Private criteriaSheet As Worksheet
Private CriteriaItemTable As ListObject
Private CriteriaRange As Range

Private onFocusListBox As Boolean
Private onFocusComboBox As Boolean
Private onFocusFrame As Boolean
Private isShowToolTip As Boolean

Private FrameFilterFullHeight As Single
Private FrameTargetFullHeight As Single
Private FrameStationFullHeight As Single
Private FrameSexFullHeight As Single
Private FrameBloodTypeFullHeight As Single
Private FrameBirthDayFullHeight As Single
Private FrameAddressFullHeight As Single

Private Const BaseTextColor As Long = &HD3D3D3
Private Const BaseBackColor As Long = &H202020
Private Const BaseBorderColor As Long = &H808080
Private Const MouseOverColor As Long = &H808080
Private Const FrameBaseColor As Long = &H404040
Private Const FrameOpenColor As Long = &H606060
Private Const TextBoxBaseBackColor As Long = &H808080
Private Const TextBoxFocusBackColor As Long = &HC0C0C0

Private startTime As Single
Private toolTipDelayTime As Double

Private Enum FormMode
    fmNewItem = 0
    fmEdit = 1
End Enum

Private Enum Sex
    sMale = 1
    sFemale = 2
    sUnknown = 4
End Enum

Private Enum BloodType
    btA = 1
    btB = 2
    btAB = 4
    btO = 8
End Enum

Private Type Field
    Controls As ControlEvents
    PrevControl As MSForms.IControl
    Mode As FormMode
End Type

Private This As Field

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
    TargetList = SheetList.ListObjects("T_�q��").DataBodyRange
       
    Dim MarginY As Long: MarginY = 2    '���ڊԂ̍���
    Dim posX As Long, posY As Long: posX = 6: posY = 30 '���ڂ̈ʒu
    Dim group As Long: group = 1    '���ڂ̃O���[�v(�ύX���������ꍇ�ɃZ�p���[�^������)
    Dim i As Long
    
    With Me.FrameTarget
        .BackColor = FrameBaseColor
        .BorderColor = BaseBorderColor
        With Me.LabelSelectTarget
            .BackColor = FrameBaseColor
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
        End With
        For i = 1 To UBound(TargetList, 1)
            If group <> TargetList(i, 2) Then
                group = TargetList(i, 2)
                posY = posY + MarginY
                With .Controls.Add("Forms.Label.1", "LabelTargetSeparator" & group - 1)
                    .SpecialEffect = fmButtonEffectFlat
                    .BorderStyle = fmBorderStyleSingle
                    .BackColor = FrameBaseColor
                    .BorderColor = BaseBorderColor
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
                .BackColor = FrameBaseColor
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
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
    
    TargetList = SheetList.ListObjects("T_�^�p��").DataBodyRange
    group = 1
    posY = 30
    With Me.FrameStation
        .BackColor = FrameBaseColor
        .BorderColor = BaseBorderColor
        With Me.LabelSelectTarget
            .BackColor = FrameBaseColor
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
        End With
        For i = 1 To UBound(TargetList, 1)
            If group <> TargetList(i, 2) Then
                group = TargetList(i, 2)
                posY = posY + MarginY
                With .Controls.Add("Forms.Label.1", "LabelStationSeparator" & group - 1)
                    .SpecialEffect = fmButtonEffectFlat
                    .BorderStyle = fmBorderStyleSingle
                    .BackColor = FrameBaseColor
                    .BorderColor = BaseBorderColor
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
                .BackColor = FrameBaseColor
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
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
    
    posY = 30
    With Me.FrameSex
        .BackColor = FrameBaseColor
        .BorderColor = BaseBorderColor
        With Me.LabelSelectTarget
            .BackColor = FrameBaseColor
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
        End With
        For i = 1 To 3
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxSex" & i)
                .Caption = Choose(i, "�j", "��", "���̑��E�s��")
'                .GroupName = group
                .SpecialEffect = fmButtonEffectFlat
                .BackColor = FrameBaseColor
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .Height + MarginY
            End With
        Next i
        FrameSexFullHeight = posY + 12
    End With
    
    posY = 30
    With Me.FrameBloodType
        .BackColor = FrameBaseColor
        .BorderColor = BaseBorderColor
        With Me.LabelSelectTarget
            .BackColor = FrameBaseColor
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
        End With
        For i = 1 To 4
            With .Controls.Add("Forms.CheckBox.1", "CheckBloodType" & i)
                .Caption = Choose(i, "A", "B", "AB", "O")
'                .GroupName = group
                .SpecialEffect = fmButtonEffectFlat
                .BackColor = FrameBaseColor
                .Width = 40
                .Height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .Height + MarginY
            End With
        Next i
        FrameBloodTypeFullHeight = posY + 12
    End With
End Sub

Private Sub WebBrowserInitialize()
    With Me.WebBrowserPreview
        .Silent = True
        .Navigate ("about:blank")
        DoEvents
        
        With .Document.Body.Style
            .backgroundColor = "#202020"
            .Color = "#D3D3D3"
'        .Document.Body.Style.FontStyle = "bold"
            .FontSize = "x-large"
            .FontFamily = "Yu Gothic UI"
        End With
        .Document.Body.Innerhtml = "<div style=""Height:700px;display:flex;justify-content:center;align-items:center;""><p>PDF�t�@�C���������Ƀh���b�v���Ă�������</p></div>"
'        .Document.Body.Innerhtml = "<p style=""color:White;text-align:center;"">PDF�������Ƀh���b�v���Ă�������</p>"
    End With
End Sub

Private Sub CommandButton1_Click()
    WebBrowserInitialize
End Sub

Private Sub UserForm_Initialize()
    startTime = Timer
    searchCriteriaAge = ""
    toolTipDelayTime = 0.3
    
    Dim listRange As Range
'    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_�s���{��").ListColumns("�s���{����").DataBodyRange
'    Dim i As Long
'    With ComboBoxAddress
'        .ForeColor = BaseTextColor
'        .BackColor = TextBoxBaseBackColor
'        .BorderColor = BaseBorderColor
'        For i = 1 To listRange.Rows.count
'            .AddItem listRange(i)
'        Next
'    End With
    Call CreateCheckBox
    
    Me.Top = 298
    Me.Left = 430.5
    Me.BackColor = BaseBackColor
    Me.BorderColor = BaseBorderColor
    
    With Me.FrameFilter
        .BackColor = BaseBackColor
        .BorderColor = BaseBorderColor
        .ScrollBars = fmScrollBarsNone
    End With
    
    With Me.LabelToolTip
'        .AutoSize = True
        .Font.size = 9
        .TextAlign = fmTextAlignLeft
    End With
    
    Call AddTemporarySheet
    Call AutoFitListbox
'    Debug.Print ListBoxFullWidth

    Call WebBrowserInitialize
'    With Me.WebBrowserPreview
'        .Silent = True
'        .Navigate ("about:blank")
'        DoEvents
'
'        With .Document.Body.Style
'            .backgroundColor = "#202020"
'            .Color = "#D3D3D3"
''        .Document.Body.Style.FontStyle = "bold"
'            .FontSize = "x-large"
'            .FontFamily = "Yu Gothic UI"
'        End With
'        .Document.Body.Innerhtml = "<div style=""Height:700px;display:flex;justify-content:center;align-items:center;""><p>PDF�t�@�C���������Ƀh���b�v���Ă�������</p></div>"
''        .Document.Body.Innerhtml = "<p style=""color:White;text-align:center;"">PDF�������Ƀh���b�v���Ă�������</p>"
'    End With

    Set This.Controls = New ControlEvents
    With This.Controls
        .Parent = Me
        .Init
    End With
    
    Call SetControlAttribute
    
    Debug.Print Timer - startTime
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not This.PrevControl Is Nothing Then
'        This.PrevControl.Object.BorderStyle = fmBorderStyleNone
        Select Case This.PrevControl.Tag
            Case "Button"
                This.PrevControl.Object.BackStyle = fmBackStyleOpaque
                This.PrevControl.Object.BorderStyle = fmBorderStyleNone
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

'--------------------�C���^�[�t�F�C�X����R�[���o�b�N����郁���o�֐�
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
        Case CtrlEvt.Control.Name = "TextBoxAgeFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxAgeTo"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxBirthDayFrom"
            Call Filter3
        Case CtrlEvt.Control.Name = "TextBoxBirthDayTo"
            Call Filter3
'        Case CtrlEvt.Control.Name = "ComboBoxEditAddress"
'            Debug.Print CtrlEvt.Control.Name & " AfterUpdate:" & CtrlEvt.Control.Text
'            searchCriteriaAddress = CtrlEvt.Control.Text
'            Call Filter2
    End Select
    Debug.Print CtrlEvt.Control.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case CtrlEvt.Control.Name = "TextBoxNewBirthDay"
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
        Case CtrlEvt.Control.Name = "TextBoxBloodTypeNoFrom"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                 If Me.OptionButtonBloodTypeNoSingle Then
                    searchCriteriaBtNo = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaBtNo = CtrlEvt.Control.value & "," & TextBoxBloodTypeNoTo.value
                End If
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxBloodTypeNoTo"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                searchCriteriaBtNo = TextBoxBloodTypeNoFrom.value & "," & CtrlEvt.Control.value
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxAgeFrom"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                 If Me.OptionButtonAgeSingle Then
                    searchCriteriaAge = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaAge = CtrlEvt.Control.value & "," & TextBoxAgeTo.value
                End If
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxAgeTo"
            If CtrlEvt.Control.Text = "" Or VBA.IsNumeric(CtrlEvt.Control.value) Then
                searchCriteriaAge = TextBoxAgeFrom.value & "," & CtrlEvt.Control.value
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxBirthDayFrom"
            If VBA.IsDate(CtrlEvt.Control.value) Then
                If Me.OptionButtonSingleDay Then
                    searchCriteriaDate = CtrlEvt.Control.value & "," & CtrlEvt.Control.value
                Else
                    searchCriteriaDate = CtrlEvt.Control.value & "," & TextBoxBirthDayTo.value
                End If
                CtrlEvt.Control.Text = Format(CtrlEvt.Control.value, "YYYY/MM/DD")
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
        Case CtrlEvt.Control.Name = "TextBoxBirthDayTo"
            If VBA.IsDate(CtrlEvt.Control.value) Then
                searchCriteriaDate = TextBoxBirthDayFrom.value & "," & CtrlEvt.Control.value
                CtrlEvt.Control.Text = Format(CtrlEvt.Control.value, "YYYY/MM/DD")
            Else
                If CtrlEvt.Control.Text <> "" Then
                    CtrlEvt.Control.SelStart = 0
                    CtrlEvt.Control.SelLength = VBA.Len(CtrlEvt.Control.Text)
                    Cancel = True
                End If
            End If
    End Select
    Debug.Print CtrlEvt.Control.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)
    Dim i As Long
    Select Case True
'        Case CtrlEvt.Control.Name = "CheckBoxEditFemale"
'            If CtrlEvt.Control.value Then searchCriteriaSex = "��"
'            Call Filter2
'        Case CtrlEvt.Control.Name = "CheckBoxEditMale"
'            If CtrlEvt.Control.value Then searchCriteriaSex = "�j"
'            Call Filter2
        Case CtrlEvt.Control.Name = "ComboBoxAddress"
            Debug.Print CtrlEvt.Control.Name & " Change:" & CtrlEvt.Control.Text
            searchCriteriaAddress = CtrlEvt.Control.Text
            Call Filter3
        Case Left(CtrlEvt.Control.Name, 14) = "CheckBloodType"
            For i = 1 To 4
                With Me.Controls("CheckBloodType" & i)
                    If .value Then
                        searchCriteriaBloodType = searchCriteriaBloodType & "," & .Caption
                    End If
                End With
            Next i
            searchCriteriaBloodType = Right(searchCriteriaBloodType, Len(searchCriteriaBloodType) - 1)
        Case Left(CtrlEvt.Control.Name, 11) = "CheckBoxSex"
            For i = 1 To 3
                With Me.Controls("CheckBoxSex" & i)
                    If .value Then
                        searchCriteriaSex = searchCriteriaSex & "," & .Caption
                    End If
                End With
            Next i
            searchCriteriaSex = Right(searchCriteriaSex, Len(searchCriteriaSex) - 1)

'        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonEditBloodType") > 0
'            searchCriteriaBloodType = Replace(CtrlEvt.Control.Name, "OptionButtonEditBloodType", "")
'            Call Filter2
'        Case CtrlEvt.Control.Name = "OptionButtonEditFemale"
'            If CtrlEvt.Control.value Then searchCriteriaSex = "��"
'            Call Filter2
'        Case CtrlEvt.Control.Name = "OptionButtonEditMale"
'            If CtrlEvt.Control.value Then searchCriteriaSex = "�j"
'            Call Filter2
'        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonMode") > 0
'            If CtrlEvt.Control.Name = "OptionButtonModeRegistorItem" Then
'                Me.MultiPageSwitchMode.value = 0
'            Else
'                Me.MultiPageSwitchMode.value = 1
'            End If
        Case Else
    End Select
    Debug.Print CtrlEvt.Control.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)
    Select Case True
'        Case CtrlEvt.Control.Name = "LabelDatePickerEdit"
'            CtrlEvt.Control.BackStyle = fmBackStyleOpaque
'            DatePicker.Init Me.TextBoxBirthDayEdit
'        Case CtrlEvt.Control.Name = "LabelDatePickerNew"
'            CtrlEvt.Control.BackStyle = fmBackStyleOpaque
'            DatePicker.Init Me.TextBoxBirthDayNew
        Case Left(CtrlEvt.Control.Name, 15) = "LabelDatePicker"
            Call OpenDatePicker(CtrlEvt)
        Case CtrlEvt.Control.Tag = "SideBar"
            Call OpenSideBar(CtrlEvt)
        Case Else
            Debug.Print CtrlEvt.Control.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Call IControlEvent_OnClick(CtrlEvt)
    DoEvents
    Cancel = True
    Debug.Print CtrlEvt.Control.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "ComboBoxAddress"
            Debug.Print onFocusComboBox
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
'            HookComboBox CtrlEvt.Control
            HookControl CtrlEvt

    End Select
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
    Select Case True
        Case CtrlEvt.Control.Name = "TextBoxEditBirthDay"
            If KeyCode = 187 And Shift = 2 Then CtrlEvt.Control.value = Format(Now, "YYYY/MM/DD") ' Ctrl + �u;�v

    End Select
    Debug.Print CtrlEvt.Control.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
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
    Select Case True
        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonEditBloodType") > 0
            searchCriteriaBloodType = Replace(CtrlEvt.Control.Name, "OptionButtonEditBloodType", "")
            Call Filter3
        Case CtrlEvt.Control.Name = "OptionButtonEditFemale"
            If CtrlEvt.Control.value Then searchCriteriaSex = "��"
            Call Filter3
        Case CtrlEvt.Control.Name = "OptionButtonEditMale"
            If CtrlEvt.Control.value Then searchCriteriaSex = "�j"
            Call Filter3
        Case InStr(1, CtrlEvt.Control.Name, "OptionButtonMode") > 0
            If CtrlEvt.Control.Name = "OptionButtonModeRegistorItem" Then
                Me.MultiPageSwitchMode.value = 0
            Else
                Me.MultiPageSwitchMode.value = 1
            End If
        Case Else
    End Select
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
    Call Hover(CtrlEvt, Button, Shift, X, Y)
'    Debug.Print CtrlEvt.Control.Name & " MouseMove:(" & X & "," & Y & ") / Button:" & Button & " / Shift:" & Shift
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

'-------------------------------------------------------------------------------
'�eControl�ɑ�����t�^����
Private Sub SetControlAttribute()

End Sub

'�e�R���g���[����Tag�v���p�e�B�ɉ�����MouseHover���̏������K�肷��
'�����ΏۂƂ���R���g���[���ɂ̓R�[�h���܂���Form�f�U�C������Tag�v���p�e�B��ݒ肵�Ă�������
Private Sub Hover(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static delay As Double
    
    'MouseOver
    Select Case True
        Case CtrlEvt.Control.Name = "ComboBoxAddress"
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
'            HookComboBox CtrlEvt.Control
            HookControl CtrlEvt
        Case CtrlEvt.Control.Name = "ListBoxEdit"
'            Debug.Print CtrlEvt.Control.Name & " MouseMove:TopIndex:" & CtrlEvt.Control.TopIndex & " / MousePointer:(" & X & "," & Y & ") / Button:" & Button & " / Shift:" & Shift
            If onFocusListBox Then
            Else
                onFocusListBox = True
                delay = Util.GetTimer
'                HookListBox CtrlEvt.Control
                HookControl CtrlEvt
            End If
            If Util.GetTimer > delay + toolTipDelayTime Then
                Call ShowListToolTip(CtrlEvt, X, Y)
            End If

        Case TypeName(CtrlEvt.Control) = "Frame"
            If Left(CtrlEvt.Control.Tag, 13) = "SelectionField" Then
                If onFocusFrame Then Exit Sub
                If Me.FrameFilter.ScrollBars = fmScrollBarsNone Then UnHook: Exit Sub
                onFocusFrame = True
'                HookFrame Me.FrameFilter
                HookControl CtrlEvt
            End If
        Case TypeName(CtrlEvt.Control) = "Label"
            Select Case True
        '        Case "SelectDay", "SelectYear", "SelectMonth"
        '            CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
                Case CtrlEvt.Control.Tag = "Button"
'                    CtrlEvt.Control.Object.BackColor = MouseOverColor
                    CtrlEvt.Control.Object.BackStyle = fmBackStyleTransparent
'                    CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
                Case CtrlEvt.Control.Tag = "SideBar"
                    CtrlEvt.Control.Object.BackColor = MouseOverColor
                Case Else
            End Select
        Case Else
            UnHook
            onFocusListBox = False
            onFocusComboBox = False
            onFocusFrame = False
            Call CloseListToolTip
            delay = 0
    End Select
'    Select Case CtrlEvt.Control.Tag
''        Case "SelectDay", "SelectYear", "SelectMonth"
''            CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
'        Case "Button"
'            CtrlEvt.Control.Object.BackStyle = fmBackStyleOpaque
'            CtrlEvt.Control.Object.BorderStyle = fmBorderStyleSingle
'        Case Else
'    End Select
    
    'MouseOut
    If Not This.PrevControl Is Nothing Then
        If Not This.PrevControl Is CtrlEvt.Control Then
'            This.PrevControl.Object.BorderStyle = fmBorderStyleNone
            Select Case True
                Case This.PrevControl.Tag = "Button"
'                    This.PrevControl.Object.BackColor = BaseBackColor
                    This.PrevControl.Object.BackStyle = fmBackStyleOpaque
'                    This.PrevControl.Object.BorderStyle = fmBorderStyleNone
                Case This.PrevControl.Tag = "SideBar"
                    This.PrevControl.Object.BackColor = FrameBaseColor
                Case This.PrevControl.Name = "ListBoxEdit"
                    UnHook
                    onFocusListBox = False
                    Call CloseListToolTip
            End Select
        End If
    End If
    
    Set This.PrevControl = CtrlEvt.Control
End Sub

Private Sub ShowListToolTip(CtrlEvt As ControlEvent, ByVal X As Single, ByVal Y As Single)
    Dim tipListIndex As Long
    Dim tipText As String
    Dim itemHeight As Single
    
    With Me.ListBoxEdit
        itemHeight = .Font.size * 1.4
        tipListIndex = Fix(Y / itemHeight + .TopIndex)
        
        If tipListIndex > .ListCount - 1 Then tipListIndex = .ListCount - 1
        
        If .ColumnHeads Then tipListIndex = tipListIndex - 1
        If tipListIndex < 0 Then CloseListToolTip: Exit Sub
        tipText = ListBoxHeaderText(1) & ":" & .List(tipListIndex, 1) & vbLf & _
                  ListBoxHeaderText(2) & ":" & .List(tipListIndex, 2) & vbLf & _
                  ListBoxHeaderText(3) & ":" & .List(tipListIndex, 3)
    End With
    Me.LabelToolTip.Caption = tipText
    Call AutoFitControl(Me.LabelToolTip)
'    Me.LabelToolTip.AutoSize = True
    
    
    Dim tipTop As Single, tipLeft As Single
    tipTop = Me.MultiPageSwitchMode.Top + CtrlEvt.Control.Top + Y + itemHeight
    tipLeft = Me.MultiPageSwitchMode.Left + CtrlEvt.Control.Left + X
    
    With Me.FrameToolTip
        .Height = Me.LabelToolTip.Height
        .Width = Me.LabelToolTip.Width
        If tipTop + .Height > .Parent.InsideHeight Then tipTop = Me.MultiPageSwitchMode.Top + CtrlEvt.Control.Top + Y - itemHeight - .Height
        If tipLeft + .Width > .Parent.InsideWidth Then tipLeft = tipLeft - .Width
        .Top = tipTop
        .Left = tipLeft
        
        If Not isShowToolTip Then
            .Visible = True
            isShowToolTip = True
        End If
    End With
End Sub

Private Sub CloseListToolTip()
    With Me.FrameToolTip
        .Caption = ""
        .Visible = False
    End With
    
    isShowToolTip = False
End Sub

'ListBox���̎w�肵�����W�ɂ��鍀�ڂ� 0 ����n�܂�C���f�b�N�X�ԍ���Ԃ��܂��B
Private Function IndexFromPoint(ListBox As MSForms.ListBox, ByVal Y As Single) As Long
    Dim itemHeight As Single
    Dim topItem As Long
    Dim curItem As Long
    
    itemHeight = ListBox.Font.size * 1.2
    topItem = ListBox.TopIndex
    
    curItem = Fix(Y / itemHeight + topItem)
    
    If curItem > ListBox.ListCount - 1 Then curItem = ListBox.ListCount - 1
    IndexFromPoint = curItem
End Function

Private Sub OpenSideBar(CtrlEvt As ControlEvent)
    Dim Frame As MSForms.Frame
    Set Frame = CtrlEvt.Control.Parent
    
    Dim FrameFullHeight As Single
    Select Case Frame.Name
        Case "FrameRegistorNo": FrameFullHeight = 132
        Case "FrameBloodTypeNo": FrameFullHeight = 132
        Case "FrameName": FrameFullHeight = 60
        Case "FrameAge": FrameFullHeight = 132
        Case "FrameTarget": FrameFullHeight = FrameTargetFullHeight
        Case "FrameStation": FrameFullHeight = FrameStationFullHeight
        Case "FrameSex": FrameFullHeight = FrameSexFullHeight
        Case "FrameBloodType": FrameFullHeight = FrameBloodTypeFullHeight
        Case "FrameBirthDay": FrameFullHeight = 130
        Case "FrameAddress": FrameFullHeight = 54
    End Select
    
    Dim isOpen As Boolean
    isOpen = Not CBool(Right(Frame.Tag, 1))
    
    If isOpen Then
        Frame.Tag = "SelectionField1"
        Frame.Height = FrameFullHeight
        CtrlEvt.Control.Object.BackColor = FrameOpenColor
    Else
        Frame.Tag = "SelectionField0"
        Frame.Height = 18
        CtrlEvt.Control.Object.BackColor = FrameBaseColor
    End If
    
    Dim FrameMargin As Single: FrameMargin = 6
    Me.FrameBloodTypeNo.Top = Me.FrameRegistorNo.Top + Me.FrameRegistorNo.Height + FrameMargin
    Me.FrameName.Top = Me.FrameBloodTypeNo.Top + Me.FrameBloodTypeNo.Height + FrameMargin
    Me.FrameAge.Top = Me.FrameName.Top + Me.FrameName.Height + FrameMargin
    Me.FrameSex.Top = Me.FrameAge.Top + Me.FrameAge.Height + FrameMargin
    Me.FrameBloodType.Top = Me.FrameSex.Top + Me.FrameSex.Height + FrameMargin
    Me.FrameBirthDay.Top = Me.FrameBloodType.Top + Me.FrameBloodType.Height + FrameMargin
    Me.FrameAddress.Top = Me.FrameBirthDay.Top + Me.FrameBirthDay.Height + FrameMargin
    Me.FrameTarget.Top = Me.FrameAddress.Top + Me.FrameAddress.Height + FrameMargin
    Me.FrameStation.Top = Me.FrameTarget.Top + Me.FrameTarget.Height + FrameMargin
    
    FrameFilterFullHeight = Me.FrameRegistorNo.Height + Me.FrameBloodTypeNo.Height + _
                            Me.FrameName.Height + Me.FrameAge.Height + _
                            Me.FrameSex.Height + Me.FrameBloodType.Height + _
                            Me.FrameBirthDay.Height + Me.FrameAddress.Height + _
                            Me.FrameTarget.Height + Me.FrameStation.Height + FrameMargin * 7 + 12

    If Me.FrameFilter.Height < FrameFilterFullHeight Then
        With Me.FrameFilter
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = FrameFilterFullHeight
'                    .ScrollTop = 0
            If isOpen Then
                If Frame.Height > FrameFilterFullHeight Then
                    .ScrollTop = Frame.Top
                Else
                    .ScrollTop = .ScrollTop + Frame.Height
                End If
            Else
                .ScrollTop = 0
            End If
        End With
        onFocusFrame = True
        'HookFrame Me.FrameFilter
        HookControl This.Controls.Item("FrameFilter")
    Else
        With Me.FrameFilter
'                    .ScrollTop = 0
            .ScrollBars = fmScrollBarsNone
        End With
        UnHook
    End If
End Sub

Private Sub OpenDatePicker(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "LabelDatePickerEdit"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayEdit
        Case CtrlEvt.Control.Name = "LabelDatePickerNew"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayNew
        Case CtrlEvt.Control.Name = "LabelDatePickerFrom"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayFrom
        Case CtrlEvt.Control.Name = "LabelDatePickerTo"
            CtrlEvt.Control.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayTo
        Case Else
    End Select
End Sub

Private Sub AddTemporarySheet()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("�䒠")
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
'            originalTable.HeaderRowRange.Copy .Range("A1")
            originalTable.Range.Copy .Range("A1")
            With .Range("A1").CurrentRegion.Font
                .Name = Me.ListBoxEdit.Font.Name
                .size = Me.ListBoxEdit.Font.size
            End With
'            .Visible = False
        End With
        
        Set CriteriaItemTable = SheetList.ListObjects("T_���o����")
        Set CriteriaRange = criteriaSheet.Range("A1")
        
'        Dim i As Long
'        With ThisWorkbook.Worksheets("List").ListObjects("T_���o����")
'            Dim col As Long: col = 0
'            For i = 1 To .ListRows.count
'                criteriaSheet.Range("A1").Offset(0, col) = .DataBodyRange(i, 1)
'                If .DataBodyRange(i, 2) = "�͈�" Then
'                    col = col + 1
'                    criteriaSheet.Range("A1").Offset(0, col) = .DataBodyRange(i, 1)
'                End If
'                col = col + 1
'            Next i
''            originalTable.HeaderRowRange.Copy .Range("A1")
''            .Visible = False
'        End With
    End With
    Application.ScreenUpdating = True
End Sub

'Private Sub Filter()
'    Application.ScreenUpdating = False
''    ThisWorkbook.Worksheets("Dummy").Activate
''    workTable.DataBodyRange.Delete
'    With originalTable
'        If Me.TextBoxName.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
'        If Me.TextBoxAge.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
'        If Me.ComboBoxAddress.value <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
'        If Me.CheckBoxFemale.value Or Me.CheckBoxMale Then _
'            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
'        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
'            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
'        If Me.TextBoxDate.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY�NMM��DD��"), VisibleDropDown:=False
'
'        Dim CellsCnt As Long    '���i�荞���ް��ٌ̾�
'        Dim ColCnt As Long      '��ð��ق̗�
'        Dim buf1 As Variant    '���e�[�u���S�̂̃f�[�^
''        buf1 = .Range.SpecialCells(xlCellTypeVisible)
'        buf1 = .Range
''        CellsCnt = .DataBodyRange.SpecialCells(xlCellTypeVisible).Count
'        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).count
'        ColCnt = UBound(buf1, 2)
''
'        Dim buf2 As Variant    '���߂�l�Ƃ���ꎞ�I�Ȕz��
'        ReDim buf2(1 To (CellsCnt / ColCnt) - 1, 1 To ColCnt)
'
'        Dim i As Long            '�������ϐ��i�z��̍s�ʒu�j
'        Dim j As Long            '�������ϐ��i�z��̗�ʒu�j
'        Dim k As Long            '�e�[�u���̃f�[�^�s�{�^�C�g���s�̍s��
'        For k = 2 To UBound(buf1, 1)
'          If .Range.Rows(k).Hidden = False Then
'            i = i + 1
'            For j = 1 To ColCnt
'              buf2(i, j) = buf1(k, j)
'            Next j
'          End If
'        Next k
'
'        '�I�[�g�t�B���^������
'        .Range.AutoFilter
'        .ShowAutoFilter = False
'    End With
'    With workTable
'        .DataBodyRange.Delete
'        .Range(2, 1).Resize(i, j) = buf2
''        .Range(2, 1).Resize(UBound(buf2, 1), UBound(buf2, 2)) = buf2
'    End With
'    Erase buf1
'    Erase buf2
'    ThisWorkbook.Worksheets("FilterResult").Activate
'    Me.ListBoxEdit.RowSource = workTable.DataBodyRange.Address
'    Application.ScreenUpdating = True
'End Sub
'
'Private Sub Filter2()
'    startTime = Timer
'    Application.ScreenUpdating = False
''    ThisWorkbook.Worksheets("Dummy").Activate
''    workTable.DataBodyRange.Delete
'    With originalTable
'        .ShowAutoFilter = False
'        If Me.TextBoxEditName.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
'        If Me.TextBoxEditAge.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
'        If Me.ComboBoxEditAddress.value <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
'        If Me.OptionButtonEditFemale.value Or Me.OptionButtonEditMale Then _
'            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
'        If Me.OptionButtonEditBloodTypeA Or Me.OptionButtonEditBloodTypeB Or Me.OptionButtonEditBloodTypeAB Or Me.OptionButtonEditBloodTypeO Then _
'            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
'        If Me.TextBoxBirthDayEdit.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY/MM/DD"), VisibleDropDown:=False
'
'        resultSheet.Cells.Clear
'        Dim CellsCnt As Long
'
'        If .ListColumns(1).Range.SpecialCells(xlCellTypeVisible).count = 1 Then
'            CellsCnt = 1
'        Else
'            CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
'        End If
'        .HeaderRowRange.Copy resultSheet.Range("A1")
''        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy resultSheet.Range("A3")
'        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy resultSheet.Range("A2")
'        .ShowAutoFilter = False
'        With resultSheet.Range("A1").CurrentRegion.Font
'            .Name = Me.ListBoxEdit.Font.Name
'            .size = Me.ListBoxEdit.Font.size
'        End With
''        With resultSheet.Range("A3").CurrentRegion.Font
''            .Name = Me.ListBoxEdit.Font.Name
''            .size = Me.ListBoxEdit.Font.size
''        End With
'        Call AutoFitListbox
'
''        Me.ListBoxEdit.RowSource = resultSheet.Name & "!" & resultSheet.Range("A3").Resize(CellsCnt, .ListColumns.count).Address
'
'    End With
'    Application.ScreenUpdating = True
'    Debug.Print Timer - startTime
'End Sub

'Private searchCriteriaRegNo As String
'Private searchCriteriaBtNo As String
'Private searchCriteriaName As String
'Private searchCriteriaAge As Long
'Private searchCriteriaDate As Date
'Private searchCriteriaSex As String
'Private searchCriteriaBloodType As String
'Private searchCriteriaAddress As String
Private Sub SetCriteria()
    Dim criteriaType As String
    Dim criteriaItem As String
    Dim tempArr As Variant
    Dim conditions As Variant
    Dim j As Long
    Dim Item As Variant
    Dim tmp As Variant
    Debug.Print "SetCriteria"
    
    CriteriaRange.CurrentRegion.ClearContents
    
    Dim i As Long
    With CriteriaItemTable
        Dim col As Long: col = 0
        For i = 1 To .ListRows.Count
            criteriaItem = .DataBodyRange(i, 1)
            criteriaType = .DataBodyRange(i, 2)
            
            
'            Select Case criteriaType
'                Case "�͈�"
'                    CriteriaRange.Offset(0, col + 1) = criteriaItem
'                Case Else
'            End Select
            
            tmp = Choose(i, searchCriteriaRegNo, searchCriteriaBtNo, _
                                searchCriteriaName, searchCriteriaAge, searchCriteriaDate, _
                                searchCriteriaSex, searchCriteriaBloodType, searchCriteriaAddress)
            Debug.Print tmp
            If tmp = "" Then
                CriteriaRange.Offset(0, col) = criteriaItem
                If criteriaType = "�͈�" Then
                    col = col + 1
                    CriteriaRange.Offset(0, col) = criteriaItem
                End If
                GoTo CONTINUE
            End If
            tempArr = Split(tmp, ",")
            Debug.Print UBound(tempArr)
            ReDim conditions(UBound(tempArr))
            j = 0
            For Each Item In tempArr
               If Not IsEmpty(Item) Then
                   If Item <> "" Then
                       conditions(j) = Item
                       Debug.Print conditions(j)
                       j = j + 1
                   End If
               End If
            Next
            ReDim Preserve conditions(j - 1)
     
            CriteriaRange.Offset(0, col) = criteriaItem
            Select Case criteriaType
                Case "����"
                    CriteriaRange.Offset(1, col) = "*" & conditions(0) & "*"
                Case "�O��"
                    CriteriaRange.Offset(1, col) = conditions(0) & "*"
                Case "�͈�"
                    CriteriaRange.Offset(1, col) = ">=" & conditions(0)
                    col = col + 1
                    CriteriaRange.Offset(0, col) = criteriaItem
                    CriteriaRange.Offset(1, col) = "<=" & conditions(1)
                Case Else
                    For j = 0 To UBound(conditions) - 1
                        CriteriaRange.Offset(1 + j, col) = conditions(j)
                    Next j
            End Select
CONTINUE:
            col = col + 1
        Next i
    End With
    Set CriteriaRange = CriteriaRange.CurrentRegion
End Sub

Private Sub Filter3()
    startTime = Timer
    Application.ScreenUpdating = False
'    ThisWorkbook.Worksheets("Dummy").Activate
'    workTable.DataBodyRange.Delete
    Dim row As Long, col As Long
    Call SetCriteria
    originalTable.Range.AdvancedFilter xlFilterCopy, CriteriaRange, resultSheet.Range("A1").CurrentRegion
    
    Call AutoFitListbox

'    Me.ListBoxEdit.RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(resultSheet.Range("A2").CurrentRegion.Rows.count - 1, originalTable.ListColumns.count).Address
    
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

Private Sub AutoFitListbox()
    resultSheet.Cells.EntireColumn.AutoFit
    
    On Error GoTo ERROR_HANDLER:
    With Me.ListBoxEdit
        .ColumnHeads = True
        .ColumnCount = originalTable.ListColumns.Count
'        Dim maxColumn As Long
'        maxColumn = .ColumnCount
        Dim cellWidth As Long
        .ColumnWidths = ""
        Dim i As Long
        For i = 1 To .ColumnCount - 1
            cellWidth = resultSheet.Cells(1, i).Width + 6
            .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & cellWidth
        Next i
        .RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(resultSheet.Range("A2").CurrentRegion.Rows.Count - 1, originalTable.ListColumns.Count).Address
        
        Dim arrayTmp As Variant
        arrayTmp = resultSheet.Range("A1").CurrentRegion.Resize(1, .ColumnCount)
        
        Dim strArray() As String
        ReDim strArray(0 To .ColumnCount - 1)
        For i = 0 To .ColumnCount - 1
            strArray(i) = arrayTmp(1, i + 1)
        Next i
        ListBoxHeaderText = strArray
    End With
'    Debug.Print "Row:" & resultSheet.Range("A2").CurrentRegion.Rows.count & "Column:" & originalTable.ListColumns.count
    Exit Sub
ERROR_HANDLER:
    MsgBox "�񕝒����G���[:" & ERR.Number & " : " & ERR.Description
End Sub

Private Sub WebBrowserPreview_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Debug.Print URL
'    Cancel = True
End Sub
