VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   12790
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   19780
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

Private originalTable As ListObject
Private workTable As ListObject
Private resultSheet As Worksheet
Private criteriaSheet As Worksheet
Private CriteriaItemTable As ListObject
Private CriteriaRange As Range

Private onFocusListBox As Boolean
Private onFocusComboBox As Boolean
Private onFocusFrame As Boolean

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

Private Type TState
    Control As ControlEvents
    PrevControl As MSForms.IControl
    Mode As FormMode
End Type

Private This As TState

'承認者:approver　承認する:Approve　承認:Approval
'署名:signature
'制約:constraint
'OperationProcedure
'Reason for operation
'Operation results
'TimeUnit
'認証:authentication

Private Sub CreateCheckBox()
    Dim TargetList() As Variant
    TargetList = Sheet3.ListObjects("T_永世").DataBodyRange
       
    Dim MarginY As Long: MarginY = 2    '項目間の高さ
    Dim posX As Long, posY As Long: posX = 6: posY = 30 '項目の位置
    Dim group As Long: group = 1    '項目のグループ(変更があった場合にセパレータを入れる)
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
                    .height = 1
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
                .height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .height + MarginY
            End With
        Next i
'        .Width = FrameWidth
'        .Height = FrameHeight
        FrameTargetFullHeight = posY + 12
    End With
    
    TargetList = Sheet3.ListObjects("T_曲").DataBodyRange
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
                    .height = 1
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
                .height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .height + MarginY
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
                .Caption = Choose(i, "男", "女", "その他・不明")
'                .GroupName = group
                .SpecialEffect = fmButtonEffectFlat
                .BackColor = FrameBaseColor
                .Width = 40
                .height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .height + MarginY
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
                .height = 20
                .Left = posX
                .Top = posY
                .ForeColor = BaseTextColor
'                .BackColor = RGB(64, 64, 64)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                posY = posY + .height + MarginY
            End With
        Next i
        FrameBloodTypeFullHeight = posY + 12
    End With
End Sub

Private Sub UserForm_Initialize()
    startTime = Timer
    searchCriteriaAge = ""
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_都道府県").ListColumns("都道府県名").DataBodyRange
    Dim i As Long
    With ComboBoxAddress
        .ForeColor = BaseTextColor
        .BackColor = TextBoxBaseBackColor
        .BorderColor = BaseBorderColor
        For i = 1 To listRange.Rows.count
            .AddItem listRange(i)
        Next
    End With
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
    
    Call AddTemporarySheet
    Call AutoFitListbox
'    Debug.Print ListBoxFullWidth

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
        .Document.Body.Innerhtml = "<div style=""Height:700px;display:flex;justify-content:center;align-items:center;""><p>PDFファイルをここにドロップしてください</p></div>"
'        .Document.Body.Innerhtml = "<p style=""color:White;text-align:center;"">PDFをここにドロップしてください</p>"
    End With

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

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "TextBoxName"
            searchCriteriaName = Cont.Text
            Call Filter3
        Case Cont.Name = "TextBoxRegNoFrom"
            Call Filter3
        Case Cont.Name = "TextBoxRegNoTo"
            Call Filter3
        Case Cont.Name = "TextBoxBtNoFrom"
            Call Filter3
        Case Cont.Name = "TextBoxBtNoTo"
            Call Filter3
        Case Cont.Name = "TextBoxAgeFrom"
            Call Filter3
        Case Cont.Name = "TextBoxAgeTo"
            Call Filter3
        Case Cont.Name = "TextBoxBirthDayFrom"
            Call Filter3
        Case Cont.Name = "TextBoxBirthDayTo"
            Call Filter3
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
'                searchCriteriaDate = Cont.value
                Cont.Text = Format(searchCriteriaDate, "YYYY/MM/DD")
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxRegNoFrom"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                 If Me.OptionButtonRegNoSingle Then
                    searchCriteriaRegNo = Cont.value & "," & Cont.value
                Else
                    searchCriteriaRegNo = Cont.value & "," & TextBoxRegNoTo.value
                End If
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxRegNoTo"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                searchCriteriaRegNo = TextBoxRegNoFrom.value & "," & Cont.value
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxBloodTypeNoFrom"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                 If Me.OptionButtonBloodTypeNoSingle Then
                    searchCriteriaBtNo = Cont.value & "," & Cont.value
                Else
                    searchCriteriaBtNo = Cont.value & "," & TextBoxBloodTypeNoTo.value
                End If
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxBloodTypeNoTo"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                searchCriteriaBtNo = TextBoxBloodTypeNoFrom.value & "," & Cont.value
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxAgeFrom"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                 If Me.OptionButtonAgeSingle Then
                    searchCriteriaAge = Cont.value & "," & Cont.value
                Else
                    searchCriteriaAge = Cont.value & "," & TextBoxAgeTo.value
                End If
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxAgeTo"
            If Cont.Text = "" Or VBA.IsNumeric(Cont.value) Then
                searchCriteriaAge = TextBoxAgeFrom.value & "," & Cont.value
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxBirthDayFrom"
            If VBA.IsDate(Cont.value) Then
                If Me.OptionButtonSingleDay Then
                    searchCriteriaDate = Cont.value & "," & Cont.value
                Else
                    searchCriteriaDate = Cont.value & "," & TextBoxBirthDayTo.value
                End If
                Cont.Text = Format(Cont.value, "YYYY/MM/DD")
            Else
                If Cont.Text <> "" Then
                    Cont.SelStart = 0
                    Cont.SelLength = VBA.Len(Cont.Text)
                    Cancel = True
                End If
            End If
        Case Cont.Name = "TextBoxBirthDayTo"
            If VBA.IsDate(Cont.value) Then
                searchCriteriaDate = TextBoxBirthDayFrom.value & "," & Cont.value
                Cont.Text = Format(Cont.value, "YYYY/MM/DD")
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
    Dim i As Long
    Select Case True
'        Case Cont.Name = "CheckBoxEditFemale"
'            If Cont.value Then searchCriteriaSex = "女"
'            Call Filter2
'        Case Cont.Name = "CheckBoxEditMale"
'            If Cont.value Then searchCriteriaSex = "男"
'            Call Filter2
        Case Cont.Name = "ComboBoxAddress"
            Debug.Print Cont.Name & " Change:" & Cont.Text
            searchCriteriaAddress = Cont.Text
            Call Filter3
        Case Left(Cont.Name, 14) = "CheckBloodType"
            For i = 1 To 4
                With Me.Controls("CheckBloodType" & i)
                    If .value Then
                        searchCriteriaBloodType = searchCriteriaBloodType & "," & .Caption
                    End If
                End With
            Next i
            searchCriteriaBloodType = Right(searchCriteriaBloodType, Len(searchCriteriaBloodType) - 1)
        Case Left(Cont.Name, 11) = "CheckBoxSex"
            For i = 1 To 3
                With Me.Controls("CheckBoxSex" & i)
                    If .value Then
                        searchCriteriaSex = searchCriteriaSex & "," & .Caption
                    End If
                End With
            Next i
            searchCriteriaSex = Right(searchCriteriaSex, Len(searchCriteriaSex) - 1)

'        Case InStr(1, Cont.Name, "OptionButtonEditBloodType") > 0
'            searchCriteriaBloodType = Replace(Cont.Name, "OptionButtonEditBloodType", "")
'            Call Filter2
'        Case Cont.Name = "OptionButtonEditFemale"
'            If Cont.value Then searchCriteriaSex = "女"
'            Call Filter2
'        Case Cont.Name = "OptionButtonEditMale"
'            If Cont.value Then searchCriteriaSex = "男"
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
    Select Case True
'        Case Cont.Name = "LabelDatePickerEdit"
'            Cont.BackStyle = fmBackStyleOpaque
'            DatePicker.Init Me.TextBoxBirthDayEdit
'        Case Cont.Name = "LabelDatePickerNew"
'            Cont.BackStyle = fmBackStyleOpaque
'            DatePicker.Init Me.TextBoxBirthDayNew
        Case Left(Cont.Name, 15) = "LabelDatePicker"
            Call OpenDatePicker(Cont)
        Case Cont.Tag = "SideBar"
            Call OpenSideBar(Cont)
        Case Else
            Debug.Print Cont.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Call IControlEvent_OnClick(Cont)
    DoEvents
    Cancel = True
    Debug.Print Cont.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "ComboBoxAddress"
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
            If KeyCode = 187 And Shift = 2 Then Cont.value = Format(Now, "YYYY/MM/DD") ' Ctrl + 「;」

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
            Call Filter3
        Case Cont.Name = "OptionButtonEditFemale"
            If Cont.value Then searchCriteriaSex = "女"
            Call Filter3
        Case Cont.Name = "OptionButtonEditMale"
            If Cont.value Then searchCriteriaSex = "男"
            Call Filter3
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

'-------------------------------------------------------------------------------
'各コントロールのTagプロパティに応じてMouseHover時の処理を規定する
'処理対象とするコントロールにはコード内またはFormデザイン時にTagプロパティを設定しておくこと
Private Sub Hover(Cont As MSForms.IControl)
    'MouseOver
    Select Case True
        Case Cont.Name = "ComboBoxAddress"
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            HookComboBox Cont
        Case Cont.Name = "ListBoxEdit"
            If onFocusListBox Then Exit Sub
            onFocusListBox = True
            HookListBox Cont
        Case TypeName(Cont) = "Frame"
            If Left(Cont.Tag, 13) = "SelectionField" Then
                If onFocusFrame Then Exit Sub
                If Me.FrameFilter.ScrollBars = fmScrollBarsNone Then UnHook: Exit Sub
                onFocusFrame = True
                HookFrame Me.FrameFilter
            End If
        Case TypeName(Cont) = "Label"
            Select Case True
        '        Case "SelectDay", "SelectYear", "SelectMonth"
        '            Cont.Object.BorderStyle = fmBorderStyleSingle
                Case Cont.Tag = "Button"
'                    Cont.Object.BackColor = MouseOverColor
                    Cont.Object.BackStyle = fmBackStyleTransparent
'                    Cont.Object.BorderStyle = fmBorderStyleSingle
                Case Cont.Tag = "SideBar"
                    Cont.Object.BackColor = MouseOverColor
                Case Else
            End Select
        Case Else
            UnHook
            onFocusListBox = False
            onFocusComboBox = False
            onFocusFrame = False
    End Select
'    Select Case Cont.Tag
''        Case "SelectDay", "SelectYear", "SelectMonth"
''            Cont.Object.BorderStyle = fmBorderStyleSingle
'        Case "Button"
'            Cont.Object.BackStyle = fmBackStyleOpaque
'            Cont.Object.BorderStyle = fmBorderStyleSingle
'        Case Else
'    End Select
    
    'MouseOut
    If Not This.PrevControl Is Nothing Then
        If Not This.PrevControl Is Cont Then
'            This.PrevControl.Object.BorderStyle = fmBorderStyleNone
            Select Case This.PrevControl.Tag
                Case "Button"
'                    This.PrevControl.Object.BackColor = BaseBackColor
                    This.PrevControl.Object.BackStyle = fmBackStyleOpaque
'                    This.PrevControl.Object.BorderStyle = fmBorderStyleNone
                Case "SideBar"
                    This.PrevControl.Object.BackColor = FrameBaseColor
            End Select
        End If
    End If
    
    Set This.PrevControl = Cont
End Sub

Private Sub OpenSideBar(Cont As MSForms.IControl)
    Dim Frame As MSForms.Frame
    Set Frame = Cont.parent
    
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
        Frame.height = FrameFullHeight
        Cont.Object.BackColor = FrameOpenColor
    Else
        Frame.Tag = "SelectionField0"
        Frame.height = 18
        Cont.Object.BackColor = FrameBaseColor
    End If
    
    Dim FrameMargin As Single: FrameMargin = 6
    Me.FrameBloodTypeNo.Top = Me.FrameRegistorNo.Top + Me.FrameRegistorNo.height + FrameMargin
    Me.FrameName.Top = Me.FrameBloodTypeNo.Top + Me.FrameBloodTypeNo.height + FrameMargin
    Me.FrameAge.Top = Me.FrameName.Top + Me.FrameName.height + FrameMargin
    Me.FrameSex.Top = Me.FrameAge.Top + Me.FrameAge.height + FrameMargin
    Me.FrameBloodType.Top = Me.FrameSex.Top + Me.FrameSex.height + FrameMargin
    Me.FrameBirthDay.Top = Me.FrameBloodType.Top + Me.FrameBloodType.height + FrameMargin
    Me.FrameAddress.Top = Me.FrameBirthDay.Top + Me.FrameBirthDay.height + FrameMargin
    Me.FrameTarget.Top = Me.FrameAddress.Top + Me.FrameAddress.height + FrameMargin
    Me.FrameStation.Top = Me.FrameTarget.Top + Me.FrameTarget.height + FrameMargin
    
    FrameFilterFullHeight = Me.FrameRegistorNo.height + Me.FrameBloodTypeNo.height + _
                            Me.FrameName.height + Me.FrameAge.height + _
                            Me.FrameSex.height + Me.FrameBloodType.height + _
                            Me.FrameBirthDay.height + Me.FrameAddress.height + _
                            Me.FrameTarget.height + Me.FrameStation.height + FrameMargin * 7 + 12

    If Me.FrameFilter.height < FrameFilterFullHeight Then
        With Me.FrameFilter
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = FrameFilterFullHeight
'                    .ScrollTop = 0
            If isOpen Then
                If Frame.height > FrameFilterFullHeight Then
                    .ScrollTop = Frame.Top
                Else
                    .ScrollTop = .ScrollTop + Frame.height
                End If
            Else
                .ScrollTop = 0
            End If
        End With
        onFocusFrame = True
        HookFrame Me.FrameFilter
    Else
        With Me.FrameFilter
'                    .ScrollTop = 0
            .ScrollBars = fmScrollBarsNone
        End With
        UnHook
    End If
End Sub

Private Sub OpenDatePicker(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "LabelDatePickerEdit"
            Cont.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayEdit
        Case Cont.Name = "LabelDatePickerNew"
            Cont.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayNew
        Case Cont.Name = "LabelDatePickerFrom"
            Cont.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayFrom
        Case Cont.Name = "LabelDatePickerTo"
            Cont.BackStyle = fmBackStyleTransparent
            DatePicker.Init Me.TextBoxBirthDayTo
        Case Else
    End Select
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
        
        Set CriteriaItemTable = ThisWorkbook.Worksheets("List").ListObjects("T_抽出項目")
        Set CriteriaRange = criteriaSheet.Range("A1")
        
'        Dim i As Long
'        With ThisWorkbook.Worksheets("List").ListObjects("T_抽出項目")
'            Dim col As Long: col = 0
'            For i = 1 To .ListRows.count
'                criteriaSheet.Range("A1").Offset(0, col) = .DataBodyRange(i, 1)
'                If .DataBodyRange(i, 2) = "範囲" Then
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
'            .Range.AutoFilter Field:=.ListColumns("氏名").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
'        If Me.TextBoxAge.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("年齢").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
'        If Me.ComboBoxAddress.value <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("住所").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
'        If Me.CheckBoxFemale.value Or Me.CheckBoxMale Then _
'            .Range.AutoFilter Field:=.ListColumns("性別").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
'        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
'            .Range.AutoFilter Field:=.ListColumns("血液型").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
'        If Me.TextBoxDate.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("生年月日").Index, Criteria1:=Format(searchCriteriaDate, "YYYY年MM月DD日"), VisibleDropDown:=False
'
'        Dim CellsCnt As Long    '←絞り込みﾃﾞｰﾀのｾﾙ個数
'        Dim ColCnt As Long      '←ﾃｰﾌﾞﾙの列数
'        Dim buf1 As Variant    '←テーブル全体のデータ
''        buf1 = .Range.SpecialCells(xlCellTypeVisible)
'        buf1 = .Range
''        CellsCnt = .DataBodyRange.SpecialCells(xlCellTypeVisible).Count
'        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).count
'        ColCnt = UBound(buf1, 2)
''
'        Dim buf2 As Variant    '←戻り値とする一時的な配列
'        ReDim buf2(1 To (CellsCnt / ColCnt) - 1, 1 To ColCnt)
'
'        Dim i As Long            '←ｶｳﾝﾀ変数（配列の行位置）
'        Dim j As Long            '←ｶｳﾝﾀ変数（配列の列位置）
'        Dim k As Long            'テーブルのデータ行＋タイトル行の行数
'        For k = 2 To UBound(buf1, 1)
'          If .Range.Rows(k).Hidden = False Then
'            i = i + 1
'            For j = 1 To ColCnt
'              buf2(i, j) = buf1(k, j)
'            Next j
'          End If
'        Next k
'
'        'オートフィルタを解除
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
'            .Range.AutoFilter Field:=.ListColumns("氏名").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
'        If Me.TextBoxEditAge.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("年齢").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
'        If Me.ComboBoxEditAddress.value <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("住所").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
'        If Me.OptionButtonEditFemale.value Or Me.OptionButtonEditMale Then _
'            .Range.AutoFilter Field:=.ListColumns("性別").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
'        If Me.OptionButtonEditBloodTypeA Or Me.OptionButtonEditBloodTypeB Or Me.OptionButtonEditBloodTypeAB Or Me.OptionButtonEditBloodTypeO Then _
'            .Range.AutoFilter Field:=.ListColumns("血液型").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
'        If Me.TextBoxBirthDayEdit.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("生年月日").Index, Criteria1:=Format(searchCriteriaDate, "YYYY/MM/DD"), VisibleDropDown:=False
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
    Dim item As Variant
    Dim tmp As Variant
    Debug.Print "SetCriteria"
    
    CriteriaRange.CurrentRegion.ClearContents
    
    Dim i As Long
    With CriteriaItemTable
        Dim col As Long: col = 0
        For i = 1 To .ListRows.count
            criteriaItem = .DataBodyRange(i, 1)
            criteriaType = .DataBodyRange(i, 2)
            
            
'            Select Case criteriaType
'                Case "範囲"
'                    CriteriaRange.Offset(0, col + 1) = criteriaItem
'                Case Else
'            End Select
            
            tmp = Choose(i, searchCriteriaRegNo, searchCriteriaBtNo, _
                                searchCriteriaName, searchCriteriaAge, searchCriteriaDate, _
                                searchCriteriaSex, searchCriteriaBloodType, searchCriteriaAddress)
            Debug.Print tmp
            If tmp = "" Then
                CriteriaRange.Offset(0, col) = criteriaItem
                If criteriaType = "範囲" Then
                    col = col + 1
                    CriteriaRange.Offset(0, col) = criteriaItem
                End If
                GoTo CONTINUE
            End If
            tempArr = Split(tmp, ",")
            Debug.Print UBound(tempArr)
            ReDim conditions(UBound(tempArr))
            j = 0
            For Each item In tempArr
               If Not IsEmpty(item) Then
                   If item <> "" Then
                       conditions(j) = item
                       Debug.Print conditions(j)
                       j = j + 1
                   End If
               End If
            Next
            ReDim Preserve conditions(j - 1)
     
            CriteriaRange.Offset(0, col) = criteriaItem
            Select Case criteriaType
                Case "部分"
                    CriteriaRange.Offset(1, col) = "*" & conditions(0) & "*"
                Case "前方"
                    CriteriaRange.Offset(1, col) = conditions(0) & "*"
                Case "範囲"
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

    Me.ListBoxEdit.RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(resultSheet.Range("A2").CurrentRegion.Rows.count - 1, originalTable.ListColumns.count).Address
    
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

Private Sub AutoFitListbox()
    resultSheet.Cells.EntireColumn.AutoFit
    
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
        Next i
        .RowSource = resultSheet.Name & "!" & resultSheet.Range("A2").Resize(resultSheet.Range("A2").CurrentRegion.Rows.count, originalTable.ListColumns.count).Address
    End With
    
    Exit Sub
ERROR_HANDLER:
    MsgBox "列幅調整エラー:" & ERR.Number & " : " & ERR.Description
End Sub

Private Sub WebBrowserPreview_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Debug.Print URL
'    Cancel = True
End Sub

