VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   12790
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10970
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

Private searchCriteriaName As String
Private searchCriteriaAge As Long
Private searchCriteriaAddress As String
Private searchCriteriaSex As String
Private searchCriteriaBloodType As String
Private searchCriteriaDate As Date
Private searchCriteriaDateLevel As Long

Private originalTable As ListObject
Private workTable As ListObject
Private tmpSheet As Worksheet
Private tmpSheet2 As Worksheet

Private Const adOpenDynamic As Long = 2
Private Const adLockOptimistic As Long = 3
Private Const adStateClosed As Long = 0
Private Const adUseClient As Long = 3

Private dpFrom As DateTimePicker
Private onFocusListBox As Boolean
Private onFocusComboBox As Boolean

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

Private this As TState

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
       
    Dim MarginX As Long, MarginY As Long
    Dim FrameWidth As Long, FrameHeight As Long
    MarginX = 2: MarginY = 2
    FrameWidth = 0: FrameHeight = 0
    Dim i As Long
    With Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewTarget
        For i = 1 To UBound(TargetList, 1)
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxNewTarget" & i)
                .Caption = TargetList(i, 1)
                .GroupName = TargetList(i, 2)
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = 6 + (.Width + MarginX) * (TargetList(i, 3) - 1)
                .Top = 12 + (.Height + MarginY) * (TargetList(i, 4) - 1)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                If .Left + .Width + 6 > FrameWidth Then FrameWidth = .Left + .Width + 6
                If .Top + .Height + 12 > FrameHeight Then FrameHeight = .Top + .Height + 12
            End With
        Next i
        .Width = FrameWidth
        .Height = FrameHeight
    End With
    FrameWidth = 0: FrameHeight = 0
    With Me.MultiPageSwitchMode("PageSearchAndEdit").FrameEditTarget
        For i = 1 To UBound(TargetList, 1)
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxEditTarget" & i)
                .Caption = TargetList(i, 1)
                .GroupName = TargetList(i, 2)
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = 6 + (.Width + MarginX) * (TargetList(i, 3) - 1)
                .Top = 12 + (.Height + MarginY) * (TargetList(i, 4) - 1)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                If .Left + .Width + 6 > FrameWidth Then FrameWidth = .Left + .Width + 6
                If .Top + .Height + 12 > FrameHeight Then FrameHeight = .Top + .Height + 12
            End With
        Next i
        .Width = FrameWidth
        .Height = FrameHeight
    End With
    
    TargetList = Sheet3.ListObjects("T_曲").DataBodyRange
    With Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewStation
        For i = 1 To UBound(TargetList, 1)
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxNewStation" & i)
                .Caption = TargetList(i, 1)
                .GroupName = TargetList(i, 2)
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = 6 + (.Width + MarginX) * (TargetList(i, 3) - 1)
                .Top = 12 + (.Height + MarginY) * (TargetList(i, 4) - 1)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                If .Left + .Width + 6 > FrameWidth Then FrameWidth = .Left + .Width + 6
                If .Top + .Height + 12 > FrameHeight Then FrameHeight = .Top + .Height + 12
            End With
        Next i
        .Left = Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewTarget.Left + Me.MultiPageSwitchMode("PageRegistorNewItem").FrameNewTarget.Width + 4
        .Width = FrameWidth
        .Height = FrameHeight
    End With
    FrameWidth = 0: FrameHeight = 0
    With Me.MultiPageSwitchMode("PageSearchAndEdit").FrameEditStation
        For i = 1 To UBound(TargetList, 1)
            With .Controls.Add("Forms.CheckBox.1", "CheckBoxEditStation" & i)
                .Caption = TargetList(i, 1)
                .GroupName = TargetList(i, 2)
                .SpecialEffect = fmButtonEffectFlat
                .Width = 40
                .Height = 20
                .Left = 6 + (.Width + MarginX) * (TargetList(i, 3) - 1)
                .Top = 12 + (.Height + MarginY) * (TargetList(i, 4) - 1)
                .Font.Name = "Yu Gothic UI"
                .Font.size = 10
                .Font.Bold = False
                If .Left + .Width + 6 > FrameWidth Then FrameWidth = .Left + .Width + 6
                If .Top + .Height + 12 > FrameHeight Then FrameHeight = .Top + .Height + 12
            End With
        Next i
        .Left = Me.MultiPageSwitchMode("PageSearchAndEdit").FrameEditTarget.Left + Me.MultiPageSwitchMode("PageSearchAndEdit").FrameEditTarget.Width + 4
        .Width = FrameWidth
        .Height = FrameHeight
    End With
    
End Sub

Private Sub UserForm_Initialize()
    startTime = Timer
    searchCriteriaAge = -1
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_都道府県").ListColumns("都道府県名").DataBodyRange
    Dim i As Long
    With ComboBoxEditAddress
        For i = 1 To listRange.Rows.count
            .AddItem listRange(i)
        Next
    End With
    Call CreateCheckBox
    
    Call AddTemporarySheet
    With Me.ListBoxEdit
        .Clear
        .ColumnHeads = True
'        .RowSource = workTable.DataBodyRange.Address
'        .RowSourceType = "Table/Query"
'        .RowSource = originalTable.DataBodyRange.Address
        .RowSource = tmpSheet2.Name & "!" & tmpSheet2.Range("A2").Resize(originalTable.ListRows.count, originalTable.ListColumns.count).Address
    End With
    Call AutoFitListbox
'    With TextBoxDatePickerTest
'        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
''        .DropButtonStyle = fmDropButtonStyleReduce
'    End With
'    ChooseHook_ListBox Me.ListBoxResultList
'    ChooseHook_ComBox Me.ComboBoxAddress
'    Set dpFrom = New DateTimePicker
'    With dpFrom
'        .Add Me.TextBoxDatePickerTest
'        .Create Me, "DD/MM/YYYY" ', _
''            BackColor:=&H492B27, _
''            TitleBack:=RGB(39, 56, 151), _
''            Trailing:=&H80000010, _
''            TitleFore:=&HFFFFFF
'    End With
    Set this.Control = New ControlEvents
    With this.Control
        .parent = Me
        .Init
    End With
    Debug.Print Timer - startTime
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.LabelEditDatePicker.SpecialEffect = fmSpecialEffectFlat
    uHook
    onFocusListBox = False
    onFocusComboBox = False
End Sub

Private Sub UserForm_Terminate()
    Set originalTable = Nothing
'    Set workTable = Nothing
'
'    With ThisWorkbook.Worksheets("Temp")
'        .Visible = True
'    End With
'
'    Application.DisplayAlerts = False
'    tmpSheet.Delete
    Set tmpSheet = Nothing
'    Application.DisplayAlerts = True
    uHook
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

'--------------------インターフェイスからコールバックされるメンバ関数
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
    End Select
    Debug.Print Cont.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(Cont As MSForms.IControl, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Select Case True
        Case Cont.Name = "TextBoxNewBirthDay"
            If VBA.IsDate(Cont.value) Then
                searchCriteriaDate = Cont.value
                Cont.Text = Format(searchCriteriaDate, "YYYY年MM月DD日")
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
                Cont.Text = Format(searchCriteriaDate, "YYYY年MM月DD日")
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
'            If Cont.value Then searchCriteriaSex = "女"
'            Call Filter2
'        Case Cont.Name = "CheckBoxEditMale"
'            If Cont.value Then searchCriteriaSex = "男"
'            Call Filter2
        Case Cont.Name = "ComboBoxEditAddress"
            searchCriteriaAddress = Cont.Text
            Call Filter3

    End Select
    Debug.Print Cont.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(Cont As MSForms.IControl)
    Dim pos As POINTAPI
    Select Case True
        Case Cont.Name = "LabelEditDatePicker"
            Cont.SpecialEffect = fmSpecialEffectFlat
            pos = GetControlPosition(Cont)
            Debug.Print pos.x
            Call DatePicker.Init(pos.y, pos.x)
            Me.TextBoxEditBirthDay.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
        Case Cont.Name = "LabelNewDatePicker"
            Cont.SpecialEffect = fmSpecialEffectFlat
            Call DatePicker.Init
            Me.TextBoxNewBirthDay.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
        Case Else
            Debug.Print Cont.Name & " Click"
    End Select
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(Cont As MSForms.IControl)
    Select Case True
        Case Cont.Name = "ComboBoxEditAddress"
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            ChooseHook_ComBox Cont

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
        Case Cont.Name = "ComboBoxEditAddress"
            If onFocusComboBox Then Exit Sub
            onFocusComboBox = True
            ChooseHook_ComBox Cont
        Case Cont.Name = "ListBoxEdit"
            If onFocusListBox Then Exit Sub
            onFocusListBox = True
            ChooseHook_ListBox Cont
        Case Cont.Name = "LabelEditDatePicker"
            Cont.SpecialEffect = fmSpecialEffectEtched
        Case Cont.Name = "LabelNewDatePicker"
            Cont.SpecialEffect = fmSpecialEffectEtched
        Case Else
            Me.LabelEditDatePicker.SpecialEffect = fmSpecialEffectFlat
            Me.LabelNewDatePicker.SpecialEffect = fmSpecialEffectFlat
    End Select
'    Debug.Print Cont.Name & " MouseMove:"
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

'-------------------------------------------------------------------------------
Private Sub AddTemporarySheet()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Dummy")
        Set originalTable = .ListObjects("T_Dummy")
        originalTable.ShowAutoFilter = False
        .Activate
'        If Util.ExistsWorksheet("Temp") Then ' ThisWorkbook.Worksheets("Temp").Delete
'            Set tmpSheet = ThisWorkbook.Worksheets("Temp")
'            tmpSheet.Cells.Clear
'        Else
'            Set tmpSheet = Sheets.Add
'            tmpSheet.Name = "Temp"
'        End If
        If Util.ExistsWorksheet("Tmp2") Then ' ThisWorkbook.Worksheets("Tmp2").Delete
            Set tmpSheet2 = ThisWorkbook.Worksheets("Tmp2")
            tmpSheet2.Cells.Clear
        Else
            Set tmpSheet2 = Sheets.Add
            tmpSheet2.Name = "Tmp2"
        End If
'        If Util.ExistsWorksheet("Temp") Then Set tmpSheet = ThisWorkbook.Worksheets("Temp")
'        Set tmpSheet2 = ThisWorkbook.Worksheets("Tmp2")
        With tmpSheet2
'            .Name = "Temp"
            originalTable.HeaderRowRange.Copy .Range("A1")
'            originalTable.Range.Copy .Range("A1")
            With .Range("A1").CurrentRegion.Font
                .Name = Me.ListBoxEdit.Font.Name
                .size = Me.ListBoxEdit.Font.size
            End With
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
            .Range.AutoFilter Field:=.ListColumns("氏名").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("年齢").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxAddress.value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("住所").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.CheckBoxFemale.value Or Me.CheckBoxMale Then _
            .Range.AutoFilter Field:=.ListColumns("性別").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("血液型").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxDate.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("生年月日").Index, Criteria1:=Format(searchCriteriaDate, "YYYY年MM月DD日"), VisibleDropDown:=False
    
        Dim CellsCnt As Long    '←絞り込みﾃﾞｰﾀのｾﾙ個数
        Dim ColCnt As Long      '←ﾃｰﾌﾞﾙの列数
        Dim buf1 As Variant    '←テーブル全体のデータ
'        buf1 = .Range.SpecialCells(xlCellTypeVisible)
        buf1 = .Range
'        CellsCnt = .DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).count
        ColCnt = UBound(buf1, 2)
'
        Dim buf2 As Variant    '←戻り値とする一時的な配列
        ReDim buf2(1 To (CellsCnt / ColCnt) - 1, 1 To ColCnt)

        Dim i As Long            '←ｶｳﾝﾀ変数（配列の行位置）
        Dim j As Long            '←ｶｳﾝﾀ変数（配列の列位置）
        Dim k As Long            'テーブルのデータ行＋タイトル行の行数
        For k = 2 To UBound(buf1, 1)
          If .Range.Rows(k).Hidden = False Then
            i = i + 1
            For j = 1 To ColCnt
              buf2(i, j) = buf1(k, j)
            Next j
          End If
        Next k
               
        'オートフィルタを解除
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
    ThisWorkbook.Worksheets("Temp").Activate
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
            .Range.AutoFilter Field:=.ListColumns("氏名").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxEditAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("年齢").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxEditAddress.value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("住所").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.OptionButtonEditFemale.value Or Me.OptionButtonEditMale Then _
            .Range.AutoFilter Field:=.ListColumns("性別").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonEditBloodTypeA Or Me.OptionButtonEditBloodTypeB Or Me.OptionButtonEditBloodTypeAB Or Me.OptionButtonEditBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("血液型").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxEditBirthDay.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("生年月日").Index, Criteria1:=Format(searchCriteriaDate, "YYYY/MM/DD"), VisibleDropDown:=False
        
        tmpSheet.Cells.Clear
        Dim CellsCnt As Long
        
        If .ListColumns(1).Range.SpecialCells(xlCellTypeVisible).count = 1 Then
            CellsCnt = 1
        Else
            CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
        End If
        .Range.SpecialCells(xlCellTypeVisible).Copy tmpSheet.Range("A1")
        With tmpSheet.Range("A1").CurrentRegion.Font
            .Name = Me.ListBoxEdit.Font.Name
            .size = Me.ListBoxEdit.Font.size
        End With
        Call AutoFitListbox

        .ShowAutoFilter = False
        Me.ListBoxEdit.RowSource = tmpSheet.Name & "!" & tmpSheet.Range("A2").Resize(CellsCnt, .ListColumns.count).Address
    End With
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

' セル範囲を"[シート名$セルアドレス(A1形式)]"に変換するヘルパー関数
Private Function ToExcelTableName(ByVal rng As Range) As String
       ToExcelTableName = "[" & rng.parent.Name & "$" & rng.Address(False, False) & "]"
End Function

Private Sub Filter3()
    startTime = Timer
    Application.ScreenUpdating = False

'    'インスタンスを作成（事前バインディングの場合）
'    Dim dbConnection As Object
'    Dim dbRecordset As Object
'    Set dbConnection = CreateObject("ADODB.Connection")
'    Set dbRecordset = CreateObject("ADODB.Recordset")

'    'ADO接続
'    dbConnection.Provider = "Microsoft.ACE.OLEDB.12.0"
'    dbConnection.Properties("Extended Properties") = "Excel 12.0;HDR=Yes"
'    dbConnection.Open ThisWorkbook.FullName

'    Public Function GetXLSConnection(DataSource As String) As Object
      Dim dbConnection             As Object
'      Dim strCNString       As String

      'レイトバインディング
        Set dbConnection = CreateObject("ADODB.Connection")
        dbConnection.Provider = "Microsoft.ACE.OLEDB.12.0"
        dbConnection.Properties("Extended Properties") = "Excel 12.0;HDR=Yes"
'        dbConnection.Provider = "MSDASQL"
'        dbConnection.Properties("Extended Properties") = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb);" & _
'                                "DBQ=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name
        dbConnection.Open ThisWorkbook.Path & "\" & ThisWorkbook.Name

'      '接続文字列
'        strCNString = "Provider=Microsoft.ACE.OLEDB.12.0;" _
'                            & "Data Source=" & DataSource & ";" _
'
'      '接続
'        dbConnection.Open strCNString

'      '接続を返す
'        Set GetXLSConnection = dbConnection
'    End Function

    Dim strSQL As String
    strSQL = "SELECT"                                           '抽出フィールド(項目)を指定
      strSQL = strSQL & "  [氏名]"
      strSQL = strSQL & ", [氏名（ひらがな）]"
      strSQL = strSQL & ", [電話番号]"
'      strSQL = strSQL & ", MONTH([誕生日]) AS [誕生月]"
    strSQL = strSQL & " FROM " & ToExcelTableName(originalTable.Range)                'データテーブルを指定
    strSQL = strSQL & " WHERE 1 = 1"                            '抽出条件
      strSQL = strSQL & " AND [性別] = '男'"                    '性別=男
      strSQL = strSQL & " AND [年齢] >= 30"                     '年齢=30歳以上
      strSQL = strSQL & " AND [年齢] <  50"                     '50歳未満
      strSQL = strSQL & " AND [血液型] = 'A'"

'    Public Function getRecordset(ByRef dbConnection As Object, ByVal strSQL As String, ByRef dbRecordset As Object) As Boolean
'      getRecordset = False
'
      On Error GoTo ERR_PROC
      Dim dbRecordset As Object
      Set dbRecordset = CreateObject("ADODB.Recordset")
      dbRecordset.CursorLocation = adUseClient
      dbRecordset.Open strSQL, dbConnection, adOpenDynamic, adLockOptimistic
'      getRecordset = True
'
      GoTo END_PROC
ERR_PROC:
      MsgBox "エラー" & ERR.Number & ":" & ERR.Description
      
'
END_PROC:
'    End Function

    'SQL文の実行（シートの範囲を指定してテーブルとする）
'    dbRecordset.Open "SELECT * FROM [Sheet1$B4:F] WHERE 区分 = '果物' ORDER BY 単価 DESC", dbConnection
'    dbRecordset.Open Source:="SELECT * FROM [Sheet1$B4:F] WHERE 区分 = '果物' ORDER BY 単価 DESC", _
'                    ActiveConnection:=dbConnection
'    Dim tmpSheet2 As Worksheet
'    Set tmpSheet2 = ThisWorkbook.Worksheets("Tmp2")
    '抽出結果を出力
    With tmpSheet2
        '出力エリアにある既存データを消去
        .Cells.Clear
'        With .Range("rngXDB_DataTop")
        With .Range("A1")
'            .CurrentRegion.ClearContents
            Dim lngF As Long
            'フィールド(項目)名を出力
            For lngF = 0 To dbRecordset.Fields.count - 1
                .Offset(, lngF).value = dbRecordset.Fields(lngF).Name
            Next lngF

            'データを出力
            .Offset(1).CopyFromRecordset dbRecordset
        End With
    End With

'    '取得した内容（Recordset）の確認
'    Do Until dbRecordset.EOF
'        Debug.Print dbRecordset!品名 & ", " & dbRecordset!単価
'        dbRecordset.MoveNext
'    Loop


    'レコードセットを閉じる
'    Public Sub CloseRecordSet(dbRecordset As Object)
      If dbRecordset.State <> adStateClosed Then
        dbRecordset.Close
      End If
      Set dbRecordset = Nothing
'    End Sub
'    Call CloseRecordSet(objRS)

    'コネクションを閉じる
'    Public Sub CloseConnection(dbConnection As Object)
      '接続された状態であるならば
      If dbConnection.State <> adStateClosed Then
        dbConnection.Close
      End If
      Set dbConnection = Nothing
'    End Sub
'    Call CloseConnection(objCN)

    'メモリの解放（無くとも構わない）
'    dbRecordset.Close: Set dbRecordset = Nothing
'    If dbConnection.State <> adStateClosed Then
'        dbConnection.Close
'    End If
'    Set dbConnection = Nothing

    '    ThisWorkbook.Worksheets("Dummy").Activate
'    workTable.DataBodyRange.Delete
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
'        If Me.TextBoxEditBirthDay.Text <> "" Then _
'            .Range.AutoFilter Field:=.ListColumns("生年月日").Index, Criteria1:=Format(searchCriteriaDate, "YYYY年MM月DD日"), VisibleDropDown:=False
'
'        tmpSheet.Cells.Clear
'        Dim CellsCnt As Long
'
'        If .ListColumns(1).Range.SpecialCells(xlCellTypeVisible).count = 1 Then
'            CellsCnt = 1
'        Else
'            CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
'        End If
'        .Range.SpecialCells(xlCellTypeVisible).Copy tmpSheet.Range("A1")
'        With tmpSheet.Range("A1").CurrentRegion.Font
'            .Name = Me.ListBoxEdit.Font.Name
'            .size = Me.ListBoxEdit.Font.size
'        End With
'        .ShowAutoFilter = False
'        Me.ListBoxEdit.RowSource = tmpSheet.Name & "!" & tmpSheet.Range("A2").Resize(CellsCnt, .ListColumns.count).Address
'    End With
    Me.ListBoxEdit.RowSource = tmpSheet2.Name & "!" & tmpSheet2.Range("A2").Resize(lngF, originalTable.ListColumns.count).Address
    Application.ScreenUpdating = True
    Debug.Print Timer - startTime
End Sub

'Private Sub FitListColumnWidthToText()
'    Dim maxColumn As Long
'    maxColumn = Me.ListBoxResultList.ColumnCount
''    Dim widthArray() As Long
'    Dim widthArray() As String
'    ReDim widthArray(0 To maxColumn - 1)
'
'    Dim i As Long
'    Dim col As Long
'    Dim textString As String
'    Dim textWidth As Long
'    With Me.ListBoxResultList
'        For i = 0 To .ListCount - 1
'            For col = 0 To maxColumn - 1
'                textString = .List(i, col)
'                textWidth = Util.MeasureTextWidth(textString, .Font.Name, .Font.size)
'                If textWidth > Val(widthArray(col)) Then widthArray(col) = CStr(textWidth)
'            Next col
'        Next i
'        .ColumnWidths = VBA.Join(widthArray, ";")
'    End With
'End Sub

Private Sub AutoFitListbox()
'    Dim WS As Worksheet
'
'    Set WS = ThisWorkbook.Sheets("Temp")
    tmpSheet2.Cells.EntireColumn.AutoFit
    
    With Me.ListBoxEdit
        Dim maxColumn As Long
        maxColumn = .ColumnCount
        .ColumnWidths = ""
        Dim i As Long
        For i = 1 To maxColumn - 1
            .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & tmpSheet2.Cells(1, i).Width
        Next i
    End With
    
'    UserForm1.ListBox1.RowSource = WS.Range("A2:E" & LS + 1).Address(External:=True)
End Sub

