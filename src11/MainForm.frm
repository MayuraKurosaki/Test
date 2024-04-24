VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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

Private dpFrom As DateTimePicker
Private onFocusListBox As Boolean
Private onFocusComboBox As Boolean

Private Sub CheckBoxFemale_AfterUpdate()
'    Call Filter
End Sub

'承認者:approver　承認する:Approve　承認:Approval
'署名:signature
'制約:constraint
'OperationProcedure
'Reason for operation
'Operation results
'TimeUnit
'認証:authentication

Private Sub CheckBoxFemale_Change()
    If Me.CheckBoxFemale.value Then searchCriteriaSex = "女"
    Call Filter2
End Sub

Private Sub CheckBoxFemale_Click()

End Sub

Private Sub CheckBoxMale_AfterUpdate()
'    Call Filter
End Sub

Private Sub CheckBoxMale_Change()
    If Me.CheckBoxMale.value Then searchCriteriaSex = "男"
    Call Filter2
End Sub

Private Sub CheckBoxMale_Click()

End Sub

Private Sub ComboBoxAddress_AfterUpdate()
'    Call Filter
End Sub

'Private Sub ComboBoxAddress_AfterUpdate()
'    Debug.Print Me.ComboBoxAddress.Text
'End Sub

Private Sub ComboBoxAddress_Change()
    searchCriteriaAddress = Me.ComboBoxAddress.Text
    Call Filter2
End Sub

Private Sub ComboBoxAddress_DropButtonClick()
'    Dim listRange As Range
'    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_都道府県").ListColumns("都道府県名").DataBodyRange
'    Dim I As Long
'    With ComboBoxAddress
'        For I = 1 To listRange.Rows.count
'            .AddItem listRange(I)
'        Next
'    End With
    If onFocusComboBox Then Exit Sub
    onFocusComboBox = True
    ChooseHook_ComBox Me.ComboBoxAddress
End Sub

Private Sub ComboBoxAddress_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If onFocusComboBox Then Exit Sub
    onFocusComboBox = True
    ChooseHook_ComBox Me.ComboBoxAddress
End Sub

Private Sub CommandButtonDatePicker_Click()
    Call DatePicker.Init
    Me.TextBoxDate.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
End Sub

Private Sub ListBoxResultList_AfterUpdate()
    Debug.Print "ListBoxUpdated"
'    Call FitListColumnWidthToText
End Sub

Private Sub ListBoxResultList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBoxResultList_Change()
    Debug.Print "ListBoxChanged"
'    Call FitListColumnWidthToText
End Sub

Private Sub ListBoxResultList_Click()
    Debug.Print Me.ListBoxResultList.ListIndex
End Sub

Private Sub ListBoxResultList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If onFocusListBox Then Exit Sub
    onFocusListBox = True
    ChooseHook_ListBox Me.ListBoxResultList
End Sub

Private Sub OptionButtonBloodTypeA_AfterUpdate()
'    Call Filter
End Sub

Private Sub OptionButtonBloodTypeB_Change()
    If Me.OptionButtonBloodTypeB.value Then searchCriteriaBloodType = "B"
    Call Filter2
End Sub

Private Sub OptionButtonBloodTypeB_Click()

End Sub

Private Sub OptionButtonBloodTypeA_Change()
    If Me.OptionButtonBloodTypeA.value Then searchCriteriaBloodType = "A"
    Call Filter2
End Sub

Private Sub OptionButtonBloodTypeA_Click()

End Sub

Private Sub OptionButtonBloodTypeAB_Change()
    If Me.OptionButtonBloodTypeAB.value Then searchCriteriaBloodType = "AB"
    Call Filter2
End Sub

Private Sub OptionButtonBloodTypeAB_Click()

End Sub

Private Sub OptionButtonBloodTypeO_Change()
    If Me.OptionButtonBloodTypeO.value Then searchCriteriaBloodType = "O"
    Call Filter2
End Sub

Private Sub OptionButtonBloodTypeO_Click()

End Sub

Private Sub OptionButtonFemale_Change()

End Sub

Private Sub OptionButtonFemale_Click()

End Sub

Private Sub OptionButtonMale_Change()

End Sub

Private Sub OptionButtonMale_Click()

End Sub

Private Sub TextBoxAge_AfterUpdate()
    If Me.TextBoxAge.Text = "" Then
        searchCriteriaAge = -1
    Else
        searchCriteriaAge = TextBoxAge.Text
    End If
    Call Filter2
End Sub

Private Sub TextBoxAge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxAge_Change()

End Sub

Private Sub TextBoxDate_AfterUpdate()
    Call Filter2
End Sub

Private Sub TextBoxDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If VBA.IsDate(Me.TextBoxDate) Then
        searchCriteriaDate = Me.TextBoxDate.value
        Me.TextBoxDate.Text = Format(searchCriteriaDate, "YYYY年MM月DD日")
    Else
        If Me.TextBoxDate.Text <> "" Then
            Me.TextBoxDate.SelStart = 0
            Me.TextBoxDate.SelLength = VBA.Len(Me.TextBoxDate.Text)
            Cancel = True
        End If
    End If
End Sub

Private Sub TextBoxDate_Change()

End Sub

Private Sub TextBoxDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 187 And Shift = 2 Then TextBoxDate.value = Format(Now, "YYYY年MM月DD日") ' Ctrl + 「;」
End Sub

Private Sub TextBoxName_AfterUpdate()
    searchCriteriaName = TextBoxName.Text
    Call Filter2
End Sub

Private Sub TextBoxName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxName_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    searchCriteriaAge = -1
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_都道府県").ListColumns("都道府県名").DataBodyRange
    Dim I As Long
    With ComboBoxAddress
        For I = 1 To listRange.Rows.count
            .AddItem listRange(I)
        Next
    End With
    
    Call AddTemporarySheet
'    Call CopyTable
    With Me.ListBoxResultList
        .Clear
        .ColumnHeads = True
'        .RowSource = workTable.DataBodyRange.Address
'        .RowSourceType = "Table/Query"
'        .RowSource = originalTable.DataBodyRange.Address
        .RowSource = tmpSheet.Name & "!" & tmpSheet.Range("A2").Resize(originalTable.ListRows.count, originalTable.ListColumns.count).Address
    End With
    Call AutoFitListbox
'    With TextBoxDatePickerTest
'        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
''        .DropButtonStyle = fmDropButtonStyleReduce
'    End With
    ChooseHook_ListBox Me.ListBoxResultList
    ChooseHook_ComBox Me.ComboBoxAddress
    Set dpFrom = New DateTimePicker
    With dpFrom
        .Add Me.TextBoxDatePickerTest
        .Create Me, "DD/MM/YYYY" ', _
'            BackColor:=&H492B27, _
'            TitleBack:=RGB(39, 56, 151), _
'            Trailing:=&H80000010, _
'            TitleFore:=&HFFFFFF
    End With
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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

Private Sub AddTemporarySheet()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Dummy")
        Set originalTable = .ListObjects("T_Dummy")
        .Activate
'        If Util.ExistsWorksheet("Temp") Then ThisWorkbook.Worksheets("Temp").Delete
'        Set tmpSheet = Sheets.Add
        If Util.ExistsWorksheet("Temp") Then Set tmpSheet = ThisWorkbook.Worksheets("Temp")
        With tmpSheet
'            .name = "Temp"
'            originalTable.HeaderRowRange.Copy tmpSheet.Range("A1")
            originalTable.Range.Copy tmpSheet.Range("A1")
            With .Range("A1").CurrentRegion.Font
                .Name = Me.ListBoxResultList.Font.Name
                .size = Me.ListBoxResultList.Font.size
            End With
'            .Visible = False
        End With
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub CopyTable()
    Application.ScreenUpdating = False
'    With ThisWorkbook
'        Dim sheetIndex As Long
'        sheetIndex = .Worksheets("Dummy").index
'        .Worksheets("Dummy").Copy After:=.Worksheets(.Worksheets.Count)
'        .Worksheets(sheetIndex).Activate
'        With .Worksheets(.Worksheets.Count)
'            .Name = "Temp"
'            .Visible = False
'        End With
'    End With

    With ThisWorkbook.Worksheets("Dummy")
        Set originalTable = .ListObjects("T_Dummy")
        .Activate
'        Set tmpSheet = Sheets.Add
'        With tmpSheet
'            .Name = "Temp"
'            originalTable.Range.Copy .Range("A1")
'            Set workTable = .ListObjects(1)
'            .Visible = False
'        End With
'        With Sheets.Add
'            .Name = "Temp"
'            originalTable.Range.Copy .Range("A1")
'            Set workTable = .ListObjects(1)
'            .Visible = False
'        End With
    End With
    
'    Set originalTable = ThisWorkbook.Worksheets("Dummy").ListObjects("T_Dummy")
'    Set workTable = ThisWorkbook.Worksheets("Temp").ListObjects(1)
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

        Dim I As Long            '←ｶｳﾝﾀ変数（配列の行位置）
        Dim j As Long            '←ｶｳﾝﾀ変数（配列の列位置）
        Dim k As Long            'テーブルのデータ行＋タイトル行の行数
        For k = 2 To UBound(buf1, 1)
          If .Range.Rows(k).Hidden = False Then
            I = I + 1
            For j = 1 To ColCnt
              buf2(I, j) = buf1(k, j)
            Next j
          End If
        Next k
               
        'オートフィルタを解除
        .Range.AutoFilter
        .ShowAutoFilter = False
    End With
    With workTable
        .DataBodyRange.Delete
        .Range(2, 1).Resize(I, j) = buf2
'        .Range(2, 1).Resize(UBound(buf2, 1), UBound(buf2, 2)) = buf2
    End With
    Erase buf1
    Erase buf2
    ThisWorkbook.Worksheets("Temp").Activate
    Me.ListBoxResultList.RowSource = workTable.DataBodyRange.Address
    Application.ScreenUpdating = True
End Sub

Private Sub Filter2()
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
        
        tmpSheet.Cells.Clear
        Dim CellsCnt As Long    '←絞り込みﾃﾞｰﾀのｾﾙ個数
        CellsCnt = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).count
        .Range.SpecialCells(xlCellTypeVisible).Copy tmpSheet.Range("A1")
            With tmpSheet.Range("A1").CurrentRegion.Font
                .Name = Me.ListBoxResultList.Font.Name
                .size = Me.ListBoxResultList.Font.size
            End With

'        Debug.Print sh.ListObjects.Count
'        If Not workTable.DataBodyRange Is Nothing Then workTable.DataBodyRange.Delete
'        Debug.Print sh.ListObjects.Count
'        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy sh.Range("A2")
'        Set workTable = sh.ListObjects(1)
'        Debug.Print sh.ListObjects.Count
'        sh.Visible = True
        
'        Dim CellsCnt As Long    '←絞り込みﾃﾞｰﾀのｾﾙ個数
'        Dim ColCnt As Long      '←ﾃｰﾌﾞﾙの列数
'        Dim buf1 As Variant    '←テーブル全体のデータ
''        buf1 = .Range.SpecialCells(xlCellTypeVisible)
'        buf1 = .Range
''        CellsCnt = .DataBodyRange.SpecialCells(xlCellTypeVisible).Count
'        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).Count
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
               
        'オートフィルタを解除
'        .Range.AutoFilter
'        .ShowAutoFilter = False
    End With
'    With workTable
'        .DataBodyRange.Delete
'        .Range(2, 1).Resize(i, j) = buf2
''        .Range(2, 1).Resize(UBound(buf2, 1), UBound(buf2, 2)) = buf2
'    End With
'    Erase buf1
'    Erase buf2
'    ThisWorkbook.Worksheets("Temp").Activate
'    Me.ListBoxResultList.RowSource = workTable.DataBodyRange.Address
'    Me.ListBoxResultList.RowSourceType = "Table/Query"
'    Me.ListBoxResultList.RowSource = originalTable.DataBodyRange.Address
'    tmpSheet.Activate
'    Me.ListBoxResultList.RowSource = ""
'    Me.ListBoxResultList.Clear
'    Debug.Print Me.ListBoxResultList.ListCount
    Me.ListBoxResultList.RowSource = tmpSheet.Name & "!" & tmpSheet.Range("A2").Resize(CellsCnt, originalTable.ListColumns.count).Address
'    Debug.Print Me.ListBoxResultList.ListCount
'    Debug.Print CellsCnt
    originalTable.ShowAutoFilter = False
    Application.ScreenUpdating = True
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
    Dim WS As Worksheet
'    Dim LS, LastColumn, i As Long
'    Dim objek As String
    
    Set WS = ThisWorkbook.Sheets("Temp")
'    LS = WS.Range("A" & Rows.count).End(xlUp).Row
'    objek = "userform1.listbox1_"
    WS.Cells.EntireColumn.AutoFit
    
        With Me.ListBoxResultList
            Dim maxColumn As Long
            maxColumn = .ColumnCount
'            .ColumnCount = 13
            .ColumnWidths = ""
            Dim I As Long
            For I = 1 To maxColumn - 1
                .ColumnWidths = .ColumnWidths & IIf(I > 1, ";", "") & WS.Cells(1, I).Width
            Next I
        End With
    
'    UserForm1.ListBox1.RowSource = WS.Range("A2:E" & LS + 1).Address(External:=True)
End Sub

