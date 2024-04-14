Attribute VB_Name = "Search"
Option Explicit

Sub ExtractData()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("ExtractedData")
    
    '各シートの最終行を取得
    Dim cmax1 As Long, cmax2 As Long
    cmax1 = ws1.Range("A65536").End(xlUp).row
    cmax2 = ws2.Range("A65536").End(xlUp).row
    
    'データをリセット
    ws2.Range("B6:B7").ClearContents
    If Not cmax2 = 9 Then: ws2.Range("A10:E" & cmax2).ClearContents
    
    '開始日と終了日を取得
    Dim startdate As Date, enddate As Date
    startdate = ws2.Range("B2").value
    enddate = ws2.Range("B3").value
    
    '取引先を取得
    Dim torihiki As String
    torihiki = ws2.Range("B4").value
    
    '開始日、終了日、取引先が空欄か判定
    Dim flag(2) As Boolean ' BooleanのDefault値はFalse
    If startdate = 0 Then: flag(0) = True
    If enddate = 0 Then: flag(1) = True
    If torihiki = "" Then: flag(2) = True
    
    '変数の初期化
    Dim n As Long: n = 10
    Dim goukei As Long: goukei = 0
    Dim kensu As Long: kensu = 0
    
    '条件に合致した行を抽出
    Dim i As Long
    For i = 2 To cmax1
        If flag(0) = False Then
            If ws1.Range("C" & i).value < startdate Then: GoTo Continue
        End If
        
        If flag(1) = False Then
            If ws1.Range("C" & i).value >= enddate Then: GoTo Continue
        End If
        
        If flag(2) = False Then
            If ws1.Range("E" & i) <> torihiki Then: GoTo Continue
        End If
    
        '条件に合致した行のデータのみを対象して分析
        ws2.Range("A" & n & ":E" & n).value = ws1.Range("A" & i & ":E" & i).value
        goukei = goukei + ws1.Range("D" & i).value
        kensu = kensu + 1
        n = n + 1
        
Continue:
    Next
        
    ws2.Range("B6").value = goukei
    ws2.Range("B7").value = kensu
End Sub

Sub Search()
    Dim c As Range
    Set c = Range("A1:C5")
    
    Dim values As Variant
    values = c.value ' (1 To Rows.Count, 1 To Columns.Count) の二次元配列で値を取得
    
    Dim formatCells As Range ' 書式を設定するための条件に一致したセル
    
    Dim row As Long
    Dim column As Long
    
    ' セルを Z 方向に検索
    For row = 1 To c.Rows.Count
        For column = 1 To c.Columns.Count
            
            ' 複数条件
            Dim value As Variant
            value = values(row, column)
            If Not (value = 条件1 And value = 条件2) Then  ' And
                GoTo Continue
            End If
            If Not (value = 条件1 Or value = 条件2) Then ' Or
                GoTo Continue
            End If
    
            ' 複数条件に一致している
            values(row, column) = "一致" ' 値を編集
            
            ' 書式を編集するためにセルを収集
            If formatCells Is Nothing Then
                Set formatCells = c.Cells(row, column)
            Else
                Set formatCells = Union(formatCells, c.Cells(row, column))
            End If
            
Continue:
        Next
    Next
    
    ' 値をまとめて設定
    c.value = values
    
    ' 書式をまとめて設定
    If Not formatCells Is Nothing Then
        With formatCells
            .Font.Color = RGB(255, 0, 0)
            .Interior.ColorIndex = 35
        End With
    End If
End Sub

Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=2, Criteria1:= _
        "=*田*", Operator:=xlAnd
End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=2, Criteria1:= _
        "=*?田*", Operator:=xlAnd
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=7, Criteria1:= _
        "A"
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=4, Criteria1:= _
        ">=24", Operator:=xlAnd
End Sub

Public Function TableArray_5(T As ListObject) As Variant
  Dim buf1 As Variant    '←テーブル全体のデータ
  Dim buf2 As Variant    '←戻り値とする一時的な配列
  Dim i As Long            '←ｶｳﾝﾀ変数（配列の行位置）
  Dim j As Long            '←ｶｳﾝﾀ変数（配列の列位置）
  Dim k As Long            'テーブルのデータ行＋タイトル行の行数
  Dim CellsCnt As Long    '←絞り込みﾃﾞｰﾀのｾﾙ個数
  Dim ColCnt As Long      '←ﾃｰﾌﾞﾙの列数
  buf1 = T.Range
  CellsCnt = T.Range.SpecialCells(xlCellTypeVisible).Count
  ColCnt = UBound(buf1, 2)
  ReDim buf2(1 To (CellsCnt / ColCnt), 1 To ColCnt)
  For k = 1 To UBound(buf1, 1)
    If T.Range.Rows(k).Hidden = False Then
      i = i + 1
      For j = 1 To ColCnt
        buf2(i, j) = buf1(k, j)
      Next j
    End If
  Next k
  TableArray_5 = buf2
End Function

'==========　?(1)　テーブルの絞り込み　============
Sub TableFilter(T As ListObject, Col As Variant, _
                Optional C1 As Variant, Optional Ope As XlAutoFilterOperator, Optional C2 As Variant)
  '// T　　：操作対象のListObjectオブジェクト
  '// col　：絞り込み列。列名(文字列)でも列位置(整数)でもOK
  '// C1　：Critical1（文字列、数値、配列）
  '// Ope：Operator（数値　1～34）
  '// C2　：Critical2（文字列、数値、配列）
  Dim Param As Integer      '←指定したパラメータの組み合わせ値(0～7)
  If Not IsMissing(C1) Then Param = Param + 4
  If Ope >= 1 And Ope <= 34 Then Param = Param + 2
  If Not IsMissing(C2) Then Param = Param + 1
  On Error Resume Next
    Select Case Param
      Case 0
        T.Range.AutoFilter Field:=T.ListColumns(Col).index
      Case 3
        T.Range.AutoFilter Field:=T.ListColumns(Col).index, _
                              Operator:=Ope, Criteria2:=C2
      Case 4
        T.Range.AutoFilter Field:=T.ListColumns(Col).index, _
                              Criteria1:=C1
      Case 6
        T.Range.AutoFilter Field:=T.ListColumns(Col).index, _
                              Criteria1:=C1, Operator:=Ope
      Case 7
        T.Range.AutoFilter Field:=T.ListColumns(Col).index, _
                              Criteria1:=C1, Operator:=Ope, Criteria2:=C2
      Case Else
        MsgBox "引数の指定が間違っています"
    End Select
    If Not Err.Number = 0 Then MsgBox Err.Description
  On Error GoTo 0
End Sub

'==========　?(2)　プロシージャへの指示（文字列）　============
Sub AutoFilter_exec_String()
  'No.1 単一文字列を絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=道")
  'No.2 単一文字列を含む絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=*?道")
  'No.3 複数文字列をORで絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, Array("東海道", "東北道"), xlFilterValues)
  'No.4 2つの文字列を絞り込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "東海道", xlOr, "東北道")
  'No.5 空白以外を絞り込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "<>")
  'No.6 空白を絞り込み
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=")
  ' 絞り込み解除
  Call TableFilter(ActiveSheet.ListObjects(1), 2)
End Sub

'==========　?(3)　プロシージャへの指示（数値列）　============
Sub AutoFilter_exec_Numeric()
  'No.7 数値の1条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 1, ">5")
  'No.8 数値の複数値をORで絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 1, Array("3", "5"), 7)
  'No.9 数値の2条件の絞り込み
  Call TableFilter(ActiveSheet.ListObjects(1), 1, ">5", xlAnd, "<8")
  'No.10 数値の上位Top10(項目)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "3", xlTop10Items)
  'No.11 数値の下位Top10(項目)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "3", xlBottom10Items)
  'No.12 数値の上位Top10(％)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "20", xlTop10Percent)
  'No.13 数値の下位Top10(％)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "20", xlBottom10Percent)
  'No.14 数値の平均より上
  Call TableFilter(ActiveSheet.ListObjects(1), 1, xlFilterAboveAverage, xlFilterDynamic)
  'No.15 数値の平均より下
  Call TableFilter(ActiveSheet.ListObjects(1), 1, xlFilterBelowAverage, xlFilterDynamic)
  ' 絞り込み解除
  Call TableFilter(ActiveSheet.ListObjects(1), 1)
End Sub

'==========　?(4)　プロシージャへの指示（日付列）　============
Sub AutoFilter_exec_Date()
  'No.16 日付の1条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">2022/8/20")
  'No.17 日付の1条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & Format(CDate("2022/8/20"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal))
  'No.18 日付の1条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & CLng(CDate("2022/8/20")))
  'No.19 複数日付をORで絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, Array("2022/8/20", "2022/8/22"), xlFilterValues)
  'No.20 複数日付をORで絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, , xlFilterValues, Array(2, "2022/8/20", 2, "2022/8/22"))
  'No.21 日付の2条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">2022/8/20", xlAnd, "<2022/8/22")
  'No.22 日付の2条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & Format(CDate("2022/8/20"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal), xlAnd, "<" & Format(CDate("2022/8/22"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal))
  'No.23 日付の2条件の絞込み
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & CLng(CDate("2022/8/20")), xlAnd, "<" & CLng(CDate("2022/8/22")))
  'No.24 既定の日付（今日を絞り込み）
  Call TableFilter(ActiveSheet.ListObjects(1), 3, xlFilterToday, xlFilterDynamic)
  ' 絞り込み解除
  Call TableFilter(ActiveSheet.ListObjects(1), 3)
End Sub

'==========　?(1)　１列ずつ解除　============
Sub TableFilterOff_01(T As ListObject)
  Dim i As Integer      '←カウンタ変数（列位置）
  If T.ShowAutoFilter = False Then
    T.ShowAutoFilter = True
    Exit Sub
  End If
  For i = 1 To T.ListColumns.Count
    If T.AutoFilter.Filters(i).On = True Then
      T.Range.AutoFilter Field:=i
    End If
  Next i
End Sub

'==========　?(2)　絞り込まれていなくても1列ずつ解除　============
Sub TableFilterOff_02(T As ListObject)
  Dim i As Integer      '←カウンタ変数（列位置）
  For i = 1 To T.ListColumns.Count
    T.Range.AutoFilter Field:=i
  Next i
End Sub

'==========　?(3)　AutoFilterメソッドでボタン消去・再表示　============
Sub TableFilterOff_03(T As ListObject)
  If Not T.AutoFilter Is Nothing Then
    T.Range.AutoFilter
  End If
  T.Range.AutoFilter
End Sub

'==========　?(4)　ShowAutoFilterプロパティでボタン消去・再表示　============
Sub TableFilterOff_04(T As ListObject)
  If T.ShowAutoFilter = True Then
    T.ShowAutoFilter = False
  End If
  T.ShowAutoFilter = True
End Sub

'==========　?(5)　ShowAllDataで全データを表示　============
Sub TableFilterOff_05(T As ListObject)
  If T.AutoFilter Is Nothing Then
    T.Range.AutoFilter
    Exit Sub
  End If
  If T.AutoFilter.FilterMode = True Then
    T.AutoFilter.ShowAllData
  End If
End Sub

'==========　?(5)　テーブルに行を挿入 ４　============
Sub TableInsert_4(T As ListObject, arrayData As Variant)
  With T.ListRows.Add
    .Range = arrayData
  End With
End Sub

'==========　?(7)　テーブルの最終行の下にデータを追加 ２　============
Sub TableInsert_6(T As ListObject, arrayData As Variant)
  T.HeaderRowRange.Offset(T.ListRows.Count + 1, 0) = arrayData
End Sub

'==========　?(2)　全行を調べ可視行の場合にデータ書き換え　============
Sub TableUpdate_1(T As ListObject, Col As Variant, uniData As Variant)
  Dim i As Long      '←テーブルのデータ行数
  For i = 1 To T.ListRows.Count
    If T.DataBodyRange.Rows(i).Hidden = False Then
      T.ListColumns(Col).DataBodyRange(i) = uniData
    End If
  Next i
End Sub

'==========　?(3)　絞り込み行のみをデータ変更　============
Sub TableUpdate_2(T As ListObject, Col As Variant, uniData As Variant)
  Dim r As Range      '←テーブル内の可視行×指定列のセル範囲
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  For Each r In T.ListColumns(Col).DataBodyRange.SpecialCells(xlCellTypeVisible)
    r = uniData
  Next r
End Sub

'==========　?(4)　テーブルの列全体に対してデータ変更　============
Sub TableUpdate_3(T As ListObject, Col As Variant, uniData As Variant)
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  T.ListColumns(Col).DataBodyRange = uniData
End Sub

'==========　?(2)　抽出行のセル範囲を取得し、絞り込み解除後に削除　============
Sub TableDel_1(T As ListObject)
  Dim r As Range    '←可視行のセル範囲（複数行に渡ることもあり）
  On Error Resume Next
    Set r = T.DataBodyRange.SpecialCells(xlCellTypeVisible)
  On Error GoTo 0
  T.ShowAutoFilter = False
  T.ShowAutoFilter = True
  If Not r Is Nothing Then r.Delete
End Sub

'==========　?(2)　絞り込み後の可視行をそのまま行削除　============
Sub TableDel_2(T As ListObject)
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  T.DataBodyRange.EntireRow.Delete
End Sub


