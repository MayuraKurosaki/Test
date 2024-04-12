Attribute VB_Name = "Module1"
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

Sub search()
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
