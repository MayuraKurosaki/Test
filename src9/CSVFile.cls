VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
' データ構造
' [Root] Type:Object(Dictionary), Name:N/A, ChildCount:row + 1
'    +-[Item(0)] Type:Object(Dictionary), Name:Header, ChildCount:MaxColumn + 1
'    |    +-[Item(0)(0)] Type:String, Name:1, Value:HeaderField(1)
'    |    +-[Item(0)(1)] Type:String, Name:2, Value:HeaderField(2)
'    |    .
'    |    .
'    |    +-[Item(0)(col)] Type:String, Name:[col+1], Value:HeaderField(col+1)
'    |
'    +-[Item(1)] Type:Object(Dictionary), Name:Content1, ChildCount:FieldCount + 1
'    |    +-[Item(1)(0)] Type:String, Name:HeaderField(1), Value:FieldData(1)(1)
'    |    +-[Item(1)(1)] Type:String, Name:HeaderField(2), Value:FieldData(1)(2)
'    |    .
'    |    .
'    +-[Item(2)] Type:Object(Dictionary), Name:Content2, ChildCount:FieldCount + 1
'    |    +-[Item(2)(0)] Type:String, Name:HeaderField(1), Value:FieldData(2)(1)
'    |    +-[Item(2)(1)] Type:String, Name:HeaderField(2), Value:FieldData(2)(2)
'    |    .
'    |    .
'    +-[Item(row)] Type:Object(Dictionary), Name:Content[row], ChildCount:FieldCount + 1
'    |    +-[Item(row)(0)] Type:String, Name:HeaderField(1), Value:FieldData(row)(1)
'    |    +-[Item(row)(1)] Type:String, Name:HeaderField(2), Value:FieldData(row)(2)
'    |    .
'    |    .
'    |    +-[Item(row)(col)] Type:String, Name:HeaderField(col+1), Value:FieldData(row)(col+1)

Private Enum State
    ST_HEADER
    ST_CONTENT
    ST_EOT          ' End Of Token
    ST_EOL          ' End Of Line
    ST_EOF          ' End Of File
    ST_UNDETERMINED
End Enum

Private containsHeader_ As Boolean
Private lineSep_ As String
Private delimiter_ As String
Private maxColumn_ As Long
Private sQuote_ As Long
Private dQuote_ As Long
Private rBracket_ As Long
Private sBracket_ As Long
Private cBracket_ As Long
Private countSQuate_ As Long
Private countDQuate_ As Long
Private countRBracket_ As Long
Private countSBracket_ As Long
Private countCBracket_ As Long
Private dicHeader As Dictionary
Private row_ As Long
Private col_ As Long
Private recentMarkStack_ As Stack    '直近のクォーテーションマークまたは開き括弧
    

Private root_ As DataNode   'ファイル内容を保持するルートノード

Private Sub Class_Initialize()
    containsHeader_ = False
    lineSep_ = VBA.vbCrLf
    delimiter_ = ","
    Set root_ = CreateInstance(New DataNode, NK_ROOT)
    Set recentMarkStack_ = New Stack
    Set dicHeader = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set root_ = Nothing
    Set recentMarkStack_ = Nothing
    Set dicHeader = Nothing
End Sub

Private Property Let ContainsHeader(aContainsHeader As Boolean)
    containsHeader_ = aContainsHeader
End Property

Public Property Get ContainsHeader() As Boolean
    ContainsHeader = containsHeader_
End Property

Private Property Let LineSeparator(aLineSep As String)
    lineSep_ = aLineSep
End Property

Private Property Get LineSeparator() As String
    LineSeparator = lineSep_
End Property

Private Property Let Delimiter(aDelimiter As String)
    delimiter_ = aDelimiter
End Property

Private Property Get Delimiter() As String
    Delimiter = delimiter_
End Property

Public Property Let Quote(aQuote As String)
    sQuote_ = VBA.CLng(aQuote Like "*'*") * -1
    dQuote_ = VBA.CLng(aQuote Like "*""*") * -1
End Property

Public Property Get Quote() As String
    Quote = VBA.IIf(sQuote_, "'", "") & VBA.IIf(dQuote_, """", "")
End Property

Public Property Let Bracket(aBracket As String)
    rBracket_ = VBA.CLng(aBracket Like "*(*") * -1
    sBracket_ = VBA.CLng(aBracket Like "*[[]*") * -1
    cBracket_ = VBA.CLng(aBracket Like "*{*") * -1
End Property

Public Property Get Bracket() As String
    Bracket = VBA.IIf(rBracket_, "(", "") & VBA.IIf(sBracket_, "[", "") & VBA.IIf(cBracket_, "{", "")
End Property

Private Property Let MaxColumn(aColumn As Long)
    maxColumn_ = aColumn
End Property

Private Property Get MaxColumn() As Long
    MaxColumn = maxColumn_
End Property

'Public Property Get ContentType(ByRef aLine As Long) As State
'    If root_(aLine - 1).Name = "Header" Then
'        ContentType = ST_HEADER
'    ElseIf VBA.InStr(0, root_(aLine - 1).Name, "Content") <> 0 Then
'        ContentType = ST_CONTENT
'    Else
'        ContentType = ST_UNKNOWN
'    End If
'End Property

Public Property Get LineCount() As Long
    LineCount = root_.ChildCount
End Property

Public Property Get Lines() As Collection
    Set Lines = root_.Children
End Property

Public Property Get fieldCount(ByRef aNameOrLine As Variant) As Long
    Select Case VBA.VarType(aNameOrLine)
        Case VBA.vbString
            fieldCount = root_.GetNode(VBA.CStr(aNameOrLine)).ChildCount
        Case VBA.vbInteger, VBA.vbLong
            fieldCount = root_.GetNodeByIndex(VBA.CLng(aNameOrLine)).ChildCount
        Case Else
            fieldCount = 0
    End Select
End Property

Public Property Let Header(ByRef aHeaderArray As Variant)
    Dim i As Long
    For i = 0 To UBound(aHeaderArray)
        Call PutLineData(dicHeader, VBA.CLng(i), VBA.CStr(aHeaderArray(i)))
    Next i
End Property

Public Function ReadLine(ByRef aNameOrLine As Variant) As Dictionary
    Set ReadLine = Nothing
    
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrLine)
        Case VBA.vbString
            Set node_ = root_.GetNode(VBA.CStr(aNameOrLine))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = root_.GetNodeByIndex(VBA.CLng(aNameOrLine))
        Case Else
            Set node_ = Nothing
    End Select
    
    If node_ Is Nothing Then Exit Function
    
    Set ReadLine = New Dictionary
    
    Dim child_ As Object    'DataNode
    For Each child_ In node_.Children
        Call ReadLine.Add(child_.Name, child_.Value)
    Next child_
    
    Set child_ = Nothing
    Set node_ = Nothing
End Function

'1行分のデータをParentノードにObject(Dictionary)ノードとして追加する
Private Sub AddLineNode(ByRef aName As String, ByRef dicLine As Dictionary)
    Dim lineNode As DataNode
    
    Set lineNode = CreateInstance(New DataNode, NK_OBJECT, aName)
    
    Dim key_ As Variant
    
    For Each key_ In dicLine.Keys
        Call lineNode.AddField(VBA.CStr(key_), dicLine(key_))
    Next
    
    Call root_.AddNode(lineNode)
    Set lineNode = Nothing
    
    Call dicLine.RemoveAll
    
    row_ = row_ + 1
    col_ = 1
End Sub

'1フィールドごとに辞書に登録
Private Sub PutLineData(ByRef dicLine As Dictionary, ByRef aKey As Variant, ByRef fieldData As String)
    '「""」を「"」で置換
    fieldData = VBA.Replace(fieldData, """""", """")
    '前後の「"」を削除
    If VBA.Left$(fieldData, 1) = """" And VBA.Right$(fieldData, 1) = """" Then
        If VBA.Len(fieldData) <= 2 Then
            fieldData = ""
        Else
            fieldData = VBA.Mid$(fieldData, 2, VBA.Len(fieldData) - 2)
        End If
    End If
    
    Call dicLine.Add(aKey, fieldData)
    col_ = col_ + 1
    fieldData = ""
End Sub

Public Sub Parse(ByRef aFile As Variant, Optional ByRef aContainsHeader As Boolean = False, Optional ByVal aLineSep As String = VBA.vbCrLf, Optional ByVal aDelimiter As String = ",", Optional ByVal aLineNum As Boolean = False)
    If VBA.VarType(aFile) <> VBA.vbString Then Exit Sub
    
    Dim contents As String
    Dim stream_ As TextStream
    
    Set stream_ = CreateInstance(New TextStream, aFile)
    contents = stream_.LoadFile
    
    ContainsHeader = aContainsHeader
    LineSeparator = aLineSep
    Delimiter = aDelimiter
    
    Dim fieldData As String:    fieldData = ""      '1フィールド分データ
    Dim dicLineData As Dictionary   '1行分のデータ
    Dim state_ As State
    
    Dim key_() As Variant
    
    row_ = 0
    col_ = 1

    Set dicLineData = New Dictionary
    
    MaxColumn = 0
    
    Dim i As Long   '文字数
    i = 1
    
    If ContainsHeader Then
        'ファイル1行目がヘッダの場合
        If aLineNum Then Call PutLineData(dicHeader, VBA.CLng(col_), VBA.CStr(row_))
        Do Until (state_ = ST_EOL)
            state_ = DivideIntoField(contents, i, fieldData)
            Call PutLineData(dicHeader, VBA.CLng(col_), fieldData)
            ReDim Preserve key_(col_ - 1)
            key_(col_ - 1) = dicHeader(col_ - 1)
        Loop
    End If
    
    Call AddLineNode("Header", dicHeader)
    
    Do
        If aLineNum Then Call PutLineData(dicLineData, VBA.CLng(col_), VBA.CStr(row_))
        Do
            state_ = DivideIntoField(contents, i, fieldData)
            '最大列数を更新
            If col_ > MaxColumn Then
                MaxColumn = col_
                ReDim Preserve key_(col_ - 1)
                key_(col_ - 1) = VBA.CLng(col_)
            End If
            Call PutLineData(dicLineData, key_(col_ - 1), fieldData)
            If state_ = ST_EOF Then Exit Do
        Loop While (state_ <> ST_EOL)
        
        Call AddLineNode("Content" & row_, dicLineData)
    Loop While (state_ <> ST_EOF)
    
    Set stream_ = Nothing
    Set dicLineData = Nothing
End Sub

Private Function DivideIntoField(ByRef aContents As String, ByRef i As Long, ByRef aOutField As String) As State
    DivideIntoField = ST_UNDETERMINED
    
    countSQuate_ = 0     'シングルクォーテーション数
    countDQuate_ = 0     'ダブルクォーテーション数
    countRBracket_ = 0   '「(」の数
    countSBracket_ = 0   '「[」の数
    countCBracket_ = 0   '「{」の数
    
    Dim charCount As Long
    Dim c As String
    
    For charCount = i To VBA.Len(aContents)
        c = VBA.Mid$(aContents, charCount, 1)
        Select Case c
            Case vbLf, vbCr
                If countDQuate_ = 0 Then   '「"」で括られた改行コードは文字として扱う
                    '改行候補としてさらに判定
                    If i + 1 < VBA.Len(aContents) Then '改行のCrLfはCrで改行判定済なので無視する
                        If VBA.Mid$(aContents, charCount, 2) = vbCrLf Then charCount = charCount + 1
                    End If
                    '改行としてフィールドの終了、行の終了を返す
                    i = charCount + 1
                    DivideIntoField = ST_EOL
                    Exit Function
                End If
            Case ",", vbTab     '引用符または括弧で括られている場合は文字として扱う
                If countSQuate_ + countDQuate_ + _
                   countRBracket_ + countSBracket_ + countCBracket_ = 0 Then
               
                    'トークンの区切りとしてトークンの終了を返す
                    DivideIntoField = ST_EOT
                    i = charCount + 1
                    Exit Function
                End If
            Case "'"
                If sQuote_ Then
                    If i + 1 < VBA.Len(aContents) Then
                        If VBA.Mid$(aContents, charCount, 2) = "''" Then
                            aOutField = aOutField & c
                            charCount = charCount + 1
                        Else
                            Call IsEnclosed(c)
                        End If
                    End If
                End If
            Case """"
                If dQuote_ Then
                    If i + 1 < VBA.Len(aContents) Then
                        If VBA.Mid$(aContents, charCount, 2) = """""" Then
                            aOutField = aOutField & c
                            charCount = charCount + 1
                        Else
                            Call IsEnclosed(c)
                        End If
                    End If
                End If
            Case "(", ")"
                If rBracket_ Then Call IsEnclosed(c)
            Case "[", "]"
                If sBracket_ Then Call IsEnclosed(c)
            Case "{", "}"
                If cBracket_ Then Call IsEnclosed(c)
        End Select
        aOutField = aOutField & c
CONTINUE:
    Next

    '最終行の最終列の処理
    DivideIntoField = ST_EOF
End Function

'引用符、括弧の処理ルール
'「"」に囲まれるものは「"」以外全て文字とする。「"」自身を文字として含めるときは「""」のように2つ続ける
'「'」に囲まれるものは「'」以外全て文字とする。「'」自身を文字として含めるときは「'」のように2つ続ける(SQL対応)
'「(」と「)」、「[」と「]」、「{」と「}」はそれぞれ対応し、それぞれのペアで入れ子関係に対応する。
'　　・内側の括弧が閉じられる前に外側の括弧を閉じる記述がある場合はエラーとする。
'　　・同一の括弧の種類で開閉の個数が合わない場合は以降の文字列解析が正しく行われない。
Private Function IsEnclosed(ByRef c As String) As Boolean
    IsEnclosed = False
    
    Dim recentMark_ As String
    recentMark_ = recentMarkStack_.Pop
    
    Select Case c
        Case "'"
            If recentMark_ = """" Then
                '「"」で括られている場合は文字として扱う
                recentMarkStack_.Push recentMark_
            ElseIf recentMark_ = "'" Then
                countSQuate_ = countSQuate_ - 1
            Else
                countSQuate_ = countSQuate_ + 1
                recentMarkStack_.Push recentMark_
                recentMarkStack_.Push c
            End If
        Case """"
            If recentMark_ = "'" Then
                '「'」で括られている場合は文字として扱う
                recentMarkStack_.Push recentMark_
            ElseIf recentMark_ = """" Then
                countDQuate_ = countDQuate_ - 1
            Else
                countDQuate_ = countDQuate_ + 1
                recentMarkStack_.Push recentMark_
                recentMarkStack_.Push c
            End If
        Case "("
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                '「"」「'」で括られている場合は文字として扱う
                recentMarkStack_.Push recentMark_
            Else
                countRBracket_ = countRBracket_ + 1
                recentMarkStack_.Push recentMark_
                recentMarkStack_.Push c
            End If
        Case ")"
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                recentMarkStack_.Push recentMark_
            ElseIf recentMark_ = "(" Then
                countRBracket_ = countRBracket_ - 1
            Else
                '括弧の対応が正しくない
            End If
        Case "["
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                recentMarkStack_.Push recentMark_
            Else
                countSBracket_ = countSBracket_ + 1
                recentMarkStack_.Push recentMark_
                recentMarkStack_.Push c
            End If
        Case "]"
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                recentMarkStack_.Push recentMark_
            ElseIf recentMark_ = "[" Then
                countSBracket_ = countSBracket_ - 1
            Else
                '括弧の対応が正しくない
            End If
        Case "{"
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                recentMarkStack_.Push recentMark_
            Else
                countCBracket_ = countCBracket_ + 1
                recentMarkStack_.Push recentMark_
                recentMarkStack_.Push c
            End If
        Case "}"
            If (recentMark_ = """") Or _
               (recentMark_ = "'") Then
                recentMarkStack_.Push recentMark_
            ElseIf recentMark_ = "{" Then
                countCBracket_ = countCBracket_ - 1
            Else
                '括弧の対応が正しくない
            End If
    End Select
    
    If countSQuate_ + countDQuate_ + _
       countRBracket_ + countSBracket_ + countCBracket_ <> 0 Then
        
        IsEnclosed = True
    End If
End Function

Public Function WriteToString() As String
    Dim content As StringBuilder
    content = New StringBuilder

    Dim first As DataNode
    Set first = root_.Children(1)

    Dim index As Long
    index = 0
    
    Dim header_ As VBA.Collection       ' List<string>
    Set header_ = New VBA.Collection
    
    Dim field As DataNode
    For Each field In first.Children
        If (index > 0) Then Call content.Append(",")
        Call content.Append(field.Name)
        Call header_.Add(field.Name)
        index = index + 1
    Next
    
    Call content.AppendLine

    Dim Item As DataNode
    For Each Item In root_.Children
        index = 0
        Dim fieldName As Variant
        For Each fieldName In header_
            field = Item.GetNode(VBA.CStr(fieldName))

            If (index > 0) Then Call content.Append(",")

            If (Not field Is Nothing) Then
                Dim escape As Boolean
                escape = (VBA.InStr(field.Value, ",") Or VBA.InStr(field.Value, VBA.vbLf))
'                escape = (VBA.InStr(field.Value, ",") Or VBA.InStr(field.Value, "\n"))

                If escape Then Call content.Append(VBA.Chr$(34))
                Call content.Append(field.Value)
                If escape Then Call content.Append(VBA.Chr$(34))
            End If

            index = index + 1
        Next

        Call content.AppendLine
    Next

    WriteToString = content.ToString()
End Function
