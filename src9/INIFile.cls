VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
' データ構造
' [Root] Type:Object(Dictionary), Name:N/A, ChildCount:SectionCount
'    +-[Item(0)] Type:Object(Dictionary), Name:[Section1], ChildCount:KeyCount
'    |    +-[Item(0)(0)] Type:String, Name:[Key1], Value:[Value1]
'    |    +-[Item(0)(1)] Type:String, Name:[Key2], Value:[Value2]
'    |    .
'    |    .
'    |    +-[Item(0)(j)] Type:String, Name:[Key(j+1)], Value:[Value(j+1)]
'    |
'    +-[Item(1)] Type:Object(Dictionary), Name:[Section2], ChildCount:KeyCount
'    |    +-[Item(1)(0)] Type:String, Name:[Key1], Value:[Value1]
'    |    +-[Item(1)(1)] Type:String, Name:[Key2], Value:[Value2]
'    |    .
'    |    .
'    +-[Item(i)] Type:Object(Dictionary), Name:[Section(i+1)], ChildCount:KeyCount
'    |    +-[Item(i)(0)] Type:String, Name:[Key1], Value:[Value1]
'    |    +-[Item(i)(1)] Type:String, Name:[Key1], Value:[Value1]
'    |    .
'    |    .
'    |    +-[Item(i)(j)] Type:String, Name:[Key(j+1)], Value:[Value(j+1)]

Private Enum State
    ST_SECTION
    ST_KEY
    ST_VALUE
    ST_COMMENT
    ST_UNDETERMINED
End Enum

Private root_ As DataNode   'ファイル内容を保持するルートノード

Private Sub Class_Initialize()
    Set root_ = CreateInstance(New DataNode, NK_ROOT)
End Sub

Private Sub Class_Terminate()
    Set root_ = Nothing
End Sub

Public Property Get SectionCount() As Long
    SectionCount = root_.ChildCount
End Property

Public Property Get Sections() As Collection
    Set Sections = root_.Children
End Property

Public Function ReadString(aKey As String, Optional aSection As String = "", Optional aDefault As String = "") As String
    If SectionCount = 0 Then Exit Function
    ReadString = root_.FindNode(aSection).GetString(aKey)
End Function

Public Function ReadInteger(aKey As String, Optional aSection As String = "", Optional aDefault As Variant = CLng("0")) As Long
    If SectionCount = 0 Then Exit Function
    ReadInteger = root_.FindNode(aSection).GetInt32(aKey)
End Function

Public Function ReadDouble(aKey As String, Optional aSection As String = "", Optional aDefault As Variant = CDbl("0")) As Double
    If SectionCount = 0 Then Exit Function
    ReadDouble = root_.FindNode(aSection).GetDouble(aKey)
End Function

Public Function ReadBool(aKey As String, Optional aSection As String = "", Optional aDefault As Boolean = False) As Boolean
    If SectionCount = 0 Then Exit Function
    ReadBool = root_.FindNode(aSection).GetBool(aKey)
End Function

Public Function ReadSection(Optional aSection As String) As Dictionary
    Set ReadSection = Nothing
    
    Dim node_ As DataNode
    Set node_ = root_.FindNode(aSection)
    If node_ Is Nothing Then GoTo FINALY
    
    Set ReadSection = New Dictionary
    
    Dim child_ As Object
    For Each child_ In node_.Children
        Call ReadSection.Add(child_.Name, child_.Value)
    Next child_
    
    Set child_ = Nothing
    
FINALY:
    Set node_ = Nothing
End Function

'セクションデータをセクションノードとして追加する
Private Sub AddSectionNode(ByRef aSectionName As String, ByRef dicSection As Dictionary)
    Dim sectionNode As DataNode
    
    Set sectionNode = CreateInstance(New DataNode, NK_OBJECT, aSectionName)
    
    Dim key_ As Variant
    For Each key_ In dicSection.Keys
        Call sectionNode.AddField(VBA.CStr(key_), dicSection(key_))
    Next
    
    Call root_.AddNode(sectionNode)
    Set sectionNode = Nothing
    
    'Sectionデータ初期化
    Call dicSection.RemoveAll
    aSectionName = ""
End Sub

Public Sub Parse(ByRef aFile As Variant)
    If VBA.VarType(aFile) <> VBA.vbString Then Exit Sub
    
    Dim contents As String
    Dim stream_ As TextStream
    
    Set stream_ = CreateInstance(New TextStream, aFile)
    contents = stream_.LoadFile
    
    Dim i As Long   '文字数
    Dim countQuate As Long:     countQuate = 0      'ダブルクォーテーション数
    Dim countBracket As Long:   countBracket = 0    '「[」の数
    Dim dicSection As Dictionary                    '1セクション分のデータ
    
    Dim sectionName_ As String: sectionName_ = ""
    Dim sectionCount_ As Long:  sectionCount_ = 0
    Dim key_ As String:         key_ = ""
    Dim value_ As String:       value_ = ""
    Dim strTmp As String:       strTmp = ""
    Dim state_ As State:        state_ = ST_UNDETERMINED
    
    Set dicSection = New Dictionary
    
    For i = 1 To VBA.Len(contents)
        Select Case VBA.Mid$(contents, i, 1)
            Case vbLf, vbCr
                '改行コードCrLfの判定(改行処理はCr判定時に実施済みのため読み飛ばす)
                If i > 1 Then
                    If VBA.Mid$(contents, i - 1, 2) = vbCrLf Then
                        'i = i + 1
                        GoTo CONTINUE
                    End If
                End If
                
                'Value処理中のみ有効行となる
                '---Key=ValueペアをSectionデータに追加---
                If state_ = ST_VALUE Then
                    value_ = strTmp
                    strTmp = ""
                    Call dicSection.Add(key_, value_)
                End If
                countQuate = 0
                countBracket = 0
                state_ = ST_UNDETERMINED
                GoTo CONTINUE
            Case " ", vbTab '空白文字(行頭でなく、かつコメントでもない場合のみ読み込む)
                Select Case state_
                    Case ST_COMMENT:        GoTo CONTINUE
                    Case ST_UNDETERMINED:   GoTo CONTINUE
                End Select
            Case ";"
                '次の場合はコメントとする
                '・最初の非空白文字の場合
                '・「"」で括られていない場合(「"」が偶数の場合)
                Select Case state_
                    Case ST_COMMENT: GoTo CONTINUE
                    Case ST_SECTION
                        'Section名の定義が終わっていないので無効行とする
                        countQuate = 0
                        countBracket = 0
                        state_ = ST_COMMENT
                        GoTo CONTINUE
                    Case ST_VALUE
                        'Key名の途中の場合はKey名の一部として、
                        'Valueの場合、「"」で括られている場合のみValueの一部として扱う
                        If countQuate Mod 2 = 0 Then    '「"」が偶数ならコメント開始、奇数ならValueの一部
                            value_ = strTmp
                            strTmp = ""
                            '---登録処理---
                            Call dicSection.Add(key_, value_)
                            countQuate = 0
                            countBracket = 0
                            state_ = ST_COMMENT
                            GoTo CONTINUE
                        End If
                    Case ST_UNDETERMINED
                        '行頭の「;」はコメント行
                        state_ = ST_COMMENT
                        GoTo CONTINUE
                End Select
            Case "="
                Select Case state_
                    Case ST_COMMENT: GoTo CONTINUE
                    Case ST_KEY
                        key_ = strTmp
                        strTmp = ""
                        countQuate = 0
                        countBracket = 0
                        state_ = ST_VALUE
                        GoTo CONTINUE
                    Case ST_UNDETERMINED
                        '行頭の「=」は無効行とする(コメント行と同じ処理)
                        state_ = ST_COMMENT
                        GoTo CONTINUE
                End Select
            Case """" '「"」のカウントをとる
                Select Case state_
                    Case ST_COMMENT:        GoTo CONTINUE
                    Case ST_UNDETERMINED:   GoTo CONTINUE
                End Select
                countQuate = countQuate + 1
            Case "["
                Select Case state_
                    Case ST_COMMENT: GoTo CONTINUE
                    Case ST_SECTION
                        '「[」をSection名の一部として読み込む
                        countBracket = countBracket + 1
                    Case ST_UNDETERMINED
                        '行頭に「[」が現れた場合:Section(「[」は読まずに次の文字へ)
                        countBracket = countBracket + 1
                        '新しいSectionの始まり
                        '---現Sectionの内容をNodeに追加---
                        If sectionName_ <> "" Then Call AddSectionNode(sectionName_, dicSection)
                        state_ = ST_SECTION
                        GoTo CONTINUE
                End Select
            Case "]"
                Select Case state_
                    Case ST_COMMENT: GoTo CONTINUE
                    Case ST_SECTION
                        countBracket = countBracket - 1
                        If countBracket = 0 Then
                            'Section名の終わり
                            '「]」は読まずにこれまでの文字列をSection名とする
                            sectionName_ = strTmp
                            strTmp = ""
                            countQuate = 0
                            countBracket = 0
                            '以降の記述は無視する(コメント行と同じ処理)
                            state_ = ST_COMMENT
                            GoTo CONTINUE
                        End If
                    Case ST_UNDETERMINED
                        '行頭に「]」が現れた場合:Key(「]」をKey名の一部として読み込む)
                        state_ = ST_KEY
                End Select
            Case Else
                Select Case state_
                    Case ST_COMMENT: GoTo CONTINUE
                    Case ST_UNDETERMINED
                        state_ = ST_KEY
                End Select
        End Select
        
        strTmp = strTmp & VBA.Mid$(contents, i, 1)
CONTINUE:
    Next

    '最終行の処理
    'ここまでのKey=ValueペアをSectionデータに追加してからSectionをNodeに追加
    If state_ <> ST_SECTION Then
        If state_ = ST_KEY Then
            value_ = strTmp
            strTmp = ""
            '---登録処理---
            Call dicSection.Add(key_, value_)
        End If
        If sectionName_ <> "" Then Call AddSectionNode(sectionName_, dicSection)
    End If
    
    Set stream_ = Nothing
    Set dicSection = Nothing
End Sub
