Attribute VB_Name = "Search"
Option Explicit

Sub ExtractData()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("ExtractedData")
    
    '�e�V�[�g�̍ŏI�s���擾
    Dim cmax1 As Long, cmax2 As Long
    cmax1 = ws1.Range("A65536").End(xlUp).row
    cmax2 = ws2.Range("A65536").End(xlUp).row
    
    '�f�[�^�����Z�b�g
    ws2.Range("B6:B7").ClearContents
    If Not cmax2 = 9 Then: ws2.Range("A10:E" & cmax2).ClearContents
    
    '�J�n���ƏI�������擾
    Dim startdate As Date, enddate As Date
    startdate = ws2.Range("B2").value
    enddate = ws2.Range("B3").value
    
    '�������擾
    Dim torihiki As String
    torihiki = ws2.Range("B4").value
    
    '�J�n���A�I�����A����悪�󗓂�����
    Dim flag(2) As Boolean ' Boolean��Default�l��False
    If startdate = 0 Then: flag(0) = True
    If enddate = 0 Then: flag(1) = True
    If torihiki = "" Then: flag(2) = True
    
    '�ϐ��̏�����
    Dim n As Long: n = 10
    Dim goukei As Long: goukei = 0
    Dim kensu As Long: kensu = 0
    
    '�����ɍ��v�����s�𒊏o
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
    
        '�����ɍ��v�����s�̃f�[�^�݂̂�Ώۂ��ĕ���
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
    values = c.value ' (1 To Rows.Count, 1 To Columns.Count) �̓񎟌��z��Œl���擾
    
    Dim formatCells As Range ' ������ݒ肷�邽�߂̏����Ɉ�v�����Z��
    
    Dim row As Long
    Dim column As Long
    
    ' �Z���� Z �����Ɍ���
    For row = 1 To c.Rows.Count
        For column = 1 To c.Columns.Count
            
            ' ��������
            Dim value As Variant
            value = values(row, column)
            If Not (value = ����1 And value = ����2) Then  ' And
                GoTo Continue
            End If
            If Not (value = ����1 Or value = ����2) Then ' Or
                GoTo Continue
            End If
    
            ' ���������Ɉ�v���Ă���
            values(row, column) = "��v" ' �l��ҏW
            
            ' ������ҏW���邽�߂ɃZ�������W
            If formatCells Is Nothing Then
                Set formatCells = c.Cells(row, column)
            Else
                Set formatCells = Union(formatCells, c.Cells(row, column))
            End If
            
Continue:
        Next
    Next
    
    ' �l���܂Ƃ߂Đݒ�
    c.value = values
    
    ' �������܂Ƃ߂Đݒ�
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
        "=*�c*", Operator:=xlAnd
End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=2, Criteria1:= _
        "=*?�c*", Operator:=xlAnd
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=7, Criteria1:= _
        "A"
    ActiveSheet.ListObjects("T_Dummy").Range.AutoFilter Field:=4, Criteria1:= _
        ">=24", Operator:=xlAnd
End Sub

Public Function TableArray_5(T As ListObject) As Variant
  Dim buf1 As Variant    '���e�[�u���S�̂̃f�[�^
  Dim buf2 As Variant    '���߂�l�Ƃ���ꎞ�I�Ȕz��
  Dim i As Long            '�������ϐ��i�z��̍s�ʒu�j
  Dim j As Long            '�������ϐ��i�z��̗�ʒu�j
  Dim k As Long            '�e�[�u���̃f�[�^�s�{�^�C�g���s�̍s��
  Dim CellsCnt As Long    '���i�荞���ް��ٌ̾�
  Dim ColCnt As Long      '��ð��ق̗�
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

'==========�@?(1)�@�e�[�u���̍i�荞�݁@============
Sub TableFilter(T As ListObject, Col As Variant, _
                Optional C1 As Variant, Optional Ope As XlAutoFilterOperator, Optional C2 As Variant)
  '// T�@�@�F����Ώۂ�ListObject�I�u�W�F�N�g
  '// col�@�F�i�荞�ݗ�B��(������)�ł���ʒu(����)�ł�OK
  '// C1�@�FCritical1�i������A���l�A�z��j
  '// Ope�FOperator�i���l�@1�`34�j
  '// C2�@�FCritical2�i������A���l�A�z��j
  Dim Param As Integer      '���w�肵���p�����[�^�̑g�ݍ��킹�l(0�`7)
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
        MsgBox "�����̎w�肪�Ԉ���Ă��܂�"
    End Select
    If Not Err.Number = 0 Then MsgBox Err.Description
  On Error GoTo 0
End Sub

'==========�@?(2)�@�v���V�[�W���ւ̎w���i������j�@============
Sub AutoFilter_exec_String()
  'No.1 �P�ꕶ������i����
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=��")
  'No.2 �P�ꕶ������܂ލi����
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=*?��")
  'No.3 �����������OR�ōi����
  Call TableFilter(ActiveSheet.ListObjects(1), 2, Array("���C��", "���k��"), xlFilterValues)
  'No.4 2�̕�������i�荞��
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "���C��", xlOr, "���k��")
  'No.5 �󔒈ȊO���i�荞��
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "<>")
  'No.6 �󔒂��i�荞��
  Call TableFilter(ActiveSheet.ListObjects(1), 2, "=")
  ' �i�荞�݉���
  Call TableFilter(ActiveSheet.ListObjects(1), 2)
End Sub

'==========�@?(3)�@�v���V�[�W���ւ̎w���i���l��j�@============
Sub AutoFilter_exec_Numeric()
  'No.7 ���l��1�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 1, ">5")
  'No.8 ���l�̕����l��OR�ōi����
  Call TableFilter(ActiveSheet.ListObjects(1), 1, Array("3", "5"), 7)
  'No.9 ���l��2�����̍i�荞��
  Call TableFilter(ActiveSheet.ListObjects(1), 1, ">5", xlAnd, "<8")
  'No.10 ���l�̏��Top10(����)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "3", xlTop10Items)
  'No.11 ���l�̉���Top10(����)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "3", xlBottom10Items)
  'No.12 ���l�̏��Top10(��)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "20", xlTop10Percent)
  'No.13 ���l�̉���Top10(��)
  Call TableFilter(ActiveSheet.ListObjects(1), 1, "20", xlBottom10Percent)
  'No.14 ���l�̕��ς���
  Call TableFilter(ActiveSheet.ListObjects(1), 1, xlFilterAboveAverage, xlFilterDynamic)
  'No.15 ���l�̕��ς�艺
  Call TableFilter(ActiveSheet.ListObjects(1), 1, xlFilterBelowAverage, xlFilterDynamic)
  ' �i�荞�݉���
  Call TableFilter(ActiveSheet.ListObjects(1), 1)
End Sub

'==========�@?(4)�@�v���V�[�W���ւ̎w���i���t��j�@============
Sub AutoFilter_exec_Date()
  'No.16 ���t��1�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">2022/8/20")
  'No.17 ���t��1�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & Format(CDate("2022/8/20"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal))
  'No.18 ���t��1�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & CLng(CDate("2022/8/20")))
  'No.19 �������t��OR�ōi����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, Array("2022/8/20", "2022/8/22"), xlFilterValues)
  'No.20 �������t��OR�ōi����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, , xlFilterValues, Array(2, "2022/8/20", 2, "2022/8/22"))
  'No.21 ���t��2�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">2022/8/20", xlAnd, "<2022/8/22")
  'No.22 ���t��2�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & Format(CDate("2022/8/20"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal), xlAnd, "<" & Format(CDate("2022/8/22"), ActiveSheet.ListObjects(1).ListColumns(3).DataBodyRange(1).NumberFormatLocal))
  'No.23 ���t��2�����̍i����
  Call TableFilter(ActiveSheet.ListObjects(1), 3, ">" & CLng(CDate("2022/8/20")), xlAnd, "<" & CLng(CDate("2022/8/22")))
  'No.24 ����̓��t�i�������i�荞�݁j
  Call TableFilter(ActiveSheet.ListObjects(1), 3, xlFilterToday, xlFilterDynamic)
  ' �i�荞�݉���
  Call TableFilter(ActiveSheet.ListObjects(1), 3)
End Sub

'==========�@?(1)�@�P�񂸂����@============
Sub TableFilterOff_01(T As ListObject)
  Dim i As Integer      '���J�E���^�ϐ��i��ʒu�j
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

'==========�@?(2)�@�i�荞�܂�Ă��Ȃ��Ă�1�񂸂����@============
Sub TableFilterOff_02(T As ListObject)
  Dim i As Integer      '���J�E���^�ϐ��i��ʒu�j
  For i = 1 To T.ListColumns.Count
    T.Range.AutoFilter Field:=i
  Next i
End Sub

'==========�@?(3)�@AutoFilter���\�b�h�Ń{�^�������E�ĕ\���@============
Sub TableFilterOff_03(T As ListObject)
  If Not T.AutoFilter Is Nothing Then
    T.Range.AutoFilter
  End If
  T.Range.AutoFilter
End Sub

'==========�@?(4)�@ShowAutoFilter�v���p�e�B�Ń{�^�������E�ĕ\���@============
Sub TableFilterOff_04(T As ListObject)
  If T.ShowAutoFilter = True Then
    T.ShowAutoFilter = False
  End If
  T.ShowAutoFilter = True
End Sub

'==========�@?(5)�@ShowAllData�őS�f�[�^��\���@============
Sub TableFilterOff_05(T As ListObject)
  If T.AutoFilter Is Nothing Then
    T.Range.AutoFilter
    Exit Sub
  End If
  If T.AutoFilter.FilterMode = True Then
    T.AutoFilter.ShowAllData
  End If
End Sub

'==========�@?(5)�@�e�[�u���ɍs��}�� �S�@============
Sub TableInsert_4(T As ListObject, arrayData As Variant)
  With T.ListRows.Add
    .Range = arrayData
  End With
End Sub

'==========�@?(7)�@�e�[�u���̍ŏI�s�̉��Ƀf�[�^��ǉ� �Q�@============
Sub TableInsert_6(T As ListObject, arrayData As Variant)
  T.HeaderRowRange.Offset(T.ListRows.Count + 1, 0) = arrayData
End Sub

'==========�@?(2)�@�S�s�𒲂׉��s�̏ꍇ�Ƀf�[�^���������@============
Sub TableUpdate_1(T As ListObject, Col As Variant, uniData As Variant)
  Dim i As Long      '���e�[�u���̃f�[�^�s��
  For i = 1 To T.ListRows.Count
    If T.DataBodyRange.Rows(i).Hidden = False Then
      T.ListColumns(Col).DataBodyRange(i) = uniData
    End If
  Next i
End Sub

'==========�@?(3)�@�i�荞�ݍs�݂̂��f�[�^�ύX�@============
Sub TableUpdate_2(T As ListObject, Col As Variant, uniData As Variant)
  Dim r As Range      '���e�[�u�����̉��s�~�w���̃Z���͈�
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  For Each r In T.ListColumns(Col).DataBodyRange.SpecialCells(xlCellTypeVisible)
    r = uniData
  Next r
End Sub

'==========�@?(4)�@�e�[�u���̗�S�̂ɑ΂��ăf�[�^�ύX�@============
Sub TableUpdate_3(T As ListObject, Col As Variant, uniData As Variant)
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  T.ListColumns(Col).DataBodyRange = uniData
End Sub

'==========�@?(2)�@���o�s�̃Z���͈͂��擾���A�i�荞�݉�����ɍ폜�@============
Sub TableDel_1(T As ListObject)
  Dim r As Range    '�����s�̃Z���͈́i�����s�ɓn�邱�Ƃ�����j
  On Error Resume Next
    Set r = T.DataBodyRange.SpecialCells(xlCellTypeVisible)
  On Error GoTo 0
  T.ShowAutoFilter = False
  T.ShowAutoFilter = True
  If Not r Is Nothing Then r.Delete
End Sub

'==========�@?(2)�@�i�荞�݌�̉��s�����̂܂܍s�폜�@============
Sub TableDel_2(T As ListObject)
  On Error Resume Next
    If T.DataBodyRange.SpecialCells(xlCellTypeVisible).Count = 0 Then
      Exit Sub
    End If
  On Error GoTo 0
  T.DataBodyRange.EntireRow.Delete
End Sub


