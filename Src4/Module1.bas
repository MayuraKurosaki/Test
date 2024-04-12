Attribute VB_Name = "Module1"
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

Sub search()
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
