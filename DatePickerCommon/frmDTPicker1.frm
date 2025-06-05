VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDTPicker1 
   Caption         =   "API��DTPicker���p �i���̃t�H�[���̃T���v���͑S�Ċ���t�H���g�j"
   ClientHeight    =   3735
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8440.001
   OleObjectBlob   =   "frmDTPicker1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmDTPicker1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DTPCBox As clsDTPickerOnCombo3
Private DTP4 As clsDTPickerOnCombo3     '���P�Ɨ��p

Private Sub UserForm_Initialize()
Dim j As Integer

    Set DTPCBox = New clsDTPickerOnCombo3
    With DTPCBox
        ' DTPicker���ڂ�ComboBox��o�^
        .Add ComboBox1
        .Add ComboBox2
        .Add ComboBox3
            ' : �ȉ��DTPicker ���ڂ���R���{�{�b�N�X��S��[Add]����

        .Create Me, "yyyy/MM/dd"       'DTPicker �̐���
    End With

    DTPCBox.DateFormat(2) = "yyyy�NM��/d��"         '[2]�����ҏW��ς��Ă݂�
    DTPCBox.MinDate(2) = DateValue("2009/3/10")     '[2]�������͉\�͈͂�ς��Ă݂�
    DTPCBox.MaxDate(2) = DateValue("2013/7/20")

    Set DTP4 = New clsDTPickerOnCombo3      '�P�Ƃŗ��p
    With DTP4
        .Add ComboBox4
        .Create Me, "yy�NMM��dd��(dddd)", _
                BackColor:=&H99FFFF, _
                TitleBack:=&H808000, _
                Trailing:=&H99FFFF
    End With

    '�z�F���̎擾�i�c�[���p�j
    With DTPCBox        'Index�ȗ����P�Ԃ̏����擾
        lbl_txt4.BackColor = .CalendarForeColor
        lbl_txt5.BackColor = .CalendarBackColor
        lbl_txt6.BackColor = .CalendarTitleForeColor
        lbl_txt7.BackColor = .CalendarTitleBackColor
        lbl_txt8.BackColor = .CalendarTrailingForeColor
    End With
    For j = 4 To 8
        Me.Controls("TextBox" & j).Value = _
            Right("000000" & Hex(Me.Controls("lbl_txt" & j).BackColor), 6)
    Next j
End Sub

Private Sub UserForm_Terminate()
    DTPCBox.Destroy         'DTPicker�̔j���y�K�{�z
    DTP4.Destroy
    Set DTPCBox = Nothing
    Set DTP4 = Nothing
End Sub



'==== ����������̕����̓R�s�[����K�v�͂���܂���(�e�X�g�p�R�[�h�ł�) =====

Private Sub cmdMsgBox_Click()
Dim strResult(1 To 4) As String

    '[DTP4]�͒P�Ɨ��p�Ȃ̂œY��/�J�b�R�͏ȗ�
    With DTPCBox
        strResult(1) = "( 1 ) Value= " & Format(.Value(1), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(1), "yyyy/mm/dd") & "�`" & _
                                       Format(.MaxDate(1), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(1)
         
        strResult(2) = "( 2 ) Value= " & Format(.Value(2), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(2), "yyyy/mm/dd") & "�`" & _
                                       Format(.MaxDate(2), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(2)
   
        strResult(3) = "( 3 ) Value= " & Format(.Value(3), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(3), "yyyy/mm/dd") & "�`" & _
                                       Format(.MaxDate(3), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(3)
    End With
    
    strResult(4) = "( 4 ) Value= " & Format(DTP4.Value, "yyyy/mm/dd") & _
                   " , Range= [" & Format(DTP4.MinDate, "yyyy/mm/dd") & "�`" & _
                                   Format(DTP4.MaxDate, "yyyy/mm/dd") & "]" & _
                   " , Enabled= " & DTP4.Enabled

    MsgBox strResult(1) & Space(5) & vbCrLf & vbCrLf & _
           strResult(2) & Space(5) & vbCrLf & vbCrLf & _
           strResult(3) & Space(5) & vbCrLf & vbCrLf & _
           strResult(4) & Space(5)
End Sub

Private Sub cmdEnabled1_Click()
    With DTPCBox
        .Enabled(1) = Not .Enabled(1)
    End With
End Sub

Private Sub cmdEnabled2_Click()
    With DTPCBox
        .Enabled(2) = Not .Enabled(2)
    End With
End Sub

Private Sub cmdEnabled3_Click()
    With DTPCBox
        .Enabled(3) = Not .Enabled(3)
    End With
End Sub

Private Sub cmdEnabled4_Click()
    With DTP4
        .Enabled = Not .Enabled     '�P�Ɨ��p�Ȃ̂œY��/�J�b�R�͏ȗ�
    End With
End Sub

Private Sub cmdShowForm2_Click()
    frmDTPicker2.Show
End Sub

'---- �z�F�ύX ----
Private Sub cmdChangeColor_Click()
    'Index ���ȗ�����ƁA�O���[�v�S�̂��ꊇ�ŕύX����
    With DTPCBox
        .CalendarForeColor = lbl_txt4.BackColor
        .CalendarBackColor = lbl_txt5.BackColor
        .CalendarTitleForeColor = lbl_txt6.BackColor
        .CalendarTitleBackColor = lbl_txt7.BackColor
        .CalendarTrailingForeColor = lbl_txt8.BackColor
    End With
End Sub

Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(4)
    End If
End Sub

Private Sub lbl_txt4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(4)
    End If
End Sub

Private Sub TextBox5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(5)
    End If
End Sub

Private Sub lbl_txt5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(5)
    End If
End Sub

Private Sub TextBox6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(6)
    End If
End Sub

Private Sub lbl_txt6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(6)
    End If
End Sub

Private Sub TextBox7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(7)
    End If
End Sub

Private Sub lbl_txt7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(7)
    End If
End Sub

Private Sub TextBox8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(8)
    End If
End Sub

Private Sub lbl_txt8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '�E�N���b�N
        Call ColorSettingDialog(8)
    End If
End Sub

Private Sub ColorSettingDialog(ByVal BoxNo As Integer)
        Load frmRGB
        frmRGB.SetRGB Me.Controls("lbl_txt" & BoxNo).BackColor
        frmRGB.Show
        
        If (frmRGB.lngColor < 0) Then
            '�L�����Z��
        Else
            Me.Controls("lbl_txt" & BoxNo).BackColor = frmRGB.lngColor
            Me.Controls("TextBox" & BoxNo).Value = frmRGB.strRGB
        End If
        
        Unload frmRGB
End Sub
