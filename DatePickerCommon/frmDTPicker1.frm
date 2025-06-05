VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDTPicker1 
   Caption         =   "APIでDTPicker利用 （このフォームのサンプルは全て既定フォント）"
   ClientHeight    =   3735
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8440.001
   OleObjectBlob   =   "frmDTPicker1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDTPicker1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DTPCBox As clsDTPickerOnCombo3
Private DTP4 As clsDTPickerOnCombo3     '←単独利用

Private Sub UserForm_Initialize()
Dim j As Integer

    Set DTPCBox = New clsDTPickerOnCombo3
    With DTPCBox
        ' DTPickerが載るComboBoxを登録
        .Add ComboBox1
        .Add ComboBox2
        .Add ComboBox3
            ' : 以下､DTPicker を載せるコンボボックスを全て[Add]する

        .Create Me, "yyyy/MM/dd"       'DTPicker の生成
    End With

    DTPCBox.DateFormat(2) = "yyyy年M月/d日"         '[2]だけ編集を変えてみる
    DTPCBox.MinDate(2) = DateValue("2009/3/10")     '[2]だけ入力可能範囲を変えてみる
    DTPCBox.MaxDate(2) = DateValue("2013/7/20")

    Set DTP4 = New clsDTPickerOnCombo3      '単独で利用
    With DTP4
        .Add ComboBox4
        .Create Me, "yy年MM月dd日(dddd)", _
                BackColor:=&H99FFFF, _
                TitleBack:=&H808000, _
                Trailing:=&H99FFFF
    End With

    '配色情報の取得（ツール用）
    With DTPCBox        'Index省略→１番の情報を取得
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
    DTPCBox.Destroy         'DTPickerの破棄【必須】
    DTP4.Destroy
    Set DTPCBox = Nothing
    Set DTP4 = Nothing
End Sub



'==== ここから後ろの部分はコピーする必要はありません(テスト用コードです) =====

Private Sub cmdMsgBox_Click()
Dim strResult(1 To 4) As String

    '[DTP4]は単独利用なので添字/カッコは省略
    With DTPCBox
        strResult(1) = "( 1 ) Value= " & Format(.Value(1), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(1), "yyyy/mm/dd") & "～" & _
                                       Format(.MaxDate(1), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(1)
         
        strResult(2) = "( 2 ) Value= " & Format(.Value(2), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(2), "yyyy/mm/dd") & "～" & _
                                       Format(.MaxDate(2), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(2)
   
        strResult(3) = "( 3 ) Value= " & Format(.Value(3), "yyyy/mm/dd") & _
                       " , Range= [" & Format(.MinDate(3), "yyyy/mm/dd") & "～" & _
                                       Format(.MaxDate(3), "yyyy/mm/dd") & "]" & _
                       " , Enabled= " & .Enabled(3)
    End With
    
    strResult(4) = "( 4 ) Value= " & Format(DTP4.Value, "yyyy/mm/dd") & _
                   " , Range= [" & Format(DTP4.MinDate, "yyyy/mm/dd") & "～" & _
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
        .Enabled = Not .Enabled     '単独利用なので添字/カッコは省略
    End With
End Sub

Private Sub cmdShowForm2_Click()
    frmDTPicker2.Show
End Sub

'---- 配色変更 ----
Private Sub cmdChangeColor_Click()
    'Index を省略すると、グループ全体を一括で変更する
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
        '右クリック
        Call ColorSettingDialog(4)
    End If
End Sub

Private Sub lbl_txt4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(4)
    End If
End Sub

Private Sub TextBox5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(5)
    End If
End Sub

Private Sub lbl_txt5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(5)
    End If
End Sub

Private Sub TextBox6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(6)
    End If
End Sub

Private Sub lbl_txt6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(6)
    End If
End Sub

Private Sub TextBox7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(7)
    End If
End Sub

Private Sub lbl_txt7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(7)
    End If
End Sub

Private Sub TextBox8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(8)
    End If
End Sub

Private Sub lbl_txt8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 2) Then
        '右クリック
        Call ColorSettingDialog(8)
    End If
End Sub

Private Sub ColorSettingDialog(ByVal BoxNo As Integer)
        Load frmRGB
        frmRGB.SetRGB Me.Controls("lbl_txt" & BoxNo).BackColor
        frmRGB.Show
        
        If (frmRGB.lngColor < 0) Then
            'キャンセル
        Else
            Me.Controls("lbl_txt" & BoxNo).BackColor = frmRGB.lngColor
            Me.Controls("TextBox" & BoxNo).Value = frmRGB.strRGB
        End If
        
        Unload frmRGB
End Sub
