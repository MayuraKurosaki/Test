VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDTPicker2 
   Caption         =   "API �� DTPicker ���p �i�t�H���g�ύX�T���v���j"
   ClientHeight    =   3525
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6630
   OleObjectBlob   =   "frmDTPicker2.frx":0000
End
Attribute VB_Name = "frmDTPicker2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DTPCBox1 As clsDTPickerOnCombo3
Private DTPCBox2 As clsDTPickerOnCombo3

Private Sub UserForm_Initialize()
    Me.Top = frmDTPicker1.Top + 50
    Me.Left = frmDTPicker1.Left + 50

    Set DTPCBox1 = New clsDTPickerOnCombo3
    With DTPCBox1
        .Add ComboBox1
        .Add ComboBox2
        .Add ComboBox3
        .Add ComboBox4
        .Create Me, "yyyy�NMMMMd��(ddd)", DefaultFONT:=False
        .Value(1) = DateValue("2004/10/10")
        .Value(2) = DateValue("2004/2/29")
    End With

    Set DTPCBox2 = New clsDTPickerOnCombo3
    With DTPCBox2
        .Add ComboBox5
        .Add ComboBox6
        .Create Me, "yyyy/MM/dd(dddd)", DefaultFONT:=2  '2:ComboBox���̂݃t�H���g�ύX
        
        .CalendarBackColor(0) = &H99FFFF           '(1)(2)�ꏏ�ɐݒ�
        .CalendarTitleBackColor(0) = &H808000      '    �V
        .CalendarTrailingForeColor(0) = &H99FFFF   '    �V
        
        .Value(1) = DateValue("2004/7/7")
        .Value(2) = DateValue("2004/5/5")
    End With
End Sub

Private Sub UserForm_Terminate()
    DTPCBox1.Destroy
    DTPCBox2.Destroy
    Set DTPCBox1 = Nothing
    Set DTPCBox2 = Nothing
End Sub

