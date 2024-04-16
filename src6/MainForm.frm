VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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

'���F��:approver�@���F����:Approve�@���F:Approval
'����:signature
'����:constraint
'OperationProcedure
'Reason for operation
'Operation results
'TimeUnit
'�F��:authentication

Private Sub CheckBoxFemale_Change()
    If Me.CheckBoxFemale.Value Then searchCriteriaSex = "��"
    Call Filter
End Sub

Private Sub CheckBoxFemale_Click()

End Sub

Private Sub CheckBoxMale_Change()
    If Me.CheckBoxMale.Value Then searchCriteriaSex = "�j"
    Call Filter
End Sub

Private Sub CheckBoxMale_Click()

End Sub

'Private Sub ComboBoxAddress_AfterUpdate()
'    Debug.Print Me.ComboBoxAddress.Text
'End Sub

Private Sub ComboBoxAddress_Change()
    searchCriteriaAddress = Me.ComboBoxAddress.Text
    Call Filter
End Sub

Private Sub ComboBoxAddress_DropButtonClick()
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_�s���{��").ListColumns("�s���{����").DataBodyRange
    Dim i As Long
    With ComboBoxAddress
        For i = 1 To listRange.Rows.Count
            .AddItem listRange(i)
        Next
    End With
End Sub

Private Sub CommandButtonDatePicker_Click()
    Call DatePicker.Init
    Me.TextBoxDate.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
End Sub

Private Sub ListBoxResultList_Click()

End Sub

Private Sub OptionButtonBloodTypeB_Change()
    If Me.OptionButtonBloodTypeB.Value Then searchCriteriaBloodType = "B"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeB_Click()

End Sub

Private Sub OptionButtonBloodTypeA_Change()
    If Me.OptionButtonBloodTypeA.Value Then searchCriteriaBloodType = "A"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeA_Click()

End Sub

Private Sub OptionButtonBloodTypeAB_Change()
    If Me.OptionButtonBloodTypeAB.Value Then searchCriteriaBloodType = "AB"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeAB_Click()

End Sub

Private Sub OptionButtonBloodTypeO_Change()
    If Me.OptionButtonBloodTypeO.Value Then searchCriteriaBloodType = "O"
    Call Filter
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
    Call Filter
End Sub

Private Sub TextBoxAge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxAge_Change()

End Sub

Private Sub TextBoxDate_AfterUpdate()
    Call Filter
End Sub

Private Sub TextBoxDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If VBA.IsDate(Me.TextBoxDate) Then
        searchCriteriaDate = Me.TextBoxDate.Value
        Me.TextBoxDate.Text = Format(searchCriteriaDate, "YYYY�NMM��DD��")
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
    If KeyCode = 187 And Shift = 2 Then TextBoxDate.Value = Format(Now, "YYYY�NMM��DD��") ' Ctrl + �u;�v
End Sub

Private Sub TextBoxName_AfterUpdate()
    searchCriteriaName = TextBoxName.Text
    Call Filter
End Sub

Private Sub TextBoxName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxName_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    searchCriteriaAge = -1
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub Filter()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Dummy").ListObjects("T_Dummy")
        If Me.TextBoxName.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("����").index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�N��").index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxAddress.Value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�Z��").index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.CheckBoxFemale.Value Or Me.CheckBoxMale Then _
            .Range.AutoFilter Field:=.ListColumns("����").index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("���t�^").index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxDate.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("���N����").index, Criteria1:=Format(searchCriteriaDate, "YYYY�NMM��DD��"), VisibleDropDown:=False
    
        Dim CellsCnt As Long    '���i�荞���ް��ٌ̾�
        Dim ColCnt As Long      '��ð��ق̗�
        Dim buf1 As Variant    '���e�[�u���S�̂̃f�[�^
        buf1 = .Range
        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).Count
        ColCnt = UBound(buf1, 2)
        
        Dim buf2 As Variant    '���߂�l�Ƃ���ꎞ�I�Ȕz��
        ReDim buf2(1 To (CellsCnt / ColCnt), 1 To ColCnt)
        
        Dim i As Long            '�������ϐ��i�z��̍s�ʒu�j
        Dim j As Long            '�������ϐ��i�z��̗�ʒu�j
        Dim k As Long            '�e�[�u���̃f�[�^�s�{�^�C�g���s�̍s��
        For k = 1 To UBound(buf1, 1)
          If .Range.Rows(k).Hidden = False Then
            i = i + 1
            For j = 1 To ColCnt
              buf2(i, j) = buf1(k, j)
            Next j
          End If
        Next k
        
        '�I�[�g�t�B���^������
        .Range.AutoFilter
        .ShowAutoFilter = False
    End With
    With Me.ListBoxResultList
        .Clear
        .List = buf2
    End With
    Application.ScreenUpdating = True

End Sub
