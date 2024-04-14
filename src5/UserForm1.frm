VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'ÉIÅ[ÉiÅ[ ÉtÉHÅ[ÉÄÇÃíÜâõ
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private searchCriteriaAge As Long
Private searchCriteriaName As String
Private searchCriteriaSex As String
Private searchCriteriaBloodType As String
Private searchCriteriaAddress As String
Private searchCriteriaDate As Date


Private Sub CheckBoxFemale_Change()

End Sub

Private Sub CheckBoxFemale_Click()

End Sub

Private Sub CheckBoxMale_Change()

End Sub

Private Sub CheckBoxMale_Click()

End Sub

'Private Sub ComboBoxAddress_AfterUpdate()
'    Debug.Print Me.ComboBoxAddress.Text
'End Sub

Private Sub ComboBoxAddress_Change()
    searchCriteriaAddress = Me.ComboBoxAddress.Text
    Debug.Print searchCriteriaAddress
End Sub

Private Sub ComboBoxAddress_DropButtonClick()
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_ìsìπï{åß").ListColumns("ìsìπï{åßñº").DataBodyRange
    Dim i As Long
    With ComboBoxAddress
        For i = 1 To listRange.Rows.Count
            .AddItem listRange(i)
        Next
    End With
End Sub

Private Sub CommandButtonDatePicker_Click()
    Call DatePicker.Init
    searchCriteriaDate = DatePicker.SelectionDate
    Me.TextBoxDate.Text = searchCriteriaDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
End Sub

Private Sub OptionButtonBllodTypeB_Change()

End Sub

Private Sub OptionButtonBllodTypeB_Click()

End Sub

Private Sub OptionButtonBloodTypeA_Change()

End Sub

Private Sub OptionButtonBloodTypeA_Click()

End Sub

Private Sub OptionButtonBloodTypeAB_Change()

End Sub

Private Sub OptionButtonBloodTypeAB_Click()

End Sub

Private Sub OptionButtonBloodTypeO_Change()

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
    searchCriteriaAge = TextBoxAge.Text
End Sub

Private Sub TextBoxAge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxAge_Change()

End Sub

Private Sub TextBoxDate_AfterUpdate()

End Sub

Private Sub TextBoxDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If VBA.IsDate(Me.TextBoxDate) Then
        Me.TextBoxDate.Text = Format(Me.TextBoxDate.Text, "YYYY/MM/DD")
    Else
        Me.TextBoxDate.SelStart = 0
        Me.TextBoxDate.SelLength = VBA.Len(Me.TextBoxDate.Text)
        Cancel = True
    End If
End Sub

Private Sub TextBoxDate_Change()

End Sub

Private Sub TextBoxName_AfterUpdate()
    searchCriteriaName = TextBoxName.Text
End Sub

Private Sub TextBoxName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxName_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub
