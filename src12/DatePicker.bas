Attribute VB_Name = "DatePicker"
Option Explicit
Option Private Module

Private Type TState
    SelectionDate As Date
    IsShown As Boolean
End Type

Private this As TState

Public Property Get SelectionDate() As Date
    SelectionDate = this.SelectionDate
End Property

Public Property Let SelectionDate(ByVal RHS As Date)
    this.SelectionDate = RHS
End Property

Private Property Get IsShown() As Boolean
    IsShown = this.IsShown
End Property

Private Property Let IsShown(ByVal RHS As Boolean)
    this.IsShown = RHS
End Property

Public Sub Init(Optional ByVal Top As Single, Optional ByVal Left As Single, Optional ByVal InitialDate As Date)
    If IsShown Then Exit Sub
    If InitialDate = 0 Then InitialDate = VBA.Now
    SelectionDate = 0
    DatePickerForm.Top = Top
    DatePickerForm.Left = Left
    DatePickerForm.Show
End Sub
