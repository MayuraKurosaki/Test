Attribute VB_Name = "DatePicker"
Option Explicit
Option Private Module

Private Type TState
    SelectionDate As Date
    IsShown As Boolean
End Type

Private This As TState

Public Property Get SelectionDate() As Date
    SelectionDate = This.SelectionDate
End Property

Public Property Let SelectionDate(ByVal RHS As Date)
    This.SelectionDate = RHS
End Property

Private Property Get IsShown() As Boolean
    IsShown = This.IsShown
End Property

Private Property Let IsShown(ByVal RHS As Boolean)
    This.IsShown = RHS
End Property

Public Sub Init(Optional ByVal Top As Single, Optional ByVal Left As Single, Optional ByVal InitialDate As Date)
    If IsShown Then Exit Sub
    If InitialDate = 0 Then InitialDate = VBA.Now
    SelectionDate = 0
    DatePickerForm.Top = 3
    DatePickerForm.Left = 6
    DatePickerForm.Show
End Sub
