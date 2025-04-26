Attribute VB_Name = "Module1"
Option Explicit

Public Function IsHiddenRow(Target As Range) As Boolean
    If Target.EntireRow.Hidden Then
        IsHiddenRow = True
    Else
        IsHiddenRow = False
    End If
End Function
