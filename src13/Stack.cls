VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Items() As Variant

Property Get Count() As Integer
    Count = UBound(Items)
End Property

Property Get Top() As Variant
    Top = Items(UBound(Items))
End Property

Public Function Pop() As Variant
    If UBound(Items) > 0 Then
        Pop = Items(UBound(Items))
        ReDim Preserve Items(UBound(Items) - 1)
    Else
        Pop = Empty
    End If
End Function

Public Sub Push(ByRef x As Variant)
    ReDim Preserve Items(UBound(Items) + 1)
    Items(UBound(Items)) = x
End Sub

Private Sub Class_Initialize()
    ReDim Items(0)
End Sub

