VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceCraftInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Field
    Name As String
    RefDate As Date
    RefPathNum As Long
    RefRevNum As Long
    TotalPath As Long
    M As Long
    N As Long
    L As Long
    FullPathList() As Long
    PathListPerDay() As Long
    SheetNo As Long
End Type

Private This As Field

Public Property Get Self() As SpaceCraftInfo
    Set Self = Me
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

'Public Property Let Name(ByVal RHS As String)
'    This.Name = RHS
'End Property
'
'Public Property Get Name() As String
'    Name = This.Name
'End Property
'Public Property Let Name(ByVal RHS As String)
'    This.Name = RHS
'End Property
'
'Public Property Get Name() As String
'    Name = This.Name
'End Property
'Public Property Let Name(ByVal RHS As String)
'    This.Name = RHS
'End Property
'
'Public Property Get Name() As String
'    Name = This.Name
'End Property






Private Sub Class_Initialize()

End Sub

