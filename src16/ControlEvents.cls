VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Field
    Parent As Object
    List As Dictionary
'    List As Object
End Type

Private This As Field

Public Property Get Parent() As Object
    Set Parent = This.Parent
End Property

Public Property Let Parent(RHS As Object)
    Set This.Parent = RHS
End Property

Public Property Get Item(Name As String) As Object
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Set Item = This.List(Name)
End Property

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

Public Sub Init()
    Dim ctrl As MSForms.Control

    If Parent Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
    Set This.List = New Dictionary
'    Set this.List = CreateObject("Scripting.Dictionary")
    For Each ctrl In Parent.Controls
        With New ControlEvent
            .Parent = Parent
            .Control = ctrl
            This.List.Add ctrl.Name, .Self
        End With
    Next ctrl
End Sub

Private Sub Class_Terminate()
    Dim ctrlKey As Variant
    If Parent Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録解除
    If Not This.List Is Nothing Then
        For Each ctrlKey In This.List.Keys
            This.List(ctrlKey).Parent = Nothing
        Next ctrlKey
    
        'リストのオブジェクトの参照を廃棄
        This.List.RemoveAll
        Set This.List = Nothing
    End If
End Sub
