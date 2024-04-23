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

Private Type TState
    Parent As Object
    List As Dictionary
'    List As Object
End Type

Private this As TState

Public Property Get Parent() As Object
    Set Parent = this.Parent
End Property

Public Property Let Parent(RHS As Object)
    Set this.Parent = RHS
End Property

Public Sub Init()
    Dim ctrl As MSForms.control

    If Parent Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
    Set this.List = New Dictionary
'    Set this.List = CreateObject("Scripting.Dictionary")
    For Each ctrl In Parent.Controls
        With New ControlEvent
            .Parent = Parent
            .Item = ctrl
            this.List.Add ctrl.Name, .Self
        End With
    Next ctrl
End Sub

Private Sub Class_Terminate()
    Dim ctrlKey As Variant
    If Parent Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録解除
    If Not this.List Is Nothing Then
        For Each ctrlKey In this.List.Keys
            this.List(ctrlKey).Parent = Nothing
        Next ctrlKey
    
        'リストのオブジェクトの参照を廃棄
        this.List.RemoveAll
        Set this.List = Nothing
    End If
End Sub
