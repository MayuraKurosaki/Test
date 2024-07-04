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
    ParentForm As MSForms.UserForm
    List As Dictionary
End Type

Private This As Field

Public Property Get ParentForm() As MSForms.UserForm
    Set ParentForm = This.ParentForm
End Property

Public Property Let ParentForm(RHS As MSForms.UserForm)
    Set This.ParentForm = RHS
End Property

Public Property Get Item(ControlName As String) As ControlEvent
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Set Item = This.List(ControlName)
End Property

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

Public Sub Init()
    Dim ctrl As MSForms.Control

    If This.ParentForm Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
'    Set This.List = New Dictionary
'    Set this.List = CreateObject("Scripting.Dictionary")
    For Each ctrl In ParentForm.Controls
        With New ControlEvent
            .ParentForm = ParentForm
            .Control = ctrl
            This.List.Add ctrl.Name, .Self
        End With
    Next ctrl
End Sub

Public Sub RegisterControl(ctrl As MSForms.Control, Optional ByVal Attr As String = "")
    If This.ParentForm Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
    With New ControlEvent
        .ParentForm = ParentForm
        .Control = ctrl
        .AttributeItems = Attr
'        .BaseStyle = BaseStyle
'        .HighlightStyle = HighlightStyle
'        .ClickStyle = ClickStyle
        If This.List.Exists(ctrl.Name) Then
            Set This.List.Item(ctrl.Name) = .Self
        Else
            This.List.Add ctrl.Name, .Self
        End If
    End With
End Sub

Private Sub Class_Initialize()
    Set This.List = New Dictionary
End Sub

Private Sub Class_Terminate()
    Dim ctrlKey As Variant
    If This.ParentForm Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録解除
    If Not This.List Is Nothing Then
        For Each ctrlKey In This.List.Keys
            This.List(ctrlKey).ParentForm = Nothing
        Next ctrlKey
    
        'リストのオブジェクトの参照を廃棄
        This.List.RemoveAll
        Set This.List = Nothing
    End If
End Sub
