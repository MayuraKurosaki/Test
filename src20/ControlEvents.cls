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
    Dim Ctrl As MSForms.Control

    If This.ParentForm Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
    For Each Ctrl In ParentForm.Controls
        With New ControlEvent
            .ParentForm = ParentForm
            .Control = Ctrl
            This.List.Add Ctrl.Name, .Self
        End With
    Next Ctrl
End Sub

Public Function AddControl(bstrProgID As String, Optional Name As String = "", Optional Visible As Boolean = True) As MSForms.Control
    Set AddControl = This.ParentForm.Controls.Add(bstrProgID, Name, Visible)
End Function

Public Sub RegisterControl(Ctrl As MSForms.Control, Optional ByVal AttributeItems As String = "")
    If This.ParentForm Is Nothing Then Exit Sub

    'コントロールの名前とクラスオブジェクトを登録
    With New ControlEvent
        .ParentForm = ParentForm
        .Control = Ctrl
        .AttributeItems = AttributeItems
        If This.List.Exists(Ctrl.Name) Then
            Set This.List.Item(Ctrl.Name) = .Self
        Else
            This.List.Add Ctrl.Name, .Self
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
