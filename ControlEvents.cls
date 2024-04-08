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
End Type

Private This As TState

Public Property Get Parent() As Object
    Set Parent = This.Parent
End Property

Public Property Let Parent(RHS As Object)
    Set This.Parent = RHS
End Property

Public Sub init()
    Dim Ctrl As MSForms.Control

    If Parent Is Nothing Then
        Exit Sub
    End If

    '�R���g���[���̖��O�ƃN���X�I�u�W�F�N�g��o�^
    Set This.List = New Dictionary
    For Each Ctrl In Parent.Controls
        With New ControlEvent
            .Parent = Parent
            .Item = Ctrl
            This.List.Add Ctrl.name, .Self
        End With
    Next Ctrl
End Sub

Private Sub Class_Terminate()
    '���X�g�̃I�u�W�F�N�g�̎Q�Ƃ�p��
    If Not This.List Is Nothing Then
        This.List.RemoveAll
        Set This.List = Nothing
    End If
End Sub