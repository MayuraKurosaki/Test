VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Field
    Parent As ISink
End Type

Private This As Field

Public Property Let Parent(RHS As ISink)
    Set This.Parent = RHS
End Property

Public Property Get Parent() As ISink
    'インターフェイスで受けます。
    Set Parent = This.Parent
End Property

'--------------------コールバック関数群
Public Sub OnChange()
Attribute OnChange.VB_UserMemId = 2
'Attribute OnChange.VB_UserMemId = 2
    Call Parent.OnChange
End Sub

Public Sub OnListClick()
Attribute OnListClick.VB_UserMemId = -610
'Attribute OnListClick.VB_UserMemId = -610
    Call Parent.OnListClick
End Sub

Public Sub OnClick()
Attribute OnClick.VB_UserMemId = -600
'Attribute OnClick.VB_UserMemId = -600
    Call Parent.OnClick
End Sub

Public Sub OnDropButtonClick()
Attribute OnDropButtonClick.VB_UserMemId = 2002
'Attribute OnDropButtonClick.VB_UserMemId = 2002
    Call Parent.OnDropButtonClick
End Sub

Public Sub OnDblClick(ByVal Cancel As MSForms.ReturnBoolean)
Attribute OnDblClick.VB_UserMemId = -601
'Attribute OnDblClick.VB_UserMemId = -601
    Call Parent.OnDblClick(Cancel)
End Sub

Public Sub OnKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                     ByVal Shift As Integer)
Attribute OnKeyDown.VB_UserMemId = -602
'Attribute OnKeyDown.VB_UserMemId = -602
    Debug.Print TypeName(Parent) & "KeyDown(" & KeyCode & ")"
    Call Parent.OnKeyDown(KeyCode, Shift)
End Sub

Public Sub OnKeyUp(ByVal KeyCode As MSForms.ReturnInteger, _
                   ByVal Shift As Integer)
Attribute OnKeyUp.VB_UserMemId = -604
'Attribute OnKeyUp.VB_UserMemId = -604
    Call Parent.OnKeyUp(KeyCode, Shift)
End Sub

Public Sub OnMouseDown(ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal x As Single, _
                       ByVal y As Single)
Attribute OnMouseDown.VB_UserMemId = -605
'Attribute OnMouseDown.VB_UserMemId = -605
    Call Parent.OnMouseDown(Button, Shift, x, y)
End Sub

Public Sub OnMouseMove(ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal x As Single, _
                       ByVal y As Single)
Attribute OnMouseMove.VB_UserMemId = -606
'Attribute OnMouseMove.VB_UserMemId = -606
    Call Parent.OnMouseMove(Button, Shift, x, y)
End Sub

Public Sub OnMouseUp(ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                     ByVal x As Single, _
                     ByVal y As Single)
Attribute OnMouseUp.VB_UserMemId = -607
'Attribute OnMouseUp.VB_UserMemId = -607
    Call Parent.OnMouseUp(Button, Shift, x, y)
End Sub

Public Sub OnKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Attribute OnKeyPress.VB_UserMemId = -603
'Attribute OnKeyPress.VB_UserMemId = -603
    Call Parent.OnKeyPress(KeyAscii)
End Sub

Public Sub OnExit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute OnExit.VB_UserMemId = -2147384829
'Attribute OnExit.VB_UserMemId = -2147384829
    Call Parent.OnExit(Cancel)
End Sub

Public Sub OnAfterUpdate()
Attribute OnAfterUpdate.VB_UserMemId = -2147384832
'Attribute OnAfterUpdate.VB_UserMemId = -2147384832
    Call Parent.OnAfterUpdate
End Sub

Public Sub OnBeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Attribute OnBeforeUpdate.VB_UserMemId = -2147384831
'Attribute OnBeforeUpdate.VB_UserMemId = -2147384831
    Call Parent.OnBeforeUpdate(Cancel)
End Sub

Public Sub OnEnter()
Attribute OnEnter.VB_UserMemId = -2147384830
'Attribute OnEnter.VB_UserMemId = -2147384830
    Call Parent.OnEnter
End Sub

'Public Sub OnBeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean,
Public Sub OnBeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
                                                        ByVal Data As MSForms.DataObject, _
                                                        ByVal x As Single, _
                                                        ByVal y As Single, _
                                                        ByVal DragState As MSForms.fmDragState, _
                                                        ByVal Effect As MSForms.ReturnEffect, _
                                                        ByVal Shift As Integer)
    Call Parent.OnBeforeDragOver(Cancel, Data, x, y, DragState, Effect, Shift)
End Sub

Public Sub OnBeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
                                                           ByVal Action As MSForms.fmAction, _
                                                           ByVal Data As MSForms.DataObject, _
                                                           ByVal x As Single, _
                                                           ByVal y As Single, _
                                                           ByVal Effect As MSForms.ReturnEffect, _
                                                           ByVal Shift As Integer)
Attribute OnBeforeDropOrPaste.VB_UserMemId = 4
'Attribute OnBeforeDropOrPaste.VB_UserMemId = 4
    Call Parent.OnBeforeDropOrPaste(Cancel, Action, Data, x, y, Effect, Shift)
End Sub

Public Sub OnError(ByVal Number As Integer, _
                                   ByVal Description As MSForms.ReturnString, _
                                   ByVal SCode As Long, ByVal Source As String, _
                                   ByVal HelpFile As String, _
                                   ByVal HelpContext As Long, _
                                   ByVal CancelDisplay As MSForms.ReturnBoolean)
Attribute OnError.VB_UserMemId = -608
'Attribute OnError.VB_UserMemId = -608
    Call Parent.OnError(Number, Description, SCode, HelpFile, HelpContext, CancelDisplay)
End Sub

'Public Sub OnGotFocus()
'Attribute OnGotFocus.VB_UserMemId = 1541
''Attribute OnGotFocus.VB_UserMemId = 1541
'    Call Parent.OnGotFocus
'End Sub
'
'Public Sub OnLostFocus()
'Attribute OnLostFocus.VB_UserMemId = 1542
''Attribute OnLostFocus.VB_UserMemId = 1542
'    Call Parent.OnLostFocus
'End Sub

Public Sub OnAddControl(ByVal Control As MSForms.Control)
'Public Sub OnAddControl(ByVal Control As MSForms.Control)
    Call Parent.OnAddControl(Control)
End Sub

Public Sub OnLayout()
'Public Sub OnLayout()
    Call Parent.OnLayout
End Sub

Public Sub OnRemoveControl(ByVal Control As MSForms.Control)
'Public Sub OnRemoveControl(ByVal Control As MSForms.Control)
    Call Parent.OnRemoveControl(Control)
End Sub

'Public Sub OnScroll(ByVal ActionX As MSForms.fmScrollAction,
Public Sub OnScroll(ByVal ActionX As MSForms.fmScrollAction, _
                                        ByVal ActionY As MSForms.fmScrollAction, _
                                        ByVal RequestDx As Single, _
                                        ByVal RequestDy As Single, _
                                        ByVal ActualDx As MSForms.ReturnSingle, _
                                        ByVal ActualDy As MSForms.ReturnSingle)
    Debug.Print Parent.Control.Name & " Scroll:"
    Call Parent.OnScroll(ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

'' ScrollBar
'Public Sub OnScroll()
''Attribute OnScroll.VB_Description = "Scrollイベントハンドラ"
'    Call Parent.OnScroll
'End Sub

Public Sub OnZoom(Percent As Integer)
'Public Sub OnZoom(Percent As Integer)
    Call Parent.OnZoom(Percent)
End Sub

Public Sub OnSpinDown()
Attribute OnSpinDown.VB_UserMemId = 6
'Attribute OnSpinDown.VB_UserMemId = 6
'0x00000006
    Call Parent.OnSpinDown
End Sub

Public Sub OnSpinUp()
Attribute OnSpinUp.VB_UserMemId = 5
'Attribute OnSpinUp.VB_UserMemId = 5
'0x00000005
    Call Parent.OnSpinUp
End Sub

