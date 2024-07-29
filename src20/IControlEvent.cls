VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub OnChange(CtrlEvt As ControlEvent)
End Sub

Public Sub OnListClick(CtrlEvt As ControlEvent)
End Sub

Public Sub OnClick(CtrlEvt As ControlEvent)
End Sub

Public Sub OnDropButtonClick(CtrlEvt As ControlEvent)
End Sub

Public Sub OnDblClick(CtrlEvt As ControlEvent, _
                      ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnKeyDown(CtrlEvt As ControlEvent, _
                     ByVal KeyCode As MSForms.ReturnInteger, _
                     ByVal Shift As Integer)
End Sub

Public Sub OnKeyUp(CtrlEvt As ControlEvent, _
                   ByVal KeyCode As MSForms.ReturnInteger, _
                   ByVal Shift As Integer)
End Sub

Public Sub OnMouseDown(CtrlEvt As ControlEvent, _
                       ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal X As Single, _
                       ByVal Y As Single)
End Sub

Public Sub OnMouseMove(CtrlEvt As ControlEvent, _
                       ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal X As Single, _
                       ByVal Y As Single)
End Sub

Public Sub OnMouseUp(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                     ByVal X As Single, _
                     ByVal Y As Single)
End Sub

Public Sub OnKeyPress(CtrlEvt As ControlEvent, _
                      ByVal KeyAscii As MSForms.ReturnInteger)
End Sub

Public Sub OnExit(CtrlEvt As ControlEvent, _
                  ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnAfterUpdate(CtrlEvt As ControlEvent)
End Sub

Public Sub OnBeforeUpdate(CtrlEvt As ControlEvent, _
                          ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnEnter(CtrlEvt As ControlEvent)
End Sub

Public Sub OnBeforeDragOver(CtrlEvt As ControlEvent, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
End Sub

Public Sub OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
End Sub

Public Sub OnError(CtrlEvt As ControlEvent, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
End Sub

'Public Sub OnGotFocus(CtrlEvt As ControlEvent)
'End Sub
'
'Public Sub OnLostFocus(CtrlEvt As ControlEvent)
'End Sub

Public Sub OnAddControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
End Sub

Public Sub OnLayout(CtrlEvt As ControlEvent)
End Sub

Public Sub OnRemoveControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.Control)
End Sub

Public Sub OnScroll(CtrlEvt As ControlEvent, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
End Sub

'' ScrollBar
'Public Sub OnScroll(CtrlEvt As ControlEvent)
'End Sub

Public Sub OnZoom(CtrlEvt As ControlEvent, Percent As Integer)
End Sub

Public Sub OnSpinDown(CtrlEvt As ControlEvent)
End Sub

Public Sub OnSpinUp(CtrlEvt As ControlEvent)
End Sub

Public Sub OnMouseOver(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                       ByVal X As Single, _
                       ByVal Y As Single)
End Sub

Public Sub OnMouseOut(CtrlEvt As ControlEvent, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                       ByVal X As Single, _
                       ByVal Y As Single)
End Sub

Public Property Get PrevControl() As ControlEvent
End Property

Public Property Let PrevControl(RHS As ControlEvent)
End Property

Public Property Get ControlEvents() As ControlEvents
End Property

Public Property Get Base() As MSForms.UserForm
End Property

