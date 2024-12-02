VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implements IControlEvent

Private Type Field
    Controls As ControlEvents
    PrevControl As MSForms.IControl
'    Mode As FormMode
End Type

Private This As Field

Private Property Get IControlEvent_PrevControl() As ControlEvent
    Set IControlEvent_PrevControl = This.PrevControl
End Property

Private Property Let IControlEvent_PrevControl(RHS As ControlEvent)
    Set This.PrevControl = RHS
End Property

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAddControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.IControl)

End Sub

Private Sub IControlEvent_OnAfterUpdate(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, ByVal Cancel As MSForms.IReturnBoolean, ByVal Data As MSForms.IDataAutoWrapper, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.IReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, ByVal Cancel As MSForms.IReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.IDataAutoWrapper, ByVal x As Single, ByVal y As Single, ByVal Effect As MSForms.IReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, ByVal Cancel As MSForms.IReturnBoolean)

End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, ByVal Cancel As MSForms.IReturnBoolean)

End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnEnter(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnError(CtrlEvt As ControlEvent, ByVal Number As Integer, ByVal Description As MSForms.IReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.IReturnBoolean)

End Sub

Private Sub IControlEvent_OnExit(CtrlEvt As ControlEvent, ByVal Cancel As MSForms.IReturnBoolean)

End Sub

Private Sub IControlEvent_OnKeyDown(CtrlEvt As ControlEvent, ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub IControlEvent_OnKeyPress(CtrlEvt As ControlEvent, ByVal KeyAscii As MSForms.IReturnInteger)

End Sub

Private Sub IControlEvent_OnKeyUp(CtrlEvt As ControlEvent, ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub IControlEvent_OnLayout(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnListClick(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnMouseDown(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub IControlEvent_OnMouseUp(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub IControlEvent_OnRemoveControl(CtrlEvt As ControlEvent, ByVal Control As MSForms.IControl)

End Sub

Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.IReturnSingle, ByVal ActualDy As MSForms.IReturnSingle)

End Sub

Private Sub IControlEvent_OnSpinDown(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnSpinUp(CtrlEvt As ControlEvent)

End Sub

Private Sub IControlEvent_OnZoom(CtrlEvt As ControlEvent, Percent As Integer)

End Sub

Private Sub IControlEvent_OnMouseOver(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal x As Single, _
                                      ByVal y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseOver:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
'    Call MouseOver(CtrlEvt, Button, Shift, X, Y)
'    Set This.PrevControl = CtrlEvt
End Sub

Private Sub IControlEvent_OnMouseOut(CtrlEvt As ControlEvent, _
                                     ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal x As Single, _
                                     ByVal y As Single)
'    Debug.Print CtrlEvt.Control.Name & " MouseOut:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
'    Call MouseOut(CtrlEvt, Button, Shift, X, Y)
End Sub

