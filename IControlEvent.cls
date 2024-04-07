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

Public Sub OnChange(Cont As MSForms.Control)
End Sub

Public Sub OnListClick(Cont As MSForms.Control)
End Sub

Public Sub OnClick(Cont As MSForms.Control)
End Sub

Public Sub OnDropButtonClick(Cont As MSForms.Control)
End Sub

Public Sub OnDblClick(Cont As MSForms.Control, _
                      ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnKeyDown(Cont As MSForms.Control, _
                     ByVal KeyCode As MSForms.ReturnInteger, _
                     ByVal Shift As Integer)
End Sub

Public Sub OnKeyUp(Cont As MSForms.Control, _
                   ByVal KeyCode As MSForms.ReturnInteger, _
                   ByVal Shift As Integer)
End Sub

Public Sub OnMouseDown(Cont As MSForms.Control, _
                       ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal x As Single, _
                       ByVal y As Single)
End Sub

Public Sub OnMouseMove(Cont As MSForms.Control, _
                       ByVal Button As Integer, _
                       ByVal Shift As Integer, _
                       ByVal x As Single, _
                       ByVal y As Single)
End Sub

Public Sub OnMouseUp(Cont As MSForms.Control, _
                     ByVal Button As Integer, _
                     ByVal Shift As Integer, _
                     ByVal x As Single, _
                     ByVal y As Single)
End Sub

Public Sub OnKeyPress(Cont As MSForms.Control, _
                      ByVal KeyAscii As MSForms.ReturnInteger)
End Sub

Public Sub OnExit(Cont As MSForms.Control, _
                  ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnAfterUpdate(Cont As MSForms.Control)
End Sub

Public Sub OnBeforeUpdate(Cont As MSForms.Control, _
                          ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Public Sub OnEnter(Cont As MSForms.Control)
End Sub

Public Sub BeforeDragOver(Cont As MSForms.Control, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
End Sub

Public Sub OnBeforeDropOrPaste(Cont As MSForms.Control, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal x As Single, _
                               ByVal y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
End Sub

Public Sub OnError(Cont As MSForms.Control, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
End Sub

'Public Sub OnGotFocus(Cont As MSForms.Control)
'End Sub
'
'Public Sub OnLostFocus(Cont As MSForms.Control)
'End Sub

Public Sub AddControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
End Sub

Public Sub Layout(Cont As MSForms.Control)
End Sub

Public Sub RemoveControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
End Sub

Public Sub Scroll(Cont As MSForms.Control, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
End Sub

'' ScrollBar
'Public Sub Scroll(Cont As MSForms.Control)
'End Sub

Public Sub Zoom(Cont As MSForms.Control, Percent As Integer)
End Sub

Public Sub OnSpinDown(Cont As MSForms.Control)
End Sub

Public Sub OnSpinUp(Cont As MSForms.Control)
End Sub

