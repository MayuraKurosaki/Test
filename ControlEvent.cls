VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'API定義 [ ConnectToConnectionPoint ]
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
 
Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" _
         (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, _
         ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, _
         ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long

Implements ISink

Private Type TState
    Caller As IControlEvent
    EventSink As Sink
    Control As Object
    Cookie As Long
End Type

Private This As TState

Public Property Get Item() As Object
    Set Item = This.Control
End Property
 
Public Property Let Item(RHS As Object)
    Set This.Control = RHS
    Call ConnectEvent(True)
End Property
 
Public Property Get Parent() As Object
    Set Parent = This.Caller
End Property
  
Public Property Let Parent(RHS As Object)
    Set This.Caller = RHS
End Property
   
Public Property Get Self() As Object
    Set Self = Me
End Property

Private Sub Class_Initialize()
    Set This.EventSink = New Sink
    With This.EventSink
        .Parent = Me
    End With
End Sub

Private Sub Class_Terminate()
    Call Clear
End Sub

Public Sub Clear()
    If (This.Cookie <> 0) Then
        Call ConnectEvent(False)
    End If

    Set This.Control = Nothing
    Set This.EventSink = Nothing
End Sub

Private Sub ConnectEvent(ByVal Connect As Boolean)
    Dim IID_IDispatch As GUID

    ' GUID {00020400-0000-0000-C000000000000046}
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    Call ConnectToConnectionPoint(This.EventSink, _
                                  IID_IDispatch, _
                                  Connect, _
                                  This.Control, _
                                  This.Cookie, _
                                  0&)
End Sub

'--------------------コールバック関数群
Private Sub ISink_OnAfterUpdate()
    Call Parent.OnAfterUpdate(Item)
End Sub

Private Sub ISink_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    Call Parent.OnBeforeUpdate(Item, Cancel)
End Sub

Private Sub ISink_OnChange()
    Call Parent.OnChange(Item)
End Sub

Private Sub ISink_OnClick()
    Call Parent.OnClick(Item)
End Sub

Private Sub ISink_OnDblClick(ByVal Cancel As MSForms.IReturnBoolean)
    Call Parent.OnDblClick(Item, Cancel)
End Sub

Private Sub ISink_OnDropButtonClick()
    Call Parent.OnDropButtonClick(Item)
End Sub

Private Sub ISink_OnEnter()
    Call Parent.OnEnter(Item)
End Sub

Private Sub ISink_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    Call Parent.OnExit(Item, Cancel)
End Sub

Private Sub ISink_OnKeyDown(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    Call Parent.OnKeyDown(Item, KeyCode, Shift)
End Sub

Private Sub ISink_OnKeyPress(ByVal KeyAscii As MSForms.IReturnInteger)
    Call Parent.OnKeyPress(Item, KeyAscii)
End Sub

Private Sub ISink_OnKeyUp(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    Call Parent.OnKeyUp(Item, KeyCode, Shift)
End Sub

Private Sub ISink_OnListClick()
    Call Parent.OnListClick(Item)
End Sub

Private Sub ISink_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call Parent.OnMouseDown(Item, Button, Shift, x, y)
End Sub

Private Sub ISink_OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call Parent.OnMouseMove(Item, Button, Shift, x, y)
End Sub

Private Sub ISink_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call Parent.OnMouseUp(Item, Button, Shift, x, y)
End Sub

Private Sub ISink_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Call Parent.OnBeforeDragOver(Item, Cancel, Data, x, y, DragState, Effect, Shift)
End Sub

Private Sub ISink_OnBeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal x As Single, _
                               ByVal y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    Call Parent.OnBeforeDropOrPaste(Item, Cancel, Action, Data, x, y, Effect, Shift)
End Sub

Private Sub ISink_OnError(ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    Call Parent.OnError(Item, Number, Description, SCode, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub ISink_AddControl(ByVal Control As MSForms.Control)
    Call Parent.AddControl(Item, Control)
End Sub

Private Sub ISink_Layout()
    Call Parent.Layout(Item)
End Sub

Private Sub ISink_RemoveControl(ByVal Control As MSForms.Control)
    Call Parent.RemoveControl(Item, Control)
End Sub

Private Sub ISink_Scroll(ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    Call Parent.Scroll(Item, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

'' ScrollBar
'Private Sub ISink_OnScroll()
'    Call Parent.OnScroll
'End Sub

Private Sub ISink_Zoom(Percent As Integer)
    Call Parent.OnZoom(Item, Percent)
End Sub

Private Sub ISink_OnSpinDown()
    Call Parent.OnSpinDown(Item)
End Sub

Private Sub ISink_OnSpinUp()
    Call Parent.OnSpinUp(Item)
End Sub


