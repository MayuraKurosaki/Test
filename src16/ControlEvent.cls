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

Private Type Field
    Caller As IControlEvent
    EventSink As Sink
    Control As Object
    Atributes As Dictionary
    Cookie As Long
End Type

Private This As Field

Public Property Get Control() As Object
Attribute Control.VB_UserMemId = 0
'Attribute Control.VB_UserMemId = 0
    Set Control = This.Control
End Property

Public Property Let Control(RHS As Object)
    Set This.Control = RHS
    Call ConnectEvent(True)
End Property

Public Property Get Parent() As Object
    Set Parent = This.Caller
End Property

Public Property Let Parent(RHS As Object)
    Set This.Caller = RHS
End Property

Public Property Get Atributes() As Dictionary
    Set Atributes = This.Atributes
End Property

Public Property Let Atributes(RHS As Dictionary)
    Set This.Atributes = RHS
End Property

Public Property Get AtributeItems() As String
    Set Atributes = This.Atributes
End Property

Public Property Let AtributeItems(RHS As String)
    Set This.Atributes = RHS
End Property

'Public Property Get AtributeItem() As String
'    Set Atributes = This.Atributes
'End Property
'
'Public Property Let AtributeItem(Key As String, RHS As String)
'    Set This.Atributes = RHS
'End Property

Public Property Get Self() As Object
    Set Self = Me
End Property

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set This.EventSink = New Sink
    With This.EventSink
        .Parent = Me
    End With
    Set This.Atributes = New Dictionary
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
    Set This.Atributes = Nothing
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
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnAfterUpdate(Me)
End Sub

Private Sub ISink_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnBeforeUpdate(Me, Cancel)
End Sub

Private Sub ISink_OnChange()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnChange(Me)
End Sub

Private Sub ISink_OnClick()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnClick(Me)
End Sub

Private Sub ISink_OnDblClick(ByVal Cancel As MSForms.IReturnBoolean)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnDblClick(Me, Cancel)
End Sub

Private Sub ISink_OnDropButtonClick()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnDropButtonClick(Me)
End Sub

Private Sub ISink_OnEnter()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnEnter(Me)
End Sub

Private Sub ISink_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnExit(Me, Cancel)
End Sub

Private Sub ISink_OnKeyDown(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    If Parent Is Nothing Then Exit Sub
    Debug.Print TypeName(Parent) & ":KeyDown(" & KeyCode & ")"
    Call Parent.OnKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub ISink_OnKeyPress(ByVal KeyAscii As MSForms.IReturnInteger)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnKeyPress(Me, KeyAscii)
End Sub

Private Sub ISink_OnKeyUp(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnKeyUp(Me, KeyCode, Shift)
End Sub

Private Sub ISink_OnListClick()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnListClick(Me)
End Sub

Private Sub ISink_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnMouseDown(Me, Button, Shift, X, Y)
End Sub

Private Sub ISink_OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnMouseMove(Me, Button, Shift, X, Y)
End Sub

Private Sub ISink_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnMouseUp(Me, Button, Shift, X, Y)
End Sub

Private Sub ISink_OnBeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnBeforeDragOver(Me, Cancel, Data, X, Y, DragState, Effect, Shift)
End Sub

Private Sub ISink_OnBeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnBeforeDropOrPaste(Me, Cancel, Action, Data, X, Y, Effect, Shift)
End Sub

Private Sub ISink_OnError(ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnError(Me, Number, Description, SCode, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub ISink_OnAddControl(ByVal Control As MSForms.Control)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnAddControl(Me, Control)
End Sub

Private Sub ISink_OnLayout()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnLayout(Me)
End Sub

Private Sub ISink_OnRemoveControl(ByVal Control As MSForms.Control)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnRemoveControl(Me, Control)
End Sub

Private Sub ISink_OnScroll(ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    If Parent Is Nothing Then Exit Sub
    Debug.Print This.Control.Name & " Scroll:"
    Call Parent.OnScroll(Me, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

'' ScrollBar
'Private Sub ISink_OnScroll()
'    Call Parent.OnScroll(Me)
'End Sub

Private Sub ISink_OnZoom(Percent As Integer)
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnZoom(Me, Percent)
End Sub

Private Sub ISink_OnSpinDown()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnSpinDown(Me)
End Sub

Private Sub ISink_OnSpinUp()
    If Parent Is Nothing Then Exit Sub
    Call Parent.OnSpinUp(Me)
End Sub


