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

'Public Enum Style
'    Base = 1
'    Highlight = 2
'    Click = 3
'End Enum

Private Type Field
    ParentForm As IControlEvent
    EventSink As Sink
    Control As MSForms.Control
    Attributes As Dictionary
'    BaseStyle As ControlStyle
'    HoverStyle As ControlStyle
'    ActiveStyle As ControlStyle
    Cookie As Long
End Type

Private This As Field

Public Property Get Control() As MSForms.Control
Attribute Control.VB_UserMemId = 0
'Attribute Control.VB_UserMemId = 0
    Set Control = This.Control
End Property

Public Property Let Control(RHS As MSForms.Control)
    Set This.Control = RHS
    Call ConnectEvent(True)
End Property

Public Property Get ParentForm() As IControlEvent
    Set ParentForm = This.ParentForm
End Property

Public Property Let ParentForm(RHS As IControlEvent)
    Set This.ParentForm = RHS
End Property

'Private Property Get PrevControl() As ControlEvent
'    If This.ParentForm Is Nothing Then Exit Property
'    Set PrevControl = This.ParentForm.PrevControl
'End Property
'
'Private Property Let PrevControl(RHS As ControlEvent)
'    This.ParentForm.PrevControl = RHS
'End Property

Public Property Get Attributes() As Dictionary
    Set Attributes = This.Attributes
End Property

Public Property Let Attributes(RHS As Dictionary)
    Set This.Attributes = RHS
End Property

Public Property Get AttributeItems() As String
    Dim Key As Variant
    For Each Key In This.Attributes.Keys
        AttributeItems = AttributeItems & Key & ":" & This.Attributes(Key) & ";"
    Next Key
    If Len(AttributeItems) > 0 Then
        AttributeItems = Left(AttributeItems, Len(AttributeItems) - 1)
    End If
End Property

Public Property Let AttributeItems(RHS As String)
    Dim ItemPair() As String
    ItemPair = Split(RHS, ";")
    Dim Pair As Variant
    Dim splitTmp() As String
    For Each Pair In ItemPair
        splitTmp = Split(Pair, ":")
        If UBound(splitTmp) > 0 Then
            This.Attributes.Add splitTmp(0), splitTmp(1)
        Else
            This.Attributes.Add splitTmp(0), ""
        End If
    Next Pair
End Property

Public Property Get AttributeItem(Key As String) As String
    AttributeItem = This.Attributes(Key)
End Property

Public Property Let AttributeItem(Key As String, RHS As String)
    This.Attributes(Key) = RHS
End Property

'Public Property Get BaseStyle() As ControlStyle
'    Set BaseStyle = This.BaseStyle
'End Property
'
'Public Property Let BaseStyle(RHS As ControlStyle)
'    Set This.BaseStyle = RHS
'End Property
'
'Public Property Get HoverStyle() As ControlStyle
'    Set HoverStyle = This.HoverStyle
'End Property
'
'Public Property Let HoverStyle(RHS As ControlStyle)
'    Set This.HoverStyle = RHS
'End Property
'
'Public Property Get ActiveStyle() As ControlStyle
'    Set ActiveStyle = This.ActiveStyle
'End Property
'
'Public Property Let ActiveStyle(RHS As ControlStyle)
'    Set This.ActiveStyle = RHS
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
    Set This.Attributes = New Dictionary
'    Set This.BaseStyle = New ControlStyle
'    Set This.HoverStyle = New ControlStyle
'    Set This.ActiveStyle = New ControlStyle
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
    Set This.Attributes = Nothing
'    Set This.BaseStyle = Nothing
'    Set This.HoverStyle = Nothing
'    Set This.ActiveStyle = Nothing
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
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnAfterUpdate(Me)
End Sub

Private Sub ISink_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnBeforeUpdate(Me, Cancel)
End Sub

Private Sub ISink_OnChange()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnChange(Me)
End Sub

Private Sub ISink_OnClick()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnClick(Me)
End Sub

Private Sub ISink_OnDblClick(ByVal Cancel As MSForms.IReturnBoolean)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnDblClick(Me, Cancel)
End Sub

Private Sub ISink_OnDropButtonClick()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnDropButtonClick(Me)
End Sub

Private Sub ISink_OnEnter()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnEnter(Me)
End Sub

Private Sub ISink_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnExit(Me, Cancel)
End Sub

Private Sub ISink_OnKeyDown(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    If ParentForm Is Nothing Then Exit Sub
    Debug.Print TypeName(ParentForm) & ":KeyDown(" & KeyCode & ")"
    Call ParentForm.OnKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub ISink_OnKeyPress(ByVal KeyAscii As MSForms.IReturnInteger)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnKeyPress(Me, KeyAscii)
End Sub

Private Sub ISink_OnKeyUp(ByVal KeyCode As MSForms.IReturnInteger, ByVal Shift As Integer)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnKeyUp(Me, KeyCode, Shift)
End Sub

Private Sub ISink_OnListClick()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnListClick(Me)
End Sub

Private Sub ISink_OnMouseDown(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal x As Single, _
                              ByVal y As Single)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnMouseDown(Me, Button, Shift, x, y)
End Sub

Private Sub ISink_OnMouseMove(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal x As Single, _
                              ByVal y As Single)
    If ParentForm Is Nothing Then Exit Sub
    If Not ParentForm.PrevControl Is Nothing Then
        If Not ParentForm.PrevControl Is Me Then
            Call ParentForm.OnMouseOut(ParentForm.PrevControl, Button, Shift, x, y)
            Call ParentForm.OnMouseOver(Me, Button, Shift, x, y)
        Else
        
        End If
    Else
        Call ParentForm.OnMouseOver(Me, Button, Shift, x, y)
    End If
    Call ParentForm.OnMouseMove(Me, Button, Shift, x, y)
End Sub

Private Sub ISink_OnMouseUp(ByVal Button As Integer, _
                            ByVal Shift As Integer, _
                            ByVal x As Single, _
                            ByVal y As Single)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnMouseUp(Me, Button, Shift, x, y)
End Sub

Private Sub ISink_OnBeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
                                   ByVal Data As MSForms.DataObject, _
                                   ByVal x As Single, _
                                   ByVal y As Single, _
                                   ByVal DragState As MSForms.fmDragState, _
                                   ByVal Effect As MSForms.ReturnEffect, _
                                   ByVal Shift As Integer)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnBeforeDragOver(Me, Cancel, Data, x, y, DragState, Effect, Shift)
End Sub

Private Sub ISink_OnBeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
                                      ByVal Action As MSForms.fmAction, _
                                      ByVal Data As MSForms.DataObject, _
                                      ByVal x As Single, _
                                      ByVal y As Single, _
                                      ByVal Effect As MSForms.ReturnEffect, _
                                      ByVal Shift As Integer)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnBeforeDropOrPaste(Me, Cancel, Action, Data, x, y, Effect, Shift)
End Sub

Private Sub ISink_OnError(ByVal Number As Integer, _
                          ByVal Description As MSForms.ReturnString, _
                          ByVal SCode As Long, ByVal Source As String, _
                          ByVal HelpFile As String, _
                          ByVal HelpContext As Long, _
                          ByVal CancelDisplay As MSForms.ReturnBoolean)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnError(Me, Number, Description, SCode, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub ISink_OnAddControl(ByVal Control As MSForms.Control)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnAddControl(Me, Control)
End Sub

Private Sub ISink_OnLayout()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnLayout(Me)
End Sub

Private Sub ISink_OnRemoveControl(ByVal Control As MSForms.Control)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnRemoveControl(Me, Control)
End Sub

Private Sub ISink_OnScroll(ByVal ActionX As MSForms.fmScrollAction, _
                           ByVal ActionY As MSForms.fmScrollAction, _
                           ByVal RequestDx As Single, _
                           ByVal RequestDy As Single, _
                           ByVal ActualDx As MSForms.ReturnSingle, _
                           ByVal ActualDy As MSForms.ReturnSingle)
    If ParentForm Is Nothing Then Exit Sub
    Debug.Print This.Control.Name & " Scroll:"
    Call ParentForm.OnScroll(Me, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

'' ScrollBar
'Private Sub ISink_OnScroll()
'    Call ParentForm.OnScroll(Me)
'End Sub

Private Sub ISink_OnZoom(Percent As Integer)
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnZoom(Me, Percent)
End Sub

Private Sub ISink_OnSpinDown()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnSpinDown(Me)
End Sub

Private Sub ISink_OnSpinUp()
    If ParentForm Is Nothing Then Exit Sub
    Call ParentForm.OnSpinUp(Me)
End Sub

'Private Sub ISink_OnMouseOver(ByVal Button As Integer, _
'                              ByVal Shift As Integer)
'    If ParentForm Is Nothing Then Exit Sub
'    Call ParentForm.OnMouseOver(Me, Button, Shift)
'End Sub
'
'Private Sub ISink_OnMouseOut(ByVal Button As Integer, _
'                             ByVal Shift As Integer)
'    If ParentForm Is Nothing Then Exit Sub
'    Call ParentForm.OnMouseOut(Me, Button, Shift)
'End Sub
