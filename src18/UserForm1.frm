VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7660
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6960
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IControlEvent

Private Type Field
    Controls As ControlEvents
    PrevControl As ControlEvent
End Type

Private This As Field
Public tBox As clsTxtControl
Public cBox As clsTxt2


Private Property Get IControlEvent_PrevControl() As ControlEvent
    Set IControlEvent_PrevControl = This.PrevControl
End Property

Private Property Let IControlEvent_PrevControl(RHS As ControlEvent)
    Set This.PrevControl = RHS
End Property

'Private Sub btnSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    btnSaveMm.Visible = True
'End Sub

'Private Sub btnSaveMm_click()
'    Call tBox.ControlTextBox(UserForm1)
'End Sub

'Private Sub btnSaveMm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Call MouseMoveIcon
'End Sub

Private Sub UserForm_Initialize()
    Me.Height = 360
    Me.Width = 389
    
    Set tBox = New clsTxtControl
    Set cBox = New clsTxt2
    
    Call TxtColor(RGB(55, 55, 55), RGB(0, 182, 233), RGB(166, 166, 166))
    Call cBox.clasBox(UserForm1)
    ComboBox1.List = Array("One", "Two", "Three")
    Set This.Controls = New ControlEvents
    With This.Controls
        .ParentForm = Me
        .Init
    End With
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call AllMouseOut
'    btnSaveMm.Visible = False
End Sub


Private Sub UserForm_Terminate()
    Set tBox = Nothing
    Set cBox = Nothing
End Sub

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(CtrlEvt As ControlEvent, _
                                         ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print CtrlEvt.Control.Name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " Change"
End Sub

Private Sub IControlEvent_OnClick(CtrlEvt As ControlEvent)
    Call ClickProcedure(CtrlEvt)
    Debug.Print CtrlEvt.Control.Name & " Click"
End Sub

Private Sub IControlEvent_OnDblClick(CtrlEvt As ControlEvent, _
                                     ByVal Cancel As MSForms.IReturnBoolean)
'    Select Case CtrlEvt.Control.Name
'        Case "FlatButtonPagePrev", "FlatButtonPageNext"
'            Call IControlEvent_OnClick(CtrlEvt)
'            DoEvents
'            Cancel = True
'    End Select
'    Debug.Print CtrlEvt.Control.Name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(CtrlEvt As ControlEvent)
    Debug.Print CtrlEvt.Control.Name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(CtrlEvt As ControlEvent, _
                                 ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print CtrlEvt.Control.Name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(CtrlEvt As ControlEvent, _
                                    ByVal KeyCode As MSForms.IReturnInteger, _
                                    ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call AllMouseOut
        Unload Me
    End If
    Debug.Print CtrlEvt.Control.Name & " KeyDown:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnKeyPress(CtrlEvt As ControlEvent, _
                                     ByVal KeyAscii As MSForms.IReturnInteger)
    Debug.Print CtrlEvt.Control.Name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(CtrlEvt As ControlEvent, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
    Debug.Print CtrlEvt.Control.Name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseDown:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnMouseMove(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseMove:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnMouseUp(CtrlEvt As ControlEvent, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseUp:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
End Sub

Private Sub IControlEvent_OnBeforeDragOver(CtrlEvt As ControlEvent, _
                                           ByVal Cancel As MSForms.ReturnBoolean, _
                                           ByVal Data As MSForms.DataObject, _
                                           ByVal X As Single, _
                                           ByVal Y As Single, _
                                           ByVal DragState As MSForms.fmDragState, _
                                           ByVal Effect As MSForms.ReturnEffect, _
                                           ByVal Shift As Integer)
'    Debug.Print CtrlEvt.Control.Name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(CtrlEvt As ControlEvent, _
                                              ByVal Cancel As MSForms.ReturnBoolean, _
                                              ByVal Action As MSForms.fmAction, _
                                              ByVal Data As MSForms.DataObject, _
                                              ByVal X As Single, _
                                              ByVal Y As Single, _
                                              ByVal Effect As MSForms.ReturnEffect, _
                                              ByVal Shift As Integer)
'    Debug.Print CtrlEvt.Control.Name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(CtrlEvt As ControlEvent, _
                                  ByVal Number As Integer, _
                                  ByVal Description As MSForms.ReturnString, _
                                  ByVal SCode As Long, _
                                  ByVal Source As String, _
                                  ByVal HelpFile As String, _
                                  ByVal HelpContext As Long, _
                                  ByVal CancelDisplay As MSForms.ReturnBoolean)
'    Debug.Print CtrlEvt.Control.Name & " Error:"
End Sub

Private Sub IControlEvent_OnAddControl(CtrlEvt As ControlEvent, _
                                       ByVal Control As MSForms.Control)
'    Debug.Print CtrlEvt.Control.Name & " AddControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnLayout(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Layout"
End Sub

Private Sub IControlEvent_OnRemoveControl(CtrlEvt As ControlEvent, _
                                          ByVal Control As MSForms.Control)
'    Debug.Print CtrlEvt.Control.Name & " RemoveControl:" & Control.Name
End Sub

Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent, _
                                   ByVal ActionX As MSForms.fmScrollAction, _
                                   ByVal ActionY As MSForms.fmScrollAction, _
                                   ByVal RequestDx As Single, _
                                   ByVal RequestDy As Single, _
                                   ByVal ActualDx As MSForms.ReturnSingle, _
                                   ByVal ActualDy As MSForms.ReturnSingle)
'    Debug.Print CtrlEvt.Control.Name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " Scroll"
'End Sub

Private Sub IControlEvent_OnZoom(CtrlEvt As ControlEvent, _
                                 Percent As Integer)
'    Debug.Print CtrlEvt.Control.Name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(CtrlEvt As ControlEvent)
'    Debug.Print CtrlEvt.Control.Name & " SpinUp"
End Sub

Private Sub IControlEvent_OnMouseOver(CtrlEvt As ControlEvent, _
                                      ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseOver:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
    Call MouseOver(CtrlEvt, Button, Shift, X, Y)
    Set This.PrevControl = CtrlEvt
End Sub

Private Sub IControlEvent_OnMouseOut(CtrlEvt As ControlEvent, _
                                     ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)
    Debug.Print CtrlEvt.Control.Name & " MouseOut:" & Button & "(" & Shift & ") (" & X & "," & Y & ")"
    Call MouseOut(CtrlEvt, Button, Shift, X, Y)
End Sub

Private Sub MouseOver(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With CtrlEvt
        Select Case True
            Case .Attributes.Exists("Picker")
                Me.Controls(VBA.Replace$(.Control.Name, "Picker", "") & "BG").BackColor = &H808080
            Case .Control.Name = "btnSave"
                .Control.Object.Picture = btnSaveMm.Picture
                Call MouseMoveIcon

            Case Else
        End Select
    End With
End Sub

Private Sub MouseOut(CtrlEvt As ControlEvent, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With This.PrevControl
        Select Case True
            Case .Attributes.Exists("Picker")
                Me.Controls(VBA.Replace$(.Control.Name, "Picker", "") & "BG").BackColor = &HFFFFFF
            Case .Control.Name = "btnSave"
                .Control.Object.Picture = btnSaveBase.Picture
            Case Else
        End Select
    End With
End Sub

Private Sub AllMouseOut()
    If Not This.PrevControl Is Nothing Then
        Call IControlEvent_OnMouseOut(This.PrevControl, 0, 0, 0, 0)
        Set This.PrevControl = Nothing
    End If
End Sub

Private Sub ClickProcedure(CtrlEvt As ControlEvent)
    Select Case True
        Case CtrlEvt.Control.Name = "btnSave"
            Call tBox.ControlTextBox(UserForm1)
        Case Else
    End Select
End Sub

Private Sub RegisterCpntrols()
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        Select Case True
            Case VBA.Left$(ctrl.Name, 10) = "FlatButton"
                Call This.Controls.RegisterControl(ctrl, "FlatButton")
            Case VBA.Left$(ctrl.Name, 6) = "Picker"
                Call This.Controls.RegisterControl(ctrl, "Picker")
            Case Else
                Call This.Controls.RegisterControl(ctrl)
        End Select
    Next ctrl
End Sub
