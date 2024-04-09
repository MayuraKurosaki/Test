VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlEvent

Private Type TState
    Control As ControlEvents
    prevControl As MSForms.IControl
End Type

Private This As TState
'Private ExForm_ As ExtendedForm

Private Sub UserForm_Initialize()
    Set This.Control = New ControlEvents       ' インスタンスの生成
    With This.Control
        .Parent = Me
        .init
    End With
    FormNonCaption Me, True
'    Call CalendarFormModule.Use
'
'    Set ExForm_ = New ExtendedForm
'    Set ExForm_.Form = Me
'
'    With ExForm_
'        .Maximize = False
'        .Minimize = True
'        .Resize = False
'        .CloseButton = True
'        .Menu = False
''        PrevProc = .SubClass(AddressOf MainViewModule.WindowProc)
'        CalendarFormModule.PrevProc = .SubClass(AddressOf CalendarFormModule.WindowProc)
'        .Redraw
'    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
End Sub

Private Sub UserForm_Terminate()
'    Set ExForm_ = Nothing
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not This.prevControl Is Nothing Then
        This.prevControl.Object.BorderStyle = fmBorderStyleNone
        Set This.prevControl = Nothing
    End If
End Sub

'--------------------インターフェイスからコールバックされるメンバ関数
Private Sub IControlEvent_OnAfterUpdate(Cont As MSForms.IControl)
    Debug.Print Cont.name & " AfterUpdate"
End Sub

Private Sub IControlEvent_OnBeforeUpdate(Cont As MSForms.IControl, _
                                       ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.name & " BeforeUpdate"
End Sub

Private Sub IControlEvent_OnChange(Cont As MSForms.IControl)
    Debug.Print Cont.name & " Change"
End Sub

Private Sub IControlEvent_OnClick(Cont As MSForms.IControl)
    Debug.Print Cont.name & " Click"
End Sub

Private Sub IControlEvent_OnDblClick(Cont As MSForms.IControl, _
                                   ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.name & " DblClick"
End Sub

Private Sub IControlEvent_OnDropButtonClick(Cont As MSForms.IControl)
    Debug.Print Cont.name & " DropButtonClick"
End Sub

Private Sub IControlEvent_OnEnter(Cont As MSForms.IControl)
    Debug.Print Cont.name & " Enter"
End Sub

Private Sub IControlEvent_OnExit(Cont As MSForms.IControl, _
                               ByVal Cancel As MSForms.IReturnBoolean)
    Debug.Print Cont.name & " Exit"
End Sub

Private Sub IControlEvent_OnKeyDown(Cont As MSForms.IControl, _
                                  ByVal KeyCode As MSForms.IReturnInteger, _
                                  ByVal Shift As Integer)
    Debug.Print Cont.name & " KeyDown:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnKeyPress(Cont As MSForms.IControl, _
                                   ByVal KeyAscii As MSForms.IReturnInteger)
    Debug.Print Cont.name & " KeyPress:" & KeyAscii
End Sub

Private Sub IControlEvent_OnKeyUp(Cont As MSForms.IControl, _
                                ByVal KeyCode As MSForms.IReturnInteger, _
                                ByVal Shift As Integer)
    Debug.Print Cont.name & " KeyUp:" & KeyCode & "(" & Shift & ")"
End Sub

Private Sub IControlEvent_OnListClick(Cont As MSForms.IControl)
    Debug.Print Cont.name & " ListClick"
End Sub

Private Sub IControlEvent_OnMouseDown(Cont As MSForms.IControl, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
    Debug.Print Cont.name & " MouseDown:"
End Sub

Private Sub IControlEvent_OnMouseMove(Cont As MSForms.IControl, _
                                    ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal x As Single, _
                                    ByVal y As Single)
    If Cont.name = "Label48" Then Exit Sub
    Select Case TypeName(Cont)
        Case "Label"
            Cont.Object.BorderStyle = fmBorderStyleSingle
            If Not This.prevControl Is Nothing Then
                If Not This.prevControl Is Cont Then
                    This.prevControl.Object.BorderStyle = fmBorderStyleNone
                End If
            End If
            Set This.prevControl = Cont
        Case Else
            If Not This.prevControl Is Nothing Then
                This.prevControl.Object.BorderStyle = fmBorderStyleNone
            End If
    End Select
    Debug.Print Cont.name & " MouseMove:"
End Sub

Private Sub IControlEvent_OnMouseUp(Cont As MSForms.IControl, _
                                  ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal x As Single, _
                                  ByVal y As Single)
    Debug.Print Cont.name & " MouseUp:"
End Sub

Private Sub IControlEvent_BeforeDragOver(Cont As MSForms.Control, _
                            ByVal Cancel As MSForms.ReturnBoolean, _
                            ByVal Data As MSForms.DataObject, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            ByVal DragState As MSForms.fmDragState, _
                            ByVal Effect As MSForms.ReturnEffect, _
                            ByVal Shift As Integer)
    Debug.Print Cont.name & " BeforeDragOver:"
End Sub

Private Sub IControlEvent_OnBeforeDropOrPaste(Cont As MSForms.Control, _
                               ByVal Cancel As MSForms.ReturnBoolean, _
                               ByVal Action As MSForms.fmAction, _
                               ByVal Data As MSForms.DataObject, _
                               ByVal x As Single, _
                               ByVal y As Single, _
                               ByVal Effect As MSForms.ReturnEffect, _
                               ByVal Shift As Integer)
    Debug.Print Cont.name & " BeforeDropOrPaste:"
End Sub

Private Sub IControlEvent_OnError(Cont As MSForms.Control, _
                   ByVal Number As Integer, _
                   ByVal Description As MSForms.ReturnString, _
                   ByVal SCode As Long, _
                   ByVal Source As String, _
                   ByVal HelpFile As String, _
                   ByVal HelpContext As Long, _
                   ByVal CancelDisplay As MSForms.ReturnBoolean)
    Debug.Print Cont.name & " Error:"
End Sub

Private Sub IControlEvent_AddControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
    Debug.Print Cont.name & " AddControl:" & Control.name
End Sub

Private Sub IControlEvent_Layout(Cont As MSForms.Control)
    Debug.Print Cont.name & " Layout"
End Sub

Private Sub IControlEvent_RemoveControl(Cont As MSForms.Control, ByVal Control As MSForms.Control)
    Debug.Print Cont.name & " RemoveControl:" & Control.name
End Sub

Private Sub IControlEvent_Scroll(Cont As MSForms.Control, _
                    ByVal ActionX As MSForms.fmScrollAction, _
                    ByVal ActionY As MSForms.fmScrollAction, _
                    ByVal RequestDx As Single, _
                    ByVal RequestDy As Single, _
                    ByVal ActualDx As MSForms.ReturnSingle, _
                    ByVal ActualDy As MSForms.ReturnSingle)
    Debug.Print Cont.name & " Scroll:"
End Sub

'' ScrollBar
'Private Sub IControlEvent_OnScroll(Cont As MSForms.Control)
'    Debug.Print Cont.Name & " Scroll"
'End Sub

Private Sub IControlEvent_Zoom(Cont As MSForms.Control, Percent As Integer)
    Debug.Print Cont.name & " Zoom:" & Percent & "%"
End Sub

Private Sub IControlEvent_OnSpinDown(Cont As MSForms.Control)
    Debug.Print Cont.name & " SpinDown"
End Sub

Private Sub IControlEvent_OnSpinUp(Cont As MSForms.Control)
    Debug.Print Cont.name & " SpinUp"
End Sub


