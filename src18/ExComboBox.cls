VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents myForm As MSForms.UserForm
Attribute myForm.VB_VarHelpID = -1
Public WithEvents myFormControl As MSForms.Frame
Attribute myFormControl.VB_VarHelpID = -1
Public WithEvents mpage As MSForms.MultiPage
Attribute mpage.VB_VarHelpID = -1
Public WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1
Public WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents barBottom As MSForms.Label
Attribute barBottom.VB_VarHelpID = -1
Public WithEvents barTitle As MSForms.Label
Attribute barTitle.VB_VarHelpID = -1
Public WithEvents DropArrow As MSForms.Label
Attribute DropArrow.VB_VarHelpID = -1
Public WithEvents ErrLabel As MSForms.Label
Attribute ErrLabel.VB_VarHelpID = -1
'Public txt As New Collection
'Public TxtBox As New clsTxtControl
Public ctbox As New Collection
Public txtKontrol
Public fColor, eColor, tColor

Private Const FONT_NAME As String = "Yu Gothic UI"
Private Const FONT_SIZE As Currency = 10

Public Sub TxtColor(fColorValue As String, eColorValue As String, tColorValue As String)
    fColor = fColorValue 'TextBox ForeColor
    eColor = eColorValue 'When TextBox Enter
    tColor = tColorValue 'Title and bottom line Color
End Sub

Public Sub clasBox(Form As MSForms.UserForm)
    
'    FontSize = 10: FontName = "Yu Gothic UI"
    Dim ctr As Control
    Set myForm = Form

'************************************************************************************************************
    '////TEXTBOX IN FRAME/////////////////
    For Each ctr In myForm.Controls
        If TypeName(ctr) = "Frame" Then
            Dim ctrl As Control
            For Each ctrl In ctr.Controls
                If TypeName(ctrl) = "TextBox" Then
                    With ctrl
                        Call AddTextBox(ctrl)
                        
                        Set barBottom = ctr.Controls.Add("Forms.Label.1", .Name & "barBottom")
                            Call AddBarBottom(ctrl)

                        Set barTitle = ctr.Controls.Add("Forms.Label.1", .Name & "barTitle")
                            Call AddBarTitle(ctrl)

                    End With
                    
                    Set txt = New clsTxt2
                    Set txt.TextBox = ctrl
                    Set txt.myFormControl = ctr
                    Set txt.myForm = myForm
                    ctbox.Add txt
                    
                '////COMBOBOX IN FRAME/////////////////
                ElseIf TypeName(ctrl) = "ComboBox" Then
                    With ctrl
                        Call AddComboBox(ctrl)
               
                        Set barBottom = ctr.Controls.Add("Forms.Label.1", .Name & "barBottom")
                            Call AddBarBottom(ctrl)
    
                        Set barTitle = ctr.Controls.Add("Forms.Label.1", .Name & "barTitle")
                            Call AddBarTitle(ctrl)

                        Set DropArrow = ctr.Controls.Add("Forms.Label.1", .Name & "DropArrow")
                            Call AddDropArrow(ctrl)
                    
                    End With
                    
                    Set txt = New clsTxt2
                    Set txt.ComboBox = ctrl
                    Set txt.DropArrow = DropArrow
                    Set txt.myFormControl = ctr
                    Set txt.myForm = myForm
                    ctbox.Add txt
                End If
        
            Next ctrl
            
            Set txt = New clsTxt2
            Set txt.myFormControl = ctr
            Set txt.myForm = myForm
            ctbox.Add txt
                
        End If
    Next ctr
            
'************************************************************************************************************
    '///TEXTBOX IN MULTIPAGE/////////
    For Each mp In myForm.Controls
        If TypeName(mp) = "MultiPage" Then
            Dim mpc As Control
            mc = mp.Pages.Count
            For i = 0 To mc - 1
                For Each mpc In mp.Pages(i).Controls
                    If TypeName(mpc) = "TextBox" Then
                        With mpc
                        Call AddTextBox(mpc)
                          
                            Set barBottom = mp.Pages(i).Controls.Add("Forms.Label.1", .Name & "barBottom")
                            Call AddBarBottom(mpc)
                            
                            Set barTitle = mp.Pages(i).Controls.Add("Forms.Label.1", .Name & "barTitle")
                            Call AddBarTitle(mpc)
    
                        End With
                        
                        Set txt = New clsTxt2
                        Set txt.TextBox = mpc
                        Set txt.myForm = myForm
                        ctbox.Add txt
                    
                    '////COMBOBOX IN MULTIPAGE/////////////////
                    ElseIf TypeName(mpc) = "ComboBox" Then
                        With mpc
                            Call AddComboBox(mpc)
                    
                            Set barBottom = mp.Pages(i).Controls.Add("Forms.Label.1", .Name & "barBottom")
                                Call AddBarBottom(mpc)
        
                            Set barTitle = mp.Pages(i).Controls.Add("Forms.Label.1", .Name & "barTitle")
                                Call AddBarTitle(mpc)
                            
                            Set DropArrow = mp.Pages(i).Controls.Add("Forms.Label.1", .Name & "DropArrow")
                                Call AddDropArrow(mpc)
                    
                        End With
                        
                        Set txt = New clsTxt2
                        Set txt.ComboBox = mpc
                        Set txt.DropArrow = DropArrow
                        Set txt.mpage = mp
                        Set txt.myForm = myForm
                        ctbox.Add txt
                    
                    End If
                Next mpc
            Next i
            
            Set txt = New clsTxt2
            Set txt.mpage = mp
            Set txt.myForm = myForm
            ctbox.Add txt
                
        End If
    Next mp
            
'************************************************************************************************************
    '/////OTHER THAN TEXTBOX IN FRAME AND MULTIPAGE/////////////////
    Dim uctrl As Control
    For Each uctrl In myForm.Controls
        If TypeName(uctrl) = "TextBox" And uctrl.ControlTipText <> "cBox" Then
            With uctrl
                Call AddTextBox(uctrl)
                
                Set barBottom = myForm.Controls.Add("Forms.Label.1", .Name & "barBottom")
                    Call AddBarBottom(uctrl)

                Set barTitle = myForm.Controls.Add("Forms.Label.1", .Name & "barTitle")
                    Call AddBarTitle(uctrl)

            End With
            Set txt = New clsTxt2
            Set txt.TextBox = uctrl
            Set txt.myForm = myForm
            ctbox.Add txt
         
        '/////OTHER THAN COMBOBOX IN FRAME AND MULTIPAGE/////////////////
        ElseIf TypeName(uctrl) = "ComboBox" And uctrl.ControlTipText <> "cBox" Then
            With uctrl
                Call AddComboBox(uctrl)

                Set barBottom = myForm.Controls.Add("Forms.Label.1", .Name & "barBottom")
                    Call AddBarBottom(uctrl)

                Set barTitle = myForm.Controls.Add("Forms.Label.1", .Name & "barTitle")
                    Call AddBarTitle(uctrl)

                Set DropArrow = myForm.Controls.Add("Forms.Label.1", .Name & "DropArrow")
                    Call AddDropArrow(uctrl)
        
            End With
            
            Set txt = New clsTxt2
            Set txt.ComboBox = uctrl
            Set txt.DropArrow = DropArrow
            Set txt.myForm = myForm
            ctbox.Add txt
                            
        End If
    Next uctrl

End Sub

Sub AddTextBox(ctrl As Control)
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = True
        .Font.Name = FONT_NAME
        .ForeColor = fColor
        .Font.size = FONT_SIZE
        .ControlTipText = "cBox"
        .Height = 24
    End With
End Sub

Sub AddComboBox(ctrl As Control)
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = True
        .Font.Name = FONT_NAME
        .ForeColor = fColor
        .ShowDropButtonWhen = fmShowDropButtonWhenNever
        .Font.size = FONT_SIZE
        .Height = 24
        .Left = ctrl.Left - 4
        .ControlTipText = "cBox"
    End With
End Sub

Sub AddBarBottom(ctrl As Control)
    With barBottom
        .Height = 30
        .Width = ctrl.Width
        .BorderStyle = fmBorderStyleSingle
        .Left = ctrl.Left
        .BackStyle = fmBackStyleTransparent
        .BorderColor = tColor
        .Top = ctrl.Top - 6
        .ZOrder (1)
    End With
End Sub

Sub AddBarTitle(ctrl As Control)
    With barTitle
        .Height = 16 ' ctrl.Height
        .AutoSize = True
        .WordWrap = False
        .Left = ctrl.Left + 4
        .Caption = Space(1) & ctrl.Tag
        .ForeColor = fColor
        .Font.Name = FONT_NAME
        .Font.size = FONT_SIZE
        .Top = ctrl.Top + 2
        .ZOrder (1)
    End With
End Sub

Sub AddDropArrow(ctrl As Control)
    With DropArrow
        .Font.Name = "Webdings"
        .Caption = "6"
        .Height = ctrl.Height
        .Width = .Height
        .Left = ctrl.Width + ctrl.Left - .Width + 8
        .BackStyle = fmBackStyleTransparent
        .ForeColor = fColor
        .BorderStyle = fmBorderStyleNone
        .Top = ctrl.Top + 3
        .Font.size = FONT_SIZE
        .ZOrder (0)
        
    End With
End Sub

Private Sub TextBox_Change()
    Debug.Print TypeName(Me) & ":TextBox_Change"
    TextBoxEnter
End Sub

Private Sub TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print TypeName(Me) & ":TextBox_MouseDown"
    TextBoxEnter
End Sub

Private Sub TextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":TextBox_KeyUp"
    TextBoxEnter
End Sub

Private Sub txtCaption_Click()
    Debug.Print TypeName(Me) & ":txtCaption_Click"
    BoxExit
End Sub

Sub TextBoxEnter()
    Debug.Print TypeName(Me) & ":TextBoxEnter"
    BoxExit
    Dim Control As Control
    For Each Control In myForm.Controls
        If TypeName(Control) = "TextBox" And Control.Name = TextBox.Name Then
            With myForm.Controls(TextBox.Name & "barBottom")
                .BorderColor = eColor

            End With
            With myForm.Controls(TextBox.Name & "barTitle")
                .ForeColor = eColor
                .Top = TextBox.Top - 10.5
                .Font.size = 8
                .Height = 8
                 If Len(.Caption) * 9 - 40 > 40 Then
                    .Width = Len(.Caption) * 9 - 40
                 Else
                    .Width = 36
                 End If
                .Left = TextBox.Left + 4
                .BackStyle = fmBackStyleOpaque
                .BackColor = RGB(240, 240, 240)
                .ZOrder (0)
            End With
        End If
    Next Control
End Sub

Sub BoxExit()
    Debug.Print TypeName(Me) & ":BoxExit"
    Dim Control As Control

    For Each Control In myForm.Controls
        If TypeName(Control) = "TextBox" Or TypeName(Control) = "ComboBox" Then
            If Len(Control) = 0 Then
                  With myForm.Controls(Control.Name & "barBottom")
                    .BorderColor = tColor
                 End With
                 With myForm.Controls(Control.Name & "barTitle")
                     .ForeColor = fColor
                     .Top = Control.Top + 2
                     .Left = Control.Left + 4
                     .Height = 16
                     .Font.size = FONT_SIZE
                     .Width = Len(.Caption) * 6 + 24
                     .BackStyle = fmBackStyleOpaque
                     .BackColor = RGB(240, 240, 240)
                     .ZOrder (1)
                 End With
             Else
                 With myForm.Controls(Control.Name & "barBottom")
                    .BorderColor = tColor
                 End With
                 With myForm.Controls(Control.Name & "barTitle")
                     .ForeColor = tColor
                 End With
             End If
         End If
            
        If TypeName(Control) = "ComboBox" Then
            With myForm.Controls(Control.Name & "DropArrow")
                .ForeColor = tColor
            End With
        End If
    Next
End Sub

Sub ComboBoxEnter()
    Debug.Print TypeName(Me) & ":ComboBoxEnter"
    BoxExit
    Dim Control As Control
    For Each Control In myForm.Controls

        If TypeName(Control) = "ComboBox" And Control.Name = ComboBox.Name Then
            
            With myForm.Controls(ComboBox.Name & "barBottom")
                .BorderColor = eColor
            End With
            With myForm.Controls(ComboBox.Name & "barTitle")
                .ForeColor = eColor
                .Top = ComboBox.Top - 10.5
                .Font.size = 8
                .Height = 8
                .Width = Len(.Caption) * 4 + 12
                .Left = ComboBox.Left + 4
                .BackStyle = fmBackStyleOpaque
                .BackColor = RGB(240, 240, 240)
                .ZOrder (0)
            End With
            
            With myForm.Controls(ComboBox.Name & "DropArrow")
                .ForeColor = eColor
            End With
            Control.DropDown
         End If
    Next
End Sub

Private Sub Class_Terminate()
    Set txt = Nothing
    Set ctbox = Nothing
    Set TxtBox = Nothing
End Sub

Private Sub ComboBox_Change()
    Debug.Print TypeName(Me) & ":ComboBox_Change"
    ComboBoxEnter
End Sub

Private Sub ComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print TypeName(Me) & ":ComboBox_KeyUp"
    ComboBoxEnter
End Sub

Private Sub ComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print TypeName(Me) & ":ComboBox_MouseDown"
    ComboBoxEnter
End Sub

Private Sub DropArrow_Click()
    Debug.Print TypeName(Me) & ":DropArrow_Click"
    BoxExit
    Dim ComboName As String
    Dim ctrl As Control
    For Each ctrl In myForm.Controls
        If TypeName(ctrl) = "ComboBox" Then
            ComboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
            If ComboName = ctrl.Name Then
                With myForm.Controls(ComboName)
                    .DropDown
                    With myForm.Controls(ComboName & "barBottom")
                         .BorderColor = eColor
                     End With
                     
                     With myForm.Controls(ComboName & "barTitle")
                         .ForeColor = eColor
                         .Top = myForm.Controls(ComboName).Top - 10.5
                         .Font.size = 8
                         .Height = 8
                         .Width = Len(.Caption) * 4 + 12
                         .Left = myForm.Controls(ComboName).Left + 4
                         .BackStyle = fmBackStyleOpaque
                         .BackColor = RGB(240, 240, 240)
                         .ZOrder (0)
                     End With
                 End With
                 With DropArrow
                    .ForeColor = eColor
                End With
            End If
        End If
    Next ctrl
End Sub

Private Sub DropArrow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print TypeName(Me) & ":DropArrow_MouseMove"
    Call MouseMoveIcon
End Sub

Private Sub mPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print TypeName(Me) & ":mPage_MouseDown"
    BoxExit
End Sub

Private Sub mPage_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Debug.Print TypeName(Me) & ":mPage_MouseUp"
    BoxExit
End Sub

Private Sub myForm_Click()
    Debug.Print TypeName(Me) & ":myForm_Click"
    BoxExit
End Sub

Private Sub myFormControl_Click()
    Debug.Print TypeName(Me) & ":myFormControl_Click"
    BoxExit
End Sub

Private Sub ErrLabel_Click()
    Debug.Print TypeName(Me) & ":ErrLabel_Click"
    Focus = Mid(ErrLabel.Name, 1, Len(ErrLabel.Name) - 8)
    With myForm.Controls(Focus)
        .SetFocus
    End With
End Sub

Public Sub ControlTextBox(Form As MSForms.UserForm)
    Dim FontName As String
    FontName = "Yu Gothic UI"
    Set txt = Nothing
    Set myForm = Form

    For Each Item In myForm.Controls
        If Right(Item.Name, 8) = "ErrLabel" Then
            myForm.Controls.Remove (Item.Name)
        End If
        If (TypeName(Item) = "TextBox" Or TypeName(Item) = "ComboBox") Then
            txt.Add Item
        End If
    Next Item
    
    For i = 1 To txt.Count
        If txt(i).Text = "" And txt(i).Tag <> "Search" Then
            Set ErrLabel = myForm.Add("Forms.Label.1", txt(i).Name & "ErrLabel")
            With ErrLabel
                .Caption = "*This fields is required"
                .Left = txt(i).Left + 2
                .Top = txt(i).Top + 24
                .Font.size = 8
                .Font.Name = txt(i).Font.Name
                .ForeColor = &HFF
                .BackStyle = fmBackStyleTransparent
                .MousePointer = fmMousePointerIBeam
                .Width = 155
                 txtKontrol = False
            End With
             Set TxtBox = New clsTxtControl
             Select Case TypeName(txt(i))
                Case Is = "TextBox"
                    Set TxtBox.TextBox = txt(i)
                Case Is = "ComboBox"
                    Set TxtBox.ComboBox = txt(i)
             End Select
             Set TxtBox.ErrLabel = ErrLabel
             Set TxtBox.myForm = Form
             txt.Add TxtBox
        Else
            Set ErrLabel = myForm.Add("Forms.Label.1", txt(i).Name & "ErrLabel")
            With ErrLabel
                .Caption = ""
                .BackStyle = fmBackStyleTransparent
            End With
        End If

        Set TxtBox = New clsTxtControl
        Select Case TypeName(txt(i))
               Case Is = "TextBox"
                   Set TxtBox.TextBox = txt(i)
               Case Is = "ComboBox"
                   Set TxtBox.ComboBox = txt(i)
        End Select
        Set TxtBox.myForm = Form
        txt.Add TxtBox
    Next i
End Sub

