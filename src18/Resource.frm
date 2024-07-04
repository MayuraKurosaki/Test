VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Resource 
   Caption         =   "Resource"
   ClientHeight    =   7180
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13020
   OleObjectBlob   =   "Resource.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OnMouseOver As Boolean
Private OnFocus As Boolean
Private FrameTop As Single

Private Sub CommandButton1_Click()
'    DatePicker.Init Me.TextBox1
    DatePicker.ShowPicker Me.TextBox1
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox2_Enter()
    ExTextBox_GotFocus
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus
End Sub

Private Sub TextBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver
End Sub

Private Sub ExTextBox_GotFocus()
    If OnFocus Then Exit Sub
    
    With Me.Frame1
        .Caption = Me.Label2.Caption
        .Top = FrameTop - 6
        FrameTop = .Top
        .Height = 30
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
    Me.Label2.Visible = False
    OnFocus = True
End Sub

Private Sub ExTextBox_LostFocus()
    If Not OnFocus Then Exit Sub
    
    With Me.Frame1
        .Caption = ""
        .Top = FrameTop + 6
        FrameTop = .Top
        .Height = 24
        .BorderColor = &H808080
    End With
    Me.Label2.Visible = True
    OnFocus = False
End Sub

Private Sub ExTextBox_MouseOver()
    If OnMouseOver Then Exit Sub
    
    With Me.Frame1
'        .Caption = Me.Label2.Caption
'        .Top = FrameTop - 6
'        FrameTop = .Top
'        .Height = 36
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
'    Me.Label2.Visible = False
    OnMouseOver = True
End Sub

Private Sub ExTextBox_MouseOut()
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    With Me.Frame1
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 30
        .BorderColor = &H808080
    End With
    Me.Label2.Visible = True
    OnMouseOver = False
End Sub

Private Sub TextBoxBody_Change()

End Sub

Private Sub TextBoxBody_Enter()
    ExTextBox_GotFocus2
End Sub

Private Sub TextBoxBody_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ExTextBox_LostFocus2
End Sub

Private Sub TextBoxBody_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver2
End Sub

Private Sub LabelTextBoxCaption_Click()
    ExTextBox_GotFocus2
End Sub

Private Sub LabelTextBoxCaption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOver2
End Sub

Private Sub ExTextBox_GotFocus2()
    If OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
    Me.LabelTextBoxCaption.Top = Me.LabelTextBoxFrame.Top - 8
    Me.LabelTextBoxCaption.Left = Me.LabelTextBoxFrame.Left + 6
    Me.LabelTextBoxCaption.FontSize = 8
    Me.LabelTextBoxCaption.ForeColor = &HFFFFC0
    Me.LabelTextBoxCaption.BackStyle = fmBackStyleTransparent
    Me.LabelTextBoxMask.Left = Me.LabelTextBoxCaption.Left
    With Me.LabelTextBoxCaption
        Me.LabelTextBoxMask.Width = MeasureTextSize(.Caption, .FontName, .FontSize).X * 1.3
    End With
'    Me.LabelTextBoxMask.Width = Me.LabelTextBoxCaption.Width
    Me.LabelTextBoxMask.Visible = True
    OnFocus = True
End Sub

Private Sub ExTextBox_LostFocus2()
    If Not OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 26
        .BorderColor = &H808080
        .ForeColor = &HC0C0C0
    End With
'    Me.LabelTextBoxCaption.Visible = True
    Me.LabelTextBoxCaption.Top = Me.LabelTextBoxFrame.Top + 3
    Me.LabelTextBoxCaption.Left = Me.LabelTextBoxFrame.Left + 1
    Me.LabelTextBoxCaption.FontSize = 10
    Me.LabelTextBoxCaption.ForeColor = &HC0C0C0
    Me.LabelTextBoxCaption.BackStyle = fmBackStyleTransparent
    Me.LabelTextBoxMask.Visible = False
    OnFocus = False
End Sub

Private Sub ExTextBox_MouseOver2()
    If OnMouseOver Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = Me.Label2.Caption
'        .Top = FrameTop - 6
'        FrameTop = .Top
'        .Height = 36
        .BorderColor = &HFFFFC0
        .ForeColor = &HFFFFC0
    End With
'    Me.LabelTextBoxCaption.Visible = False
    OnMouseOver = True
End Sub

Private Sub ExTextBox_MouseOut2()
    If Not OnMouseOver Then Exit Sub
    If OnFocus Then Exit Sub
    
    With Me.LabelTextBoxFrame
'        .Caption = ""
'        .Top = FrameTop + 6
'        FrameTop = .Top
'        .Height = 30
        .BorderColor = &H808080
        .ForeColor = &HFFFFC0
    End With
'    Me.LabelTextBoxCaption.Visible = True
    OnMouseOver = False
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    FrameTop = Me.Frame1.Top
    OnMouseOver = False
    Call Util.MakeTransparentFrame(Frame1)
    DatePicker.Init 'Me.TextBox1
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExTextBox_MouseOut2
End Sub

Private Sub UserForm_Terminate()
    Unload DatePicker
End Sub
