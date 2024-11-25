VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxtControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents myForm As MSForms.UserForm
Attribute myForm.VB_VarHelpID = -1
Public WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1
Public WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents ErrLabel As MSForms.Label
Attribute ErrLabel.VB_VarHelpID = -1
Public txt As New Collection
Public TxtBox As New clsTxtControl
Public txtKontrol

Public Sub ControlTextBox(Form As MSForms.UserForm)
    Dim FontName As String
    FontName = "Poppins"
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
    
    For I = 1 To txt.Count
        If txt(I).Text = "" And txt(I).Tag <> "Search" Then
            Set ErrLabel = myForm.Add("Forms.Label.1", txt(I).Name & "ErrLabel")
            With ErrLabel
                .Caption = "*This fields is required"
                .Left = txt(I).Left + 2
                .Top = txt(I).Top + 24
                .Font.size = 8
                .Font.Name = txt(I).Font.Name
                .ForeColor = &HFF
                .BackStyle = fmBackStyleTransparent
                .MousePointer = fmMousePointerIBeam
                .Width = 155
                 txtKontrol = False
            End With
             Set TxtBox = New clsTxtControl
             Select Case TypeName(txt(I))
                Case Is = "TextBox"
                    Set TxtBox.TextBox = txt(I)
                Case Is = "ComboBox"
                    Set TxtBox.ComboBox = txt(I)
             End Select
             Set TxtBox.ErrLabel = ErrLabel
             Set TxtBox.myForm = Form
             txt.Add TxtBox
        Else
            Set ErrLabel = myForm.Add("Forms.Label.1", txt(I).Name & "ErrLabel")
            With ErrLabel
                .Caption = ""
                .BackStyle = fmBackStyleTransparent
            End With
        End If

        Set TxtBox = New clsTxtControl
        Select Case TypeName(txt(I))
               Case Is = "TextBox"
                   Set TxtBox.TextBox = txt(I)
               Case Is = "ComboBox"
                   Set TxtBox.ComboBox = txt(I)
        End Select
        Set TxtBox.myForm = Form
        txt.Add TxtBox
    Next I
End Sub

Private Sub Class_Terminate()
    Set txt = Nothing
    Set TxtBox = Nothing
End Sub

Private Sub ErrLabel_Click()
    Debug.Print TypeName(Me) & ":ErrLabel_Click"
    Focus = Mid(ErrLabel.Name, 1, Len(ErrLabel.Name) - 8)
    With myForm.Controls(Focus)
        .SetFocus
    End With
End Sub

Private Sub TextBox_Change()
    Debug.Print TypeName(Me) & ":TextBox_Change"
    Dim I As Integer
    For Each Ctrl In myForm.Controls
    
       If Ctrl.Name = TextBox.Name Then
       
            With myForm.Controls(TextBox.Name & "ErrLabel")
               .Caption = ""
            End With
        End If
    Next Ctrl
End Sub

Private Sub ComboBox_Change()
    Debug.Print TypeName(Me) & ":ComboBox_Change"
    Dim I As Integer
    For Each Ctrl In myForm.Controls
    
       If Ctrl.Name = ComboBox.Name Then
            With myForm.Controls(ComboBox.Name & "ErrLabel")
               .Caption = ""
            End With
        End If
    Next Ctrl
End Sub
