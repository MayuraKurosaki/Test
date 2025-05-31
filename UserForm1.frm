VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4720
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   4620
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents LView As ListView
Attribute LView.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Dim I As Long
    Dim Item1 As String
    Dim Item2 As String
    Dim Item3 As String
    
    '-------------------
    Frame1.Height = 170
    Frame1.Width = 225
    '-------------------

    Set LView = New ListView
    LView.Init Frame1
    SendMessage TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
'    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]

    'Hheader
    With LView
        .InsertColumn "Item", 0, 70
        .InsertColumn "subItem1", 1, 110
        .InsertColumn "subItem2", 2, 110
    'Item
        For I = 0 To 10
            Item1 = "Item" & I
            Item2 = "subItem1-" & I
            Item3 = "subItem2-" & I
            
            .InsertItem Item1, I
            .SetItem Item2, I, 1
            .SetItem Item3, I, 2
            'アイテムにチェックを入れる
           .SetCheckState I, 1
        Next
    End With
 
End Sub

Private Sub LView_ItemClick(ByVal iItem&, ByVal iSubItem&)  '変数名&→Long
    TextBox1.Text = iItem
    TextBox2.Text = LView.LabelText(iItem, 2) 'iSubItem or 1 ~2

    If iItem <> -1 Then
        If LView.GetCheckState(iItem) Then
            LView.SetCheckState iItem, 1
        Else
            LView.SetCheckState iItem, 0
        End If
    End If
End Sub

Private Sub LView_ItemSelected(ByVal iItem&, ByVal iSubItem&)
    TextBox1.Text = iItem
    TextBox2.Text = LView.LabelText(iItem, 2) 'iSubItem or 1 ~2
End Sub

Private Sub CommandButton1_Click()
    Dim I As Long
    Dim Item1 As String
    Dim Item2 As String
    Dim Item3 As String
    
    With LView
    'Item
        For I = 0 To 10
            Item1 = "Item" & I
            Item2 = "subItem1-" & I
            Item3 = "subItem2-" & I
            
            .InsertItem Item1, I
            .SetItem Item2, I, 1
            .SetItem Item3, I, 2
            'アイテムにチェックを入れる
           .SetCheckState I, 1
        Next
    End With

End Sub

Private Sub CommandButton2_Click()
    LView.DeleteAllItems
End Sub

Private Sub CommandButton3_Click()
    MsgBox LView.GetItemCount
End Sub
Private Sub CommandButton4_Click()
    UserForm1.Hide
    Unload Me
End Sub

Private Sub CommandButton5_Click()
    Dim I As Long
    Dim cnt As Long
    cnt = LView.GetItemCount
With LView
    For I = 0 To cnt
       If LView.GetCheckState(I) = False Then
           Debug.Print LView.LabelText(I, 0)
       End If
    Next
End With
  
End Sub




