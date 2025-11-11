VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4608
   ClientLeft      =   48
   ClientTop       =   444
   ClientWidth     =   5688
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
    Dim i As Long
    Dim Item1 As String
    Dim Item2 As String
    Dim Item3 As String
    
    '-------------------
    Frame1.Height = 170
    Frame1.Width = 225
    '-------------------
    
    Set LView = New ListView
    If Not LView.Init(Frame1) Then Set LView = Nothing: Exit Sub
'    If Not LView.SetTextColor(RGB(0, 0, 0)) Then Set LView = Nothing: Exit Sub
    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
''    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]
'    SendMessage TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
'    If Not LView.SetTextBkColor(RGB(0, 0, 255)) Then Set LView = Nothing: Exit Sub
    
    'Hheader
    With LView
        .InsertColumn "Item", 0, 70
        .InsertColumn "subItem1", 1, 110
        .InsertColumn "subItem2", 2, 110
        
    'Item
        For i = 0 To 10
            Item1 = "アイテム" & i
            Item2 = "subItem1-" & i
            Item3 = "subItem2-" & i
            
            .InsertItem Item1, i
            .SetItem Item2, i, 1
            .SetItem Item3, i, 2
            'アイテムにチェックを入れる
           .SetCheckState i, 1
        Next
    End With
    
'    Dim myStr As String
'
'    ' Itemの追加...
'    For i = 0 To 32
'        myStr = Format$(i + 1, """Item"" 0")
'        LView.InsertItem myStr, i
'        'アイテムにチェックを入れる
'        LView.SetCheckState i, 1
'    Next
     
End Sub

Private Sub LView_ItemClick(ByVal iItem As Long, ByVal iSubItem As Long)
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

Private Sub LView_ItemSelected(ByVal iItem As Long, ByVal iSubItem As Long)
    TextBox1.Text = iItem
    TextBox2.Text = LView.LabelText(iItem, 2) 'iSubItem or 1 ~2
End Sub

Private Sub CommandButton1_Click()
    Dim i As Long
    Dim Item1 As String
    Dim Item2 As String
    Dim Item3 As String
    
    With LView
    'Item
        For i = 0 To 10
            Item1 = "Item" & i
            Item2 = "subItem1-" & i
            Item3 = "subItem2-" & i
            
            .InsertItem Item1, i
            .SetItem Item2, i, 1
            .SetItem Item3, i, 2
            'アイテムにチェックを入れる
           .SetCheckState i, 1
        Next
    End With
    
    
    
'    Dim myStr As String
'    For i = 0 To 19
'        myStr = Format$(i + 1, """Item"" 0")
'        LView.InsertItem myStr, i
'    Next
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
    Dim i As Long
    Dim cnt As Long
    cnt = LView.GetItemCount
    With LView
        For i = 0 To cnt
           If LView.GetCheckState(i) = False Then
'               Debug.Print LView.LabelText(i, 0)
           End If
        Next
    End With
End Sub

Private Sub CommandButton6_Click()
    With LView
        If CommandButton6.Caption = "Non Check" Then
            .SetCheckState -1, 0
            CommandButton6.Caption = "All Check"
        Else
            .SetCheckState -1, 1
            CommandButton6.Caption = "Non Check"
        End If
    End With
End Sub

Private Sub CommandButton7_Click()
    If Not LView.SetTextColor(RGB(255, 0, 0)) Then Exit Sub
    LView.Update 1
'    RedrawWindow TT.hChild, ByVal 0&, ByVal 0, 0
End Sub

Private Sub UserForm_Terminate()
    If Not LView Is Nothing Then
        Set LView = Nothing
    End If
End Sub
