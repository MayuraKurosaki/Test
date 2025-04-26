VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6490
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10040
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lx As Single, ly As Single

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_Click()
    Dim HTI As LVHITTESTINFO

    With HTI
        .pt.x = lx
        .pt.y = ly
        .Flags = LVHT_ONITEM
    End With

    Call SendMessage(ListView1.hwnd, LVM_SUBITEMHITTEST, 0, HTI)

    Dim lst As ListItem
    If (HTI.iItem > -1) Then
        Set lst = ListView1.ListItems(HTI.iItem + 1)

        If HTI.iSubItem = 3 Then
            lst.ListSubItems(HTI.iSubItem).Text = "osita"
        End If
        MsgBox "Clicked item " & HTI.iItem + 1 & " and SubItem " & HTI.iSubItem
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

End Sub

Private Sub ListView1_DblClick()

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub ListView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub ListView1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
   lx = x
   ly = y
End Sub

Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
'   lx = X
'   ly = Y
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub UserForm_Initialize()
    With ListView1

        '初期化
        .View = lvwReport           '外観表示指定
        .LabelEdit = lvwManual      '左端項目の編集設定
        .HideSelection = False      'フォーカス移動時の選択解除設定
        .AllowColumnReorder = True  '列幅の変更有無
        .FullRowSelect = True       '行全体を選択有無
        .Gridlines = True           'グリッド線表示有無

        '列見出し
        .ColumnHeaders.Clear
        .ColumnHeaders.Add
        .ColumnHeaders.Add , "_LN", "L/N", , lvwColumnRight
        .ColumnHeaders.Add , "_Target", "対象", , lvwColumnCenter
        .ColumnHeaders.Add , "_Target2", "対象2", , lvwColumnCenter
        .ColumnHeaders.Add , "_Command", "命令", , lvwColumnLeft
        .ColumnHeaders.Add , "_Parameter", "パラメータ", , lvwColumnLeft
        .ColumnHeaders.Add , "_Remark", "備考", , lvwColumnLeft

        Dim i As Long, j As Long
        'データの登録
        For i = 1 To 30
            With .ListItems.Add
                For j = 1 To 3
                    .SubItems(j) = i & "-" & j
                    .ListSubItems(j).ForeColor _
                        = vbBlack
                Next
            End With
        Next
    End With
    Call MakeTransparentFrame(Frame1)
'    Call MakeTransparentFrame(ListView1)
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub

Private Sub UserForm_Terminate()

End Sub
