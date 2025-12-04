VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4608
   ClientLeft      =   48
   ClientTop       =   444
   ClientWidth     =   6240
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClassNameA Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc.dll" (ByVal inIAccessible As Office.IAccessible, ByRef phWnd As LongPtr) As Long

Private WithEvents LView As ListView
Attribute LView.VB_VarHelpID = -1
Private hWnd As LongPtr

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim Item1 As String
    Dim Item2 As String
    Dim Item3 As String
    
    hWnd = GetUserFormHwndAsThunderDFrame(Me)
    '-------------------
    Frame1.Height = 170
    Frame1.Width = 225
    '-------------------
    
    Set LView = New ListView
    If Not LView.Init(Frame1) Then Set LView = Nothing: Exit Sub
'    If Not LView.SetTextColor(RGB(0, 0, 0)) Then Set LView = Nothing: Exit Sub
'    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, ByVal RGB(0, 0, 0)    '[黒]
    
'    SendMessage TT.hChild, LVM_SETTEXTCOLOR, 0, ByVal RGB(0, 0, 0)    '[黒]
'    SendMessage hChild, LVM_SETTEXTCOLOR, 0, ByVal RGB(0, 0, 0)    '[黒]
'    LView.SetTextColor ByVal RGB(0, 255, 0)
''    SendMessageW TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 255)  '[青]
'    SendMessage TT.hChild, LVM_SETTEXTCOLOR, 0, RGB(0, 0, 0)    '[黒]
    
'    If Not LView.SetTextColor(RGB(0, 255, 0)) Then Set LView = Nothing: Exit Sub
'    If Not LView.SetTextBkColor(RGB(0, 0, 255)) Then Set LView = Nothing: Exit Sub
    
    Dim HeaderItems(2) As String
    Dim RecordItem(4) As String
    
    HeaderItems(0) = "Item"
    HeaderItems(1) = "subItem1"
    HeaderItems(2) = "subItem2"
    'Hheader
    With LView
'        .SetTextColor RGB(0, 255, 0)

'        .InsertColumn "Item", 0, 70
'        .InsertColumn "subItem1", 1, 110
'        .InsertColumn "subItem2", 2, 110
        .SetReportHeader HeaderItems
    'Item
        For i = 0 To 10
'            Item1 = "アイテム" & i
'            Item2 = "subItem1-" & i
'            Item3 = "subItem2-" & i
'
'            .InsertItem Item1, i
'            .SetItem Item2, i, 1
'            .SetItem Item3, i, 2
            
            RecordItem(0) = "アイテム" & i
            RecordItem(1) = "subItem1-" & i
            RecordItem(2) = "subItem2-" & i
            RecordItem(3) = "subItem3-" & i
            RecordItem(4) = "subItem4-" & i
            .SetReportRecord RecordItem
            'アイテムにチェックを入れる
           .SetCheckState i, 1
        Next
        .SetAllColumnWidthAuto
'        .SetColumnWidth 0, LVSCW_AUTOSIZE_USEHEADER
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
    Debug.Print "Item:" & iItem & " / SubItem:" & iSubItem & " Clicked"
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
    Debug.Print "Item:" & iItem & " / SubItem:" & iSubItem & " Selected"
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
    Debug.Print LView.LabelText(5, 0)
'    With LView
'        For i = 0 To cnt
'           If LView.GetCheckState(i) = False Then
'               Debug.Print LView.LabelText(i, 0)
'           End If
'        Next
'    End With
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
    Dim lColor As Long
    If Not ChooseColorDialog(lColor) Then Exit Sub
'    If Not LView.SetTextColor(RGB(255, 0, 0)) Then Exit Sub
    If Not LView.SetTextColor(lColor) Then Exit Sub
    LView.RedrawItems 0, LView.GetItemCount
'    LView.Update 1
'    RedrawWindow TT.hChild, ByVal 0&, ByVal 0, 0
End Sub

Private Sub CommandButton8_Click()
    Dim lf As LOGFONTA
    Dim cf As CHOOSE_FONTA

    lf.lfHeight = 13                                    'デフォルト高さ
    lf.lfWidth = 0                                      'デフォルト幅
    lf.lfEscapement = 0                                 'ベースラインと文字送りベクトルの間の角度
    lf.lfOrientation = 0                                'ベースラインとオリエンテーションベクトルの間の角度
    lf.lfWeight = FW_NORMAL                             '標準
    lf.lfCharSet = DEFAULT_CHARSET                      'デフォルトキャラクタセット
    lf.lfOutPrecision = OUT_DEFAULT_PRECIS              'デフォルト精度マッピング
    lf.lfClipPrecision = CLIP_DEFAULT_PRECIS            'デフォルトクリッピング精度
    lf.lfQuality = DEFAULT_QUALITY                      'デフォルト品質
    lf.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN     'デフォルトピッチ
'    lf.lfFaceName = CByte("Yu Gothic UI") ' & Chr$(0)            'デフォルト指定フォント名

'    Dim Style As String * 256

    cf.lStructSize = Len(cf)                            '構造体サイズ
    cf.hwndOwner = hWnd                              'ダイアログボックスのウィンドウ
    cf.hdc = 0                                          'デフォルトプリンタのデバイスコンテキスト
    cf.lpLogFont = VarPtr(lf)                                 'LOGFONTメモリーブロックバッファへのポインタ
    cf.iPointSize = 100                                 '10 ポイントフォント(1/10 point)
    cf.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    cf.rgbColors = &H0                                  '黒 RGB(0, 0, 0)
    cf.lpTemplateName = 0
'    cf.hInstance = GethInst                             'インスタンス
'    cf.lpszStyle = StrAdr(String$(256, Chr$(0)))
'    cf.lpszStyle = StrPtr(Style)
    cf.nFontType = REGULAR_FONTTYPE                     'レギュラーフォントタイプ
    cf.MISSING_ALIGNMENT = 0
    cf.nSizeMin = 10                                    '最小ポイントサイズ
    cf.nSizeMax = 72                                    '最大ポイントサイズ
    
    ChooseFontDialog cf
    
    Debug.Print lf.lfFaceName
End Sub

Private Sub UserForm_Terminate()
    If Not LView Is Nothing Then
        Set LView = Nothing
    End If
End Sub

Public Function GetUserFormHwndAsThunderDFrame(ByVal inUserForm As Office.IAccessible) As LongPtr               'HWND
'ユーザーフォームから ThunderDFrame としての HWND を取得する。
'https://qiita.com/nukie_53/items/39c28d000d521329548b
'inUserForm :VBA の UserForm を指定する。 UserForm 内から呼び出す場合は Me を指定すればよい。
'           :Dim h As LongPtr
'           :h = GetUserFormHwndAsThunderDFrame(Me)
'return     :inUserForm の ThunderDFrame としての HWND を返す。意図したクラス名を取得できなかった場合はエラー。
    Const ExpectClassName = "ThunderDFrame"
    If inUserForm Is Nothing Then Err.Raise 91
    
    'inUserForm から HWND を取得する。
    Dim hWnd1 As LongPtr
    Dim hr As Long 'HRESULT
    hr = WindowFromAccessibleObject(inUserForm, hWnd1)
    
    '成功した場合は S_OK 、失敗した場合はそれ以外の値となる。
    Const S_OK = 0
    If hr <> S_OK Then Err.Raise 5, , "ユーザーフォームのHWNDを取得できませんでした。"
    
    'HWND のクラス名を確認。
    
    '想定されるクラス名は長くても20文字程度。
    Dim classNameBuffer As String
    classNameBuffer = VBA.Strings.String(20, 0)
    
    'hWnd1 からクラス名を取得。
    Dim usedLen As Long
'    usedLen = GetClassNameW(hWnd1, VBA.[_HiddenModule].StrPtr(classNameBuffer), VBA.Strings.Len(classNameBuffer))
    usedLen = GetClassNameA(hWnd1, classNameBuffer, VBA.Strings.Len(classNameBuffer))
    
    Dim className1 As String
    className1 = VBA.Strings.Left$(classNameBuffer, usedLen)
    
    If className1 = ExpectClassName Then
        '意図したクラス名なのでここで終了。
        Let GetUserFormHwndAsThunderDFrame = hWnd1
        Exit Function
    End If
    
    If className1 Like "F3 Server *" Then
        '365 環境？だと、"F3 Server eb8e0000"のようなクラス名になる。
        'この場合は、親の HWND をたどれば ThunderDFrame を取得できる。
        'Next
    Else
        Err.Raise 5, , "意図したクラス名のHWNDを取得できませんでした。" & vbLf & "HWND : " & hWnd1 & vbLf & "ClassName : " & className1
    End If
    
    '親の HWND を取得する。
    Dim hWnd2 As LongPtr
    hWnd2 = GetParent(hWnd1)
    
    'クラス名を確認。
'    usedLen = GetClassNameW(hWnd2, VBA.[_HiddenModule].StrPtr(classNameBuffer), VBA.Strings.Len(classNameBuffer))
    usedLen = GetClassNameA(hWnd2, classNameBuffer, VBA.Strings.Len(classNameBuffer))
    
    Dim className2 As String
    className2 = VBA.Strings.Left$(classNameBuffer, usedLen)
    
    If className2 = ExpectClassName Then
        Let GetUserFormHwndAsThunderDFrame = hWnd2
        Exit Function
    End If
    
    Err.Raise 5, , "意図したクラス名のHWNDを取得できませんでした。" & vbLf & "HWND : " & hWnd2 & vbLf & "ClassName : " & className2
End Function

