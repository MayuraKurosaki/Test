VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MouseEventTestForm2 
   Caption         =   "MouseEventsForm2"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "MouseEventTestForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MouseEventTestForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents Form As MouseEventForm 'MouseEventFormクラスのオブジェクト変数宣言
Attribute Form.VB_VarHelpID = -1

Private Sub CommandButton1_Click()
    Unload Me
End Sub

'DropFilesイベントの記述例
'UserFormにドラッグ&ドロップされたファイルを取得するイベント
'引数 DropFile：ドロップされたフルファイル名
Private Sub Form_DropFiles(ByVal DropFile As String)
    Debug.Print "Form_DropFiles"
    On Error Resume Next
    Dim i As Long
    With ListBox1
        For i = 1 To .ListCount
            If DropFile = .List(i - 1, 2) Then Exit Sub
        Next
        .AddItem .ListCount + 1
        .List(.ListCount - 1, 1) = Dir(DropFile, vbReadOnly Or vbHidden Or vbSystem)
        .List(.ListCount - 1, 2) = DropFile
    End With
End Sub

'MouseWheelイベントの記述例
'UserFormにてマウスホイールのスクロールを取得するイベント
'引数 Control：UserFormのアクティブコントロール
'　　 wParam：正数=Up　負数=Down
'　　 Shift：1=Shiftキー, 2=Ctrlキー, 4=Altキー
Private Sub Form_MouseWheel(ByVal Control As MSForms.Control, ByVal wParam As Long, ByVal Shift As Long)
    Debug.Print "Form_MouseWheel"
    On Error Resume Next
    Dim scroll As Long
    Const MINS = 3, MAXS = MINS * 4
    Select Case TypeName(Control)
    Case "ListBox", "ComboBox"
        scroll = IIf(Shift, MAXS, MINS)
        With Control
            If TypeOf Control Is MSForms.ComboBox Then .DropDown
            If 0 < wParam Then
                .TopIndex = IIf(.TopIndex < scroll, 0, .TopIndex - scroll)
            Else
                .TopIndex = .TopIndex + scroll
            End If
        End With
    End Select
End Sub

Private Sub UserForm_Activate()
    'MouseEventFormクラスの開始
    If Form Is Nothing Then
        Set Form = New MouseEventForm
        Form.Initialize Me
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 1 To 100
        ComboBox1.AddItem i & ": " & "Combo_" & i * 100
    Next
    ComboBox1.ListIndex = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'MouseEventFormクラスの終了
    Form.Terminate
    Set Form = Nothing
End Sub
