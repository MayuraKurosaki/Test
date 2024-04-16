VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseEventForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright(C) 2000-2009 kmado. All Rights Reserved.
'ユーザーフォームに独自イベント kEventsForm Version 1.01
'クラスモジュール MouseEventForm
Option Explicit
'Private Declare Function WindowFromObject Lib "oleacc" Alias "WindowFromAccessibleObject" (ByVal pacc As Object, phwnd As Long) As Long
'Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAcessible As Object, ByRef hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long

'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'Private Declare Sub DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal fAccept As Long)
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As LongPtr, ByVal fAccept As Long)

'Private Declare Sub DragFinish Lib "shell32" (ByVal HDROP As Long)
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)

'Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Const GWL_WNDPROC = (-4)
'Private dwnd&, dhook&, duf As MSForms.UserForm
Private dhook As LongPtr
Private duf As MSForms.UserForm
Private dwnd As LongPtr

'MouseWheelイベント(UserFormにてマウスホイールのスクロールを取得する)
Public Event MouseWheel(ByVal Control As MSForms.Control, ByVal wParam As LongPtr, ByVal Shift As Long)
'DropFilesイベント(UserFormにドラッグ&ドロップされたファイルを取得する)
Public Event DropFiles(ByVal DropFile As String)

'Initializeメソッド　CEventFormクラスの開始
'注)CEventFormクラスのInitializeメソッドはUserFormを表示して実行の事
'　 UserFormのInitializeイベントでの記述は不可
Public Sub Initialize(ByVal uf As MSForms.UserForm)
    Dim cf As MouseEventForm
'  WindowFromObject uf, dwnd
'  WindowFromAccessibleObject uf, dwnd
    Call WindowFromAccessibleObject(uf, dwnd)
    Set duf = uf
    DragAcceptFiles dwnd, 1
    dhook = SetWindowLongPtr(dwnd, GWL_WNDPROC, AddressOf FormHook)
    If gparam Is Nothing Then Set gparam = New Collection
    gparam.Add Me, CStr(dwnd)
    Do While dwnd
        DoEvents
        Sleep 50
    Loop
End Sub

'Terminateメソッド　CEventFormクラスの終了
Public Sub Terminate()
    SetWindowLongPtr dwnd, GWL_WNDPROC, dhook
    If Not gparam Is Nothing Then
        gparam.Remove CStr(dwnd)
        If gparam.Count = 0 Then Set gparam = Nothing
    End If
    DragAcceptFiles dwnd, 0
    dhook = 0
    dwnd = 0
End Sub

Public Property Get Hook() As Long
    Hook = dhook
End Property

Public Function EventMouseWheel(ByVal Control As MSForms.Control, ByVal wParam As LongPtr, ByVal Shift As Long) As Boolean
    RaiseEvent MouseWheel(Control, wParam, Shift)
End Function

Public Function EventDropFiles(ByVal hwnd As Long, ByVal wParam As Long) As Boolean
    Dim i As Long, l As Long
    Dim s As String
    For i = 0 To DragQueryFile(wParam, -1, vbNullString, 0) - 1
        l = DragQueryFile(wParam, i, vbNullString, 0)
        If l > 0 Then
            s = String$(l, 0)
            If DragQueryFile(wParam, i, s, l + 1) Then RaiseEvent DropFiles(s)
        End If
    Next
    DragFinish wParam
End Function

'UserFormのアクティブコントロールを取得
Public Function Control(Optional ByVal uf As MSForms.UserForm) As MSForms.Control
    Dim ctrl As MSForms.Control
    If uf Is Nothing Then Set uf = duf
    Set ctrl = uf.ActiveControl
    Do
        Set Control = ctrl
        If TypeOf ctrl Is MSForms.Frame Then
            Set ctrl = ctrl.ActiveControl
        ElseIf TypeOf ctrl Is MSForms.MultiPage Then
            Set ctrl = ctrl.SelectedItem().ActiveControl
        Else
            Exit Do
        End If
    Loop
End Function

'MouseEventFormクラスの仕様
'機能
'ユーザーフォームにマウスホイールとファイルドロップのイベントを実装します

'モジュール構成(以下の2モジュールが必須)
'クラスモジュール[MouseEventForm]、標準モジュール[MouseEventModule]

'定義
'Private WithEvents expression As MouseEventForm
'expression 定義する変数の名前を指定します

'クラスのインスタンス作成
'Set expression = New MouseEventForm

'メソッド
'Initialize メソッド
'機能
'ユーザーフォームの独自イベントを開始します
'構文
'expression.Initialize ignore
'Sub Initialize(ByVal uf As MSForms.UserForm)
'引数 uf:UserFormオブジェクト

'Terminate メソッド
'機能
'ユーザーフォームの独自イベントを終了します
'構文
'expression.Terminate

'イベント
''注)イベントプロシージャでエラー停止するとExcelがハングアップします。必ずエラートラップを行うこと
'MouseWheel イベント
'機能
'UserFormのマウスホイールのスクロールを取得する
'構文
'Public Event MouseWheel(ByVal Control As MSForms.Control, ByVal wParam As Long, ByVal Shift As Long)
'引数 Control：UserFormのアクティブコントロール
'　　 wParam：正数=Up　負数=Down
'　　 Shift：1=Shiftキー, 2=Ctrlキー, 4=Altキー

'DropFiles イベント
'機能
'UserFormにドラッグ&ドロップされたファイルを取得する
'構文
'Public Event DropFiles(ByVal DropFile As String)
'引数 DropFile：ドロップされたフルファイル名

'使用例
''任意のUserFormモジュール
'Option Explicit
'Private WithEvents Form As MouseEventForm 'MouseEventFormクラスのオブジェクト変数宣言

''ユーザーフォームの独自イベントを開始
''注)CEventFormクラスのInitializeメソッドはUserFormを表示して実行の事
''   UserFormのInitializeイベントでの記述ではまだ表示されていないので不可
'Private Sub UserForm_Activate()
'  '他のコードがあれば先に記述のこと
'  If Form Is Nothing Then
'    Set Form = New MouseEventForm
'    Form.Initialize Me
'  End If
'End Sub

''ユーザーフォームの独自イベントを終了
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'  Form.Terminate
'  Set Form = Nothing
'End Sub
