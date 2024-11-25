VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtendedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'参照用変数
Private TargetForm_ As MSForms.UserForm
Private Minimize_   As Boolean
Private Maximize_   As Boolean
Private Resize_     As Boolean
Private hWnd_       As LongPtr
Private WndProc_    As LongPtr
Private PrevProc_   As LongPtr
Private Style_      As LongPtr
Private CloseButton_ As Boolean
Private MenuHandle_ As LongPtr
Private Menu_ As Boolean
Private module_ As Object

'最小化ボタンの有効/無効
Public Property Get Minimize() As Boolean
    Minimize = Minimize_
End Property

Public Property Let Minimize(ByVal flag As Boolean)
    Minimize_ = flag
End Property

'最大化ボタンの有効/無効
Public Property Get Maximize() As Boolean
    Maximize = Maximize_
End Property

Public Property Let Maximize(ByVal flag As Boolean)
    Maximize_ = flag
End Property

'リサイズの有効/無効
Public Property Get Resize() As Boolean
    Resize = Resize_
End Property

Public Property Let Resize(ByVal flag As Boolean)
    Resize_ = flag
End Property

'閉じるボタンの有効/無効
Public Property Get CloseButton() As Boolean
    CloseButton = CloseButton_
End Property

Public Property Let CloseButton(ByVal flag As Boolean)
    CloseButton_ = flag
End Property

'メニューバーの有効/無効
Public Property Get Menu() As LongPtr
    Menu = Menu_
End Property

Public Property Let Menu(ByVal flag As LongPtr)
    Menu_ = flag
End Property

'メニューバーのハンドル
Public Property Get MenuHandle() As LongPtr
    MenuHandle = MenuHandle_
End Property

Public Property Let MenuHandle(ByVal hMenu As LongPtr)
    MenuHandle_ = hMenu
End Property

'
Public Property Get WndProc() As LongPtr
    WndProc = WndProc_
End Property

Public Property Let WndProc(ByVal aWndProc As LongPtr)
    WndProc_ = aWndProc
End Property

'対象フォーム設定
Public Property Set Form(ByRef aTargetForm As MSForms.UserForm)
    Set TargetForm_ = aTargetForm
    Debug.Print "Form Set:" & MainView.hwnd
    hWnd_ = WinAPI.FindWindow("ThunderDFrame", TargetForm_.Caption)
    Debug.Print "Form Set:" & hWnd_
    Call Redraw
End Property

Public Property Get Form() As MSForms.UserForm
    Set Form = TargetForm_
End Property

'対象モジュール設定
Public Property Set Module(ByRef aTargetModule As Object)
    Set module_ = aTargetModule
End Property

Public Property Get Module() As Object
    Set Module = module_
End Property

'初期値設定
Private Sub Class_Initialize()
    'VBA UserForm の初期値
    PrevProc_ = 0
    Minimize_ = False
    Maximize_ = False
    Resize_ = False
    CloseButton_ = True
    MenuHandle_ = 0
    Menu_ = False
End Sub

'フォーム再描写
Public Sub Redraw()
    If Form Is Nothing Then Exit Sub
    
    Style_ = WinAPI.GetWindowLongPtr(hWnd_, GWL_STYLE)
    
    If Minimize Then Style_ = Style_ Or WS_MINIMIZEBOX
    If Maximize Then Style_ = Style_ Or WS_MAXIMIZEBOX
    If Resize Then Style_ = Style_ Or WS_THICKFRAME
    
    Call WinAPI.SetWindowLongPtr(hWnd_, GWL_STYLE, Style_)
    
    If Not CloseButton Then
        Dim hMenu_ As LongPtr
        hMenu_ = WinAPI.GetSystemMenu(hWnd_, 0&)
        Call WinAPI.DeleteMenu(hMenu_, SC_CLOSE, MF_BYCOMMAND)
    End If
    
    Call WinAPI.DrawMenuBar(hWnd_)

    If Menu Then
        If MenuHandle <> 0 Then Call WinAPI.SetMenu(hWnd_, MenuHandle)
    End If
End Sub

'Public Function SubClass(ByVal aPtrWndProc As LongPtr) As LongPtr
'    PrevProc_ = WinAPI.SetWindowLongPtr(hWnd_, GWL_WNDPROC, aPtrWndProc)
'    SubClass = PrevProc_
'End Function
'
'Private Sub UnSubClass()
'    Dim ret As LongPtr
'
'    If PrevProc_ <> 0 Then
'        ret = WinAPI.SetWindowLongPtr(hWnd_, GWL_WNDPROC, PrevProc_)
'        PrevProc_ = 0
'    End If
'End Sub

Public Function SubClass(ByVal aPtrWndProc As LongPtr) As LongPtr
    WndProc_ = aPtrWndProc
    SubClass = WinAPI.SetWindowSubclass(hWnd_, aPtrWndProc, hWnd_, CLngPtr(0&))
    Debug.Print "Subclass"
End Function

Private Function UnSubClass() As LongPtr
    UnSubClass = RemoveWindowSubclass(hWnd_, WndProc_, hWnd_)
    Debug.Print "UnSubclass"
End Function

Private Sub Class_Terminate()
    Call UnSubClass
End Sub
