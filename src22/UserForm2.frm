VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8820.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

Private Const ICC_DATE_CLASSES = &H100
Private Const DTS_SHORTDATEFORMAT = &H0     'YYYY/MM/DD
Private Const DTS_LONGDATEFORMAT = &H4      'YYYY年MM月DD日

Private Const GWL_HINSTANCE As Long = (-6)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_GROUP = &H20000

Private Const DTM_FIRST = &H1000
Private Const DTM_GETSYSTEMTIME = (DTM_FIRST + 1)   'ｺﾝﾄﾛｰﾙの日付/時刻を取得
Private Const DTM_SETSYSTEMTIME = (DTM_FIRST + 2)   'ｺﾝﾄﾛｰﾙの日付/時刻をｾｯﾄ
Private Const DTM_GETRANGE = (DTM_FIRST + 3)        'ｺﾝﾄﾛｰﾙの日付範囲を取得
Private Const DTM_SETRANGE = (DTM_FIRST + 4)        'ｺﾝﾄﾛｰﾙの日付範囲を設定

'ｺﾓﾝｺﾝﾄﾛｰﾙ初期化用構造体
Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

'ｼｽﾃﾑﾀｲﾑ構造体
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private hWndForm    As LongPtr
Private hWndClient  As LongPtr
Private hWndExcel   As LongPtr
Private hWndDate    As LongPtr
Private hInst       As LongPtr
Private tSysTime    As SYSTEMTIME

'------------------------------------------------------------------
'   UserForm起動時ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim tInitCmnCtrl As tagINITCOMMONCONTROLSEX
    Dim lRet As Long
    
    'UserFormのｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ取得
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    hWndClient = FindWindowEx(hWndForm, 0, vbNullString, vbNullString)

    'ｺﾓﾝｺﾝﾄﾛｰﾙ初期化
    With tInitCmnCtrl
        .dwICC = ICC_DATE_CLASSES
        .dwSize = Len(tInitCmnCtrl)
    End With
    lRet = InitCommonControlsEx(tInitCmnCtrl)
    
    'ｱﾌﾟﾘｹｰｼｮﾝ(Excel)のｲﾝｽﾀﾝｽﾊﾝﾄﾞﾙ取得
    hWndExcel = FindWindow("XLMAIN", Application.Caption)
    hInst = GetWindowLongPtr(hWndExcel, GWL_HINSTANCE)
    
    '日付選択ｺﾝﾄﾛｰﾙ作成
    hWndDate = CreateWindowEx(0, "SysDateTimePick32", vbNullString, _
                              WS_CHILD Or WS_VISIBLE Or DTS_SHORTDATEFORMAT Or WS_GROUP, _
                              PtToPx(Me.ComboBox1.Left), PtToPx(Me.ComboBox1.Top), _
                              PtToPx(Me.ComboBox1.Width), PtToPx(Me.ComboBox1.Height), _
                              hWndClient, 0, hInst, vbNullString)

End Sub
'------------------------------------------------------------------
'   UserForm終了時ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub UserForm_Terminate()
 
    '日付選択ｺﾝﾄﾛｰﾙを破棄
    Call DestroyWindow(hWndDate)
 
End Sub

'------------------------------------------------------------------
'   ｺﾏﾝﾄﾞﾎﾞﾀﾝ押下ｲﾍﾞﾝﾄ
'------------------------------------------------------------------
Private Sub CommandButton1_Click()

    Dim lRet As LongPtr
    Dim sMsg As String
    Dim sDayOfWeek As String
    
    '日付選択ｺﾝﾄﾛｰﾙの値取得
    lRet = SendMessage(hWndDate, DTM_GETSYSTEMTIME, 0, tSysTime)

    '取得した値から年/月/日/曜日を取得
    Select Case tSysTime.wDayOfWeek
        Case 0: sDayOfWeek = "日曜日"
        Case 1: sDayOfWeek = "月曜日"
        Case 2: sDayOfWeek = "火曜日"
        Case 3: sDayOfWeek = "水曜日"
        Case 4: sDayOfWeek = "木曜日"
        Case 5: sDayOfWeek = "金曜日"
        Case 6: sDayOfWeek = "土曜日"
    End Select

    sMsg = "年:　 " & tSysTime.wYear & vbLf & _
           "月:　 " & tSysTime.wMonth & vbLf & _
           "日:　 " & tSysTime.wDay & vbLf & _
           "曜日: " & sDayOfWeek
           
    '取得した日付を表示
    Call MsgBox(sMsg)
    
End Sub

'------------------------------------------------------------------
'   ﾎﾟｲﾝﾄ(pt)→ﾋﾟｸｾﾙ(px)変換
'------------------------------------------------------------------
Function PtToPx(ByVal dPt As Double) As Double

    PtToPx = dPt * 96 / 72
    
End Function

