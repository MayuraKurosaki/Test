VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTPickerOnCombo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
'   ComboBox上にDateTimePickerを表示させるクラス(Ver3.3)
'---------------------------------------------------------------------
'   井上[Excelでお仕事] http://www.asahi-net.or.jp/~ef2o-inue/
'
'   角田[AddinBox] http://addinbox.sakura.ne.jp/index.htm
'                  http://addinbox.sakura.ne.jp/Excel_Tips06_3B.htm
'                  (旧サイト) http://www.h3.dion.ne.jp/~sakatsu/index.htm
'*******************************************************************************
'2003/07/22(2.03) 新規作成(by 井上)
'2003/07/23(2.03) 親ｺﾝﾄﾛｰﾙをTextBoxからComboBoxに変更(形が似ているから)(by 井上)
'2003/07/24(2.03) DateTimePickerの元ｺﾝﾄﾛｰﾙがｺﾝﾃﾅ上にあるｹｰｽの対応を追加(by 井上)
'-------------------------------------------------------------------------------
'2003/08/10
'         井上さん作のクラスモジュールの仕組みを AddinBox で解説
'  【以下、角田による改良】
'2004/04/27 [Ver1.0]
'         配色用プロパティの追加(バージョン表記を AddinBox版Ver1.0 と改める)
'2004/05/02 [Ver1.1]
'         配色用プロパティでシステムカラー利用可
'
'2004/06/05 [Ver2.0]
'         入力可能範囲(MinDate/MaxDate)プロパティを追加
'         Enabledプロパティを追加
'         クラス内で配列化する事により[modDTPickerOnComboBox]を廃止
'         クラス名を変更[clsDTPickerOnCombo2]
'
'2004/11/25 [Ver3.0]
'         土台のComboBoxからフォントを引き継ぐ機能を追加
'         Createメソッドに引数を追加
'2004/11/26 [Ver3.1]
'         DefaultFONT引数の内容を[True/False]から[True/Fale/2 or 3]に
'         変更（2:ComboBox部のみ変更 , 3:カレンダー部のみ変更）
'2005/3/10  [Ver3.2]
'         クラスやコレクション等の宣言時に付けていた[New]を外し、
'         Initializeにて[Set ～ = New ～]とする
'         [New] を付けているとTerminateでNothing代入しても再割当されてしまう為
'2011/9/27  [Ver3.3]
'         [x64]対応でAPI周りの処理を修正
'
'*******************************************************************************

'------------------ ＡＰＩ宣言 -----------------------
' コントロール初期化の指示子 (dwSize,dwICC は DWORD なのでx64でも Long でOK)
Private Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type
' コントロールから引き渡される日付時刻構造体
Private Type SYSTEMTIME
    Year            As Integer
    Month           As Integer
    DayOfWeek       As Integer
    Day             As Integer
    Hour            As Integer
    Minute          As Integer
    Second          As Integer
    Milliseconds    As Integer
End Type
' DTM_GETRANGE DTM_SETRANGE用の宣言
Private Type SYSTEMTIMERANGE
    MinYear         As Integer
    MinMonth        As Integer
    MinDayOfWeek    As Integer
    MinDay          As Integer
    MinHour         As Integer
    MinMinute       As Integer
    MinSecond       As Integer
    MinMilliseconds As Integer
    MaxYear         As Integer
    MaxMonth        As Integer
    MaxDayOfWeek    As Integer
    MaxDay          As Integer
    MaxHour         As Integer
    MaxMinute       As Integer
    MaxSecond       As Integer
    MaxMilliseconds As Integer
End Type

Private Const DATETIMEPICK_CLASS = "SysDateTimePick32"
Private Const ICC_DATE_CLASSES = &H100
Private Const DTS_SHORTDATEFORMAT = &H0     'コントロールパネル[日付：短い形式]
Private Const DTS_LONGDATEFORMAT = &H4      'コントロールパネル[日付：長い形式]

Private Const GDT_VALID = 0
Private Const GDTR_MIN = &H1
Private Const GDTR_MAX = &H2

Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Const DTM_FIRST = &H1000
Private Const DTM_GETSYSTEMTIME = (DTM_FIRST + 1)   ' コントロールの日付/時刻を取得
Private Const DTM_SETSYSTEMTIME = (DTM_FIRST + 2)   ' コントロールの日付/時刻をｾｯﾄ
Private Const DTM_GETRANGE = (DTM_FIRST + 3)        ' コントロールの日付範囲を取得
Private Const DTM_SETRANGE = (DTM_FIRST + 4)        ' コントロールの日付範囲を設定
Private Const DTM_SETFORMAT = (DTM_FIRST + 5)       ' コントロールに整形文字列をｾｯﾄ
Private Const DTM_SETMCCOLOR = (DTM_FIRST + 6)      ' ドロップカレンダーの色を取得
Private Const DTM_GETMCCOLOR = (DTM_FIRST + 7)      ' ドロップカレンダーの色を設定
Private Const DTM_GETMONTHCAL = (DTM_FIRST + 8)     ' ドロップカレンダーのウィンドウハンドルを取得
Private Const DTM_SETMCFONT = (DTM_FIRST + 9)       ' DTPコントロールのフォントをセット
Private Const DTM_GETMCFONT = (DTM_FIRST + 10)      ' DTPコントロールのフォントを取得

Private Const MCSC_BACKGROUND = 0                   ' 月間背景色を指定
Private Const MCSC_TEXT = 1                         ' 選択日付の色
Private Const MCSC_TITLEBK = 2                      ' タイトルバーの背景色
Private Const MCSC_TITLETEXT = 3                    ' タイトルバーのテキスト色
Private Const MCSC_MONTHBK = 4                      ' カレンダーの背景色
Private Const MCSC_TRAILINGTEXT = 5                 ' カレンダーの非選択日付のテキスト色

Private Const POINT_PER_INCH = 72
Private Const LOGPIXELSX = 88                       ' ポイント→ピクセル変換指定(横)
Private Const LOGPIXELSY = 90                       ' ポイント→ピクセル変換指定(縦)
Private Const GWL_HINSTANCE As Long = (-6)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

'CreateFont用
Private Const FW_DONTCARE = 0
Private Const FW_BOLD = 700
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SHIFTJIS_CHARSET = 128
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0

'GetStockObject
Private Const DEFAULT_GUI_FONT = 17     ' ユーザーインターフェイス用のデフォルトフォント

'---------------------------------------------------------------

' コモンコントロールの初期化API
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" _
        (LPINITCOMMONCONTROLSEX As Any) As Long
#Else
    Private Declare Function InitCommonControlsEx Lib "ComCtl32" _
        (LPINITCOMMONCONTROLSEX As Any) As Long
#End If

' ウィンドウハンドル取得API
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long
#End If

' 拡張ウィンドウハンドル取得API
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hwndParent As LongPtr, _
         ByVal hwndChildAfter As LongPtr, _
         ByVal lpszClass As String, _
         ByVal lpszWindow As String) As LongPtr
#Else
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hwndParent As Long, _
         ByVal hwndChildAfter As Long, _
         ByVal lpszClass As String, _
         ByVal lpszWindow As String) As Long
#End If

' ポイント→ピクセル変換係数取得API
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
        (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDeviceCaps Lib "gdi32" _
        (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

' アプリケーションのインスタンスを取得
' [x86]ではGetWindowLong , [x64]ではGetWindowLongPtr を使用する
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
        (ByVal hwnd As LongPtr, _
         ByVal nIndex As Long) As LongPtr

#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, _
         ByVal nIndex As Long) As Long
#End If

' DeskTopWindow取得
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
#End If

' デバイスコンテキスト取得
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDC Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
#Else
    Private Declare Function GetDC Lib "user32" _
        (ByVal hwnd As Long) As Long
#End If

' デバイスコンテキスト解放
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ReleaseDC Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else
    Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hwnd As Long, ByVal hdc As Long) As Long
#End If

' ウィンドウの作成(今回はDTPickerを作成)
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
        (ByVal dwExStyle As Long, _
         ByVal lpClassName As String, _
         ByVal lpWindowName As String, _
         ByVal dwStyle As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal hwndParent As LongPtr, _
         ByVal hMenu As LongPtr, _
         ByVal hInstance As LongPtr, _
         lpParam As Any) As LongPtr
#Else
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
        (ByVal dwExStyle As Long, _
         ByVal lpClassName As String, _
         ByVal lpWindowName As String, _
         ByVal dwStyle As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal hwndParent As Long, _
         ByVal hMenu As Long, _
         ByVal hInstance As Long, _
         lpParam As Any) As Long
#End If

' ウィンドウコントロールの操作
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As LongPtr, _
         ByVal wMsg As Long, _
         ByVal wParam As LongPtr, _
         lParam As Any) As LongPtr
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long
#End If

' ウィンドウの移動､拡縮(現在､未使用)
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function MoveWindow Lib "user32" _
        (ByVal hwnd As LongPtr, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long
#Else
    Private Declare Function MoveWindow Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long
#End If

' ウィンドウの撤去
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function DestroyWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function DestroyWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
#End If

' ウィンドウが作成されているかの判定
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function IsWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function IsWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
#End If

' Tabキー等でのフォーカス操作
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function SetFocus Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
#Else
    Private Declare Function SetFocus Lib "user32" _
        (ByVal hwnd As Long) As Long
#End If

' コントロールのEnabled(True/False)の設定
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function EnableWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal fEnable As Long) As Long
#Else
    Private Declare Function EnableWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal fEnable As Long) As Long
#End If

' コントロールのEnabled(True/False)状態の取得
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function IsWindowEnabled Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function IsWindowEnabled Lib "user32" _
        (ByVal hwnd As Long) As Long
#End If

' システムカラー(vbButtonFace等)のRGB値への変換
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long
#End If

' フォントの作成
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" _
        (ByVal nHeight As Long, _
         ByVal nWidth As Long, _
         ByVal nEscapement As Long, _
         ByVal nOrientation As Long, _
         ByVal nWeight As Long, _
         ByVal fdwItalic As Long, _
         ByVal fdwUnderline As Long, _
         ByVal fdwStrikeOut As Long, _
         ByVal fdwCharSet As Long, _
         ByVal fdwOutPrecision As Long, _
         ByVal fdwClipPrecision As Long, _
         ByVal fdwQuality As Long, _
         ByVal fdwPitchAndFamily As Long, _
         ByVal pszFaceName As String) As LongPtr
#Else
    Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
        (ByVal nHeight As Long, _
         ByVal nWidth As Long, _
         ByVal nEscapement As Long, _
         ByVal nOrientation As Long, _
         ByVal nWeight As Long, _
         ByVal fdwItalic As Long, _
         ByVal fdwUnderline As Long, _
         ByVal fdwStrikeOut As Long, _
         ByVal fdwCharSet As Long, _
         ByVal fdwOutPrecision As Long, _
         ByVal fdwClipPrecision As Long, _
         ByVal fdwQuality As Long, _
         ByVal fdwPitchAndFamily As Long, _
         ByVal pszFaceName As String) As Long
#End If

'フォントオブイェクトのリリース用
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" _
        (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function DeleteObject Lib "gdi32" _
        (ByVal hObject As Long) As Long
#End If

'既定のフォントなどの取得関数
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetStockObject Lib "gdi32" _
        (ByVal nIndex As Long) As LongPtr
#Else
    Private Declare Function GetStockObject Lib "gdi32" _
        (ByVal nIndex As Long) As Long
#End If

'---------- クラスモジュール内で使う変数の宣言(実行中固定) ----------------------
Private lngPixelsX As Long                  ' Point→Pixel係数(横)
Private lngPixelsY As Long                  ' Point→Pixel係数(縦)

Private strThunder As String
Private blnCreate As Boolean                ' Create済(True)/未(False)
Private MyForm As MSForms.UserForm          ' 呼出元UserForm
Private colBaseCtrl As Collection           ' 土台のコントロール(ComboBox)

#If VBA7 And Win64 Then
    Private lnghWnd_Excel As LongPtr            ' Excelのウィンドウハンドル
    Private lnghInstance As LongPtr             '   〃 のインスタンス
    Private lnghWnd_Form As LongPtr             ' UserFormのウィンドハンドル
    Private lnghWnd_Sub As LongPtr              ' 透明ウィンドウのウィンドウハンドル
    Private lnghwndDTP() As LongPtr             ' DTPickerのウィンドウハンドル
    Private lnghwndFONT() As LongPtr            ' フォントデータのウィンドウハンドル
    Private lnghwndDefaultFONT As LongPtr       ' 既定フォントのウィンドウハンドル
#Else
    Private lnghWnd_Excel As Long               ' Excelのウィンドウハンドル
    Private lnghInstance As Long                '   〃 のインスタンス
    Private lnghWnd_Form As Long                ' UserFormのウィンドハンドル
    Private lnghWnd_Sub As Long                 ' 透明ウィンドウのウィンドウハンドル
    Private lnghwndDTP() As Long                ' DTPickerのウィンドウハンドル
    Private lnghwndFONT() As Long               ' フォントデータのウィンドウハンドル
    Private lnghwndDefaultFONT As Long          ' 既定フォントのウィンドウハンドル
#End If

Private Const cstMask As Long = &H7FFFFFFF  'SystemColor定数の先頭1bitをoffにするマスク


'*******************************************************************************
'   イベントの記述
'*******************************************************************************
' ClassのInitialize
Private Sub Class_Initialize()
    blnCreate = False
    Set colBaseCtrl = New Collection
    
    'Point→Pixel変換係数算出
    Call GetLogPixelsXY
    
    If (Val(Application.Version) <= 8) Then
        strThunder = "ThunderXFrame"        ' Excel97
    Else
        strThunder = "ThunderDFrame"        ' Excel2000～
    End If

    '既定フォントのハンドル取得
    lnghwndDefaultFONT = GetStockObject(DEFAULT_GUI_FONT)
End Sub

' ClassのTerminate
Private Sub Class_Terminate()
    '後始末忘れ対策
    Me.Destroy
End Sub

'*******************************************************************************
'   メソッドの記述
'*******************************************************************************
' Add(基盤となるコントロールの登録)
Public Sub Add(NewCtrl As MSForms.ComboBox)
    If (blnCreate = False) Then
        colBaseCtrl.Add NewCtrl
    Else
        'Create済
    End If
End Sub

' Create(DTPickerコントロールの生成)
Public Sub Create(ByVal NewForm As MSForms.UserForm, _
                  Optional ByVal NewFormat As String = "", _
                  Optional ByVal NewMinDate As Date = #1/1/1601#, _
                  Optional ByVal NewMaxDate As Date = #12/31/9999#, _
                  Optional ByVal ForeColor As Long = -1, _
                  Optional ByVal BackColor As Long = -1, _
                  Optional ByVal TitleFore As Long = -1, _
                  Optional ByVal TitleBack As Long = -1, _
                  Optional ByVal Trailing As Long = -1, _
                  Optional ByVal DefaultFONT As Variant = True)
Dim i As Integer
Dim intFONT As Integer
    If (VarType(DefaultFONT) = vbBoolean) Then
        If (DefaultFONT = True) Then
            intFONT = 1
        Else
            intFONT = 0
        End If
    Else
        Select Case DefaultFONT
          Case 0 To 3
            intFONT = DefaultFONT
          Case Else
            intFONT = 1     '既定(True)
        End Select
    End If

    If (blnCreate = True) Then
        'Create済
    ElseIf (colBaseCtrl.Count < 1) Or (NewForm Is Nothing) Then
        'UserForm未指定/土台Ctrl未指定
        blnCreate = False
    Else
        Set MyForm = NewForm
        'UserFormのHWNDの取得
        lnghWnd_Form = FindWindow(strThunder, MyForm.Caption)
        If (lnghWnd_Form = 0) Then
            blnCreate = False
            Exit Sub
        End If
    
        lnghWnd_Excel = FindWindow("XLMAIN", Application.Caption)
        
        #If VBA7 And Win64 Then
            lnghInstance = GetWindowLongPtr(lnghWnd_Excel, GWL_HINSTANCE)
        #Else
            lnghInstance = GetWindowLong(lnghWnd_Excel, GWL_HINSTANCE)
        #End If
        
        'UserForm上にある透明WindowのHWNDの取得
        lnghWnd_Sub = FindWindowEx(lnghWnd_Form, GDT_VALID, vbNullString, vbNullString)
        
        'DTPickerの生成
        ReDim lnghwndDTP(1 To colBaseCtrl.Count)
        ReDim lnghwndFONT(1 To colBaseCtrl.Count)   'FONT変更なしでも終了処理用に配列は用意する
        For i = 1 To colBaseCtrl.Count
            Call Create_Sub(i, NewFormat, NewMinDate, NewMaxDate, intFONT)
        Next i
        blnCreate = True
        
        ' 日付の表示形式の初期設定
        If (NewFormat <> "") Then
            Me.DateFormat(0) = NewFormat
        End If
    
        '日付入力範囲の初期設定
        If (NewMinDate <= NewMaxDate) Then
            Me.MinDate(0) = NewMinDate
            Me.MaxDate(0) = NewMaxDate
        Else
            Me.MinDate(0) = #1/1/1601#
            Me.MaxDate(0) = #12/31/9999#
        End If

        'カレンダー配色の初期設定
        Me.CalendarForeColor(0) = ForeColor
        Me.CalendarBackColor(0) = BackColor
        Me.CalendarTitleForeColor(0) = TitleFore
        Me.CalendarTitleBackColor(0) = TitleBack
        Me.CalendarTrailingForeColor(0) = Trailing
    End If
End Sub

' Destroy(Window/コレクションの破棄)
Public Sub Destroy()
Dim i As Integer
    'Windowの破棄
    If (blnCreate = False) Then
        '不要
    Else
      If (colBaseCtrl.Count > 0) Then
        For i = 1 To colBaseCtrl.Count
            If (IsWindow(lnghwndDTP(i)) <> 0) Then
                Call DestroyWindow(lnghwndDTP(i))
            End If
            If (lnghwndFONT(i) <> 0) Then   'FONT変更なしでも配列は確保してある
                Call DeleteObject(lnghwndFONT(i))
            End If
        Next i
      End If
    End If

    Set MyForm = Nothing
    Set colBaseCtrl = Nothing
    Erase lnghwndDTP
    Erase lnghwndFONT
    blnCreate = False
End Sub

'(DTPicker コントロールの生成サブ)
Private Sub Create_Sub(ByVal Index As Integer, _
                       ByVal NewFormat As String, _
                       ByVal NewMinDate As Date, _
                       ByVal NewMaxDate As Date, _
                       ByVal NewFONT As Integer)
Dim icce As tagINITCOMMONCONTROLSEX
Dim dblTop As Double
Dim dblLeft As Double
Dim dblHeight As Double
Dim dblWidth As Double
Dim dblFontHeight As Double
Dim lngWeight As Long
Dim lngResult As Long

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    '土台のComboBoxからDTPickerコントロールのサイズを算出
    dblLeft = colBaseCtrl(Index).Left * lngPixelsX / POINT_PER_INCH
    dblTop = colBaseCtrl(Index).Top * lngPixelsY / POINT_PER_INCH
    dblWidth = colBaseCtrl(Index).Width * lngPixelsX / POINT_PER_INCH
    dblHeight = colBaseCtrl(Index).Height * lngPixelsY / POINT_PER_INCH

    '土台のComboBoxからフォントサイズ(高さ)を算出
    dblFontHeight = colBaseCtrl(Index).Font.Size * lngPixelsY / POINT_PER_INCH
    
    'DTPickerを[Enabled=False]にすると、ドロップボタン
    'クリックにより、土台のComboBoxリストが出てしまう為、
    'ComboBoxを最初から[Enabled=False]にする
    colBaseCtrl(Index).Enabled = False

    ' 既にウィンドウが存在する場合はウィンドウの破棄
    If (IsWindow(lnghwndDTP(Index)) <> 0) Then
        Call DestroyWindow(lnghwndDTP(Index))
    End If
    
    ' INITCOMMONCONTROLSEX構造体に値を代入
    With icce
        .dwICC = ICC_DATE_CLASSES
        .dwSize = Len(icce)
    End With
    
    ' コモンコントロールを初期化
    lngResult = InitCommonControlsEx(icce)

    ' ComboBoxコントロール上にDateTimePickerを作成(DTS_SHORTDATEFORMAT)
    lnghwndDTP(Index) = CreateWindowEx _
                            (GDT_VALID, DATETIMEPICK_CLASS, vbNullString, _
                             WS_CHILD Or WS_VISIBLE Or DTS_SHORTDATEFORMAT, _
                             dblLeft, dblTop, dblWidth, dblHeight, _
                             lnghWnd_Sub, GDT_VALID, lnghInstance, vbNullString)    ' 日付：短い形式

    If (NewFONT = 1) Then
        '既定フォント指定(True)なので、そのまま
    Else
        '土台のComboBoxからフォントを引き継ぐ
        If (colBaseCtrl(Index).Font.Bold = True) Then
            lngWeight = FW_BOLD
        Else
            lngWeight = FW_DONTCARE
        End If
        'フォントオブジェクトの生成(Heightはマイナス指定で指定値そのままの定義になる）
        lnghwndFONT(Index) = _
            CreateFont(CLng(dblFontHeight * (-1)), _
                       GDT_VALID, GDT_VALID, GDT_VALID, _
                       lngWeight, _
                       GDT_VALID, GDT_VALID, GDT_VALID, _
                       SHIFTJIS_CHARSET, _
                       OUT_DEFAULT_PRECIS, _
                       CLIP_DEFAULT_PRECIS, _
                       DEFAULT_QUALITY, _
                       (DEFAULT_PITCH + FF_DONTCARE), _
                       colBaseCtrl(Index).Font.Name)
        'Initialize内なので､SendMessageの4opは0(False:ReDrawなし)で良い
        Select Case NewFONT
          Case 0    '[False] ComboBox部＆カレンダー部とも変更
            'DTPicker(ComboBox部)のフォント設定
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      WM_SETFONT, lnghwndFONT(Index), ByVal GDT_VALID)
            '[WM_SETFONT]を送ると[DTM_SETMCFONT]も一緒に反映される
          Case 2    'ComboBox部のみ
            'DTPicker(ComboBox部)のフォント設定
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      WM_SETFONT, lnghwndFONT(Index), ByVal GDT_VALID)
            '[WM_SETFONT]を送ると[DTM_SETMCFONT]も一緒に反映されるので
            '既定フォントでカレンダー部を再設定する
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_SETMCFONT, lnghwndDefaultFONT, ByVal GDT_VALID)
          Case 3    'カレンダー部のみ
            'DTPicker(カレンダー部)のフォント設定
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_SETMCFONT, lnghwndFONT(Index), ByVal GDT_VALID)
        End Select
    End If
End Sub

' 画面精細度のPoint→Pixel変換係数算出
Private Sub GetLogPixelsXY()
#If VBA7 And Win64 Then
    Dim lnghwnd As LongPtr
    Dim lngDC As LongPtr
#Else
    Dim lnghwnd As Long
    Dim lngDC As Long
#End If

    'GetDCを使った後は、必ずReleaseDCで元に戻す。
    '画面表示用の専用メモリ(リソース)が消費され、
    '最後はWindowsがハングアップする危険があります。
    lnghwnd = GetDesktopWindow()
    lngDC = GetDC(lnghwnd)
    lngPixelsX = GetDeviceCaps(lngDC, LOGPIXELSX)
    lngPixelsY = GetDeviceCaps(lngDC, LOGPIXELSY)
    ReleaseDC lnghwnd, lngDC
End Sub

'*******************************************************************************
'   プロパティの記述
'*******************************************************************************
' Count(登録コントロール数の取得)
Public Property Get Count() As Integer
    Count = colBaseCtrl.Count
End Property

'--------------------------------------------------------------
' DateFormat(日付編集形式の設定)
Public Property Let DateFormat(Optional ByVal Index As Integer = 0, _
                               ByVal NewFormat As String)
Dim i As Integer

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    If (blnCreate = True) And (NewFormat <> "") Then
        Select Case Index
          Case 0        '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), _
                                          DTM_SETFORMAT, GDT_VALID, ByVal NewFormat)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_SETFORMAT, GDT_VALID, ByVal NewFormat)
          Case Else
        End Select
    Else
    End If
End Property

'--------------------------------------------------------------
' Value(シリアル値の設定と取得)
Public Property Get Value(Optional ByVal Index As Integer = 0) As Date
Dim st As SYSTEMTIME

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    If (blnCreate = True) Then
        Select Case Index
          Case 0        '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETSYSTEMTIME, GDT_VALID, st)
            Value = DateSerial(st.Year, st.Month, st.Day)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETSYSTEMTIME, GDT_VALID, st)
            Value = DateSerial(st.Year, st.Month, st.Day)
          Case Else
            Value = CDate(0)
        End Select
    Else
        Value = CDate(0)
    End If
End Property

Public Property Let Value(Optional ByVal Index As Integer = 0, _
                          ByVal NewValue As Date)
Dim st As SYSTEMTIME
Dim i As Integer

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    If (blnCreate = True) Then
        ' SYSTEMTIME構造体に引数の日付時間を分解して代入。
        With st
            .Year = Year(NewValue)
            .Month = Month(NewValue)
            .Day = Day(NewValue)
            .DayOfWeek = Weekday(NewValue) - 1  ' APIでは日曜日が0
            .Hour = 0
            .Minute = 0
            .Second = 0
        End With

        Select Case Index
          Case 0        '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
              lngSMResult = SendMessage(lnghwndDTP(i), _
                                        DTM_SETSYSTEMTIME, GDT_VALID, st)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_SETSYSTEMTIME, GDT_VALID, st)
          Case Else
        End Select
    Else
    End If
End Property

'-------------------------------------------------------------------------------
' Enabled(入力可否)
Public Property Get Enabled(Optional ByVal Index As Integer = 0) As Boolean
    If (blnCreate = True) Then
        Select Case Index
          Case 0        '[省略/0]→『１番』
            Enabled = CBool(IsWindowEnabled(lnghwndDTP(1)))
          Case 1 To colBaseCtrl.Count
            Enabled = CBool(IsWindowEnabled(lnghwndDTP(Index)))
          Case Else
            Enabled = False
        End Select
    Else
        Enabled = False
    End If
End Property

Public Property Let Enabled(Optional ByVal Index As Integer = 0, _
                            ByVal blnNewEnabled As Boolean)
Dim lngResult As Long
Dim i As Integer

    If (blnCreate = True) Then
        Select Case Index
          Case 0        '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngResult = EnableWindow(lnghwndDTP(i), blnNewEnabled)
            Next i
          Case 1 To colBaseCtrl.Count
            lngResult = EnableWindow(lnghwndDTP(Index), blnNewEnabled)
          Case Else
        End Select
    Else
    End If
End Property

'-------------------------------------------------------------------------------
' MinDate(選択可能日付範囲の下限)
Public Property Get MinDate(Optional ByVal Index As Integer = 0) As Date
Dim st As SYSTEMTIMERANGE

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0        '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETRANGE, GDT_VALID, st)
            MinDate = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETRANGE, GDT_VALID, st)
            MinDate = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
          Case Else
            MinDate = CDate(0)
        End Select
    Else
      MinDate = CDate(0)
    End If
End Property

Public Property Let MinDate(Optional ByVal Index As Integer = 0, _
                            ByVal dtmNewValue As Date)
Dim i As Integer
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                Call MinDate_Sub(i, dtmNewValue)
            Next i
          Case 1 To colBaseCtrl.Count
            Call MinDate_Sub(Index, dtmNewValue)
          Case Else
        End Select
    Else
    End If
End Property

Private Sub MinDate_Sub(ByVal Index As Integer, _
                        ByVal dtmNewValue As Date)
Dim dtmMax As Date
Dim st As SYSTEMTIMERANGE

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    ' 現在設定されている内容を取得する
    lngSMResult = SendMessage(lnghwndDTP(Index), _
                              DTM_GETRANGE, GDT_VALID, st)
    
    ' 取得したSYSTEMTIMERANGEの[Min]を引数の日付に置換
    With st
        .MinYear = Year(dtmNewValue)
        .MinMonth = Month(dtmNewValue)
        .MinDay = Day(dtmNewValue)
        .MinDayOfWeek = Weekday(dtmNewValue) - 1     ' APIでは日曜日が0
        .MinHour = 0
        .MinMinute = 0
        .MinSecond = 0
    End With
    
    If (st.MaxYear = 0) Then
        ' 上限が設定されていないので下限のみ設定する
        lngSMResult = SendMessage(lnghwndDTP(Index), _
                                  DTM_SETRANGE, GDTR_MIN, st)
    Else
        ' 既に設定されている上限とともに設定
        dtmMax = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
        If (dtmMax < dtmNewValue) Then
            '上限が新しい下限値より過去になるので、下限値でリセット
            With st
                .MaxYear = .MinYear
                .MaxMonth = .MinMonth
                .MaxDay = .MinDay
                .MaxDayOfWeek = .MinDayOfWeek
                .MaxHour = 0
                .MaxMinute = 0
                .MaxSecond = 0
            End With
        End If
        lngSMResult = SendMessage(lnghwndDTP(Index), _
                                  DTM_SETRANGE, GDTR_MIN Or GDTR_MAX, st)
    End If
End Sub

'-------------------------------------------------------------------------------
' MaxDate(選択可能日付範囲の上限)
Public Property Get MaxDate(Optional ByVal Index As Integer = 0) As Date
Dim st As SYSTEMTIMERANGE

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    If (blnCreate = True) Then
        Select Case Index
          Case 0        '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETRANGE, GDT_VALID, st)
            MaxDate = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETRANGE, GDT_VALID, st)
            MaxDate = DateSerial(st.MaxYear, st.MaxMonth, st.MaxDay)
          Case Else
            MaxDate = CDate(0)
        End Select
    Else
      MaxDate = CDate(0)
    End If
End Property

Public Property Let MaxDate(Optional ByVal Index As Integer = 0, _
                            ByVal dtmNewValue As Date)
Dim i As Integer
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                Call MaxDate_Sub(i, dtmNewValue)
            Next i
          Case 1 To colBaseCtrl.Count
            Call MaxDate_Sub(Index, dtmNewValue)
          Case Else
        End Select
    Else
    End If
End Property

Private Sub MaxDate_Sub(ByVal Index As Integer, _
                        ByVal dtmNewValue As Date)
Dim dtmMin As Date
Dim st As SYSTEMTIMERANGE

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    ' 現在設定されている内容を取得する
    lngSMResult = SendMessage(lnghwndDTP(Index), _
                              DTM_GETRANGE, GDT_VALID, st)
    
    ' 取得したSYSTEMTIMERANGEの[Max]を引数の日付に置換
    With st
        .MaxYear = Year(dtmNewValue)
        .MaxMonth = Month(dtmNewValue)
        .MaxDay = Day(dtmNewValue)
        .MaxDayOfWeek = Weekday(dtmNewValue) - 1    ' APIでは日曜日が0
        .MaxHour = 0
        .MaxMinute = 0
        .MaxSecond = 0
    End With
    
    If (st.MinYear = 0) Then
        ' 下限が設定されていないので上限のみ設定する
        lngSMResult = SendMessage(lnghwndDTP(Index), _
                                  DTM_SETRANGE, GDTR_MAX, st)
    Else
        ' 既に設定されている下限とともに設定
        dtmMin = DateSerial(st.MinYear, st.MinMonth, st.MinDay)
        If (dtmMin > dtmNewValue) Then
            '下限が新しい上限値より未来になるので、上限値でリセット
            With st
                .MinYear = .MaxYear
                .MinMonth = .MaxMonth
                .MinDay = .MaxDay
                .MinDayOfWeek = .MaxDayOfWeek
                .MinHour = 0
                .MinMinute = 0
                .MinSecond = 0
            End With
        End If
        lngSMResult = SendMessage(lnghwndDTP(Index), _
                                  DTM_SETRANGE, GDTR_MAX Or GDTR_MIN, st)
    End If
End Sub

'--------------------------------------------------------------
'カレンダー背景色
Public Property Get CalendarBackColor(Optional ByVal Index As Integer = 0) As Long
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETMCCOLOR, MCSC_MONTHBK, ByVal GDT_VALID)
            CalendarBackColor = CLng(lngSMResult)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETMCCOLOR, MCSC_MONTHBK, ByVal GDT_VALID)
            CalendarBackColor = CLng(lngSMResult)
          Case Else
            CalendarBackColor = -1
        End Select
    Else
        CalendarBackColor = -1
    End If
End Property

Public Property Let CalendarBackColor(Optional ByVal Index As Integer = 0, _
                                      ByVal NewColor As Long)
Dim lngRGB As Long
Dim i As Integer

#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If

    If (blnCreate = True) Then
        Select Case NewColor
          'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
          Case vbScrollBars To vbInfoBackground
            lngRGB = GetSysColor(NewColor And cstMask)
          Case &H0& To &HFFFFFF
            lngRGB = NewColor
          Case Else
            Exit Property
        End Select

        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, _
                                          MCSC_MONTHBK, ByVal lngRGB)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, _
                                      MCSC_MONTHBK, ByVal lngRGB)
          Case Else
        End Select
    Else
    End If
End Property

'-------------------------------------------------------------
'カレンダー文字色
Public Property Get CalendarForeColor(Optional ByVal Index As Integer = 0) As Long
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETMCCOLOR, MCSC_TEXT, ByVal GDT_VALID)
            CalendarForeColor = CLng(lngSMResult)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETMCCOLOR, MCSC_TEXT, ByVal GDT_VALID)
            CalendarForeColor = CLng(lngSMResult)
          Case Else
            CalendarForeColor = -1
        End Select
    Else
        CalendarForeColor = -1
    End If
End Property

Public Property Let CalendarForeColor(Optional ByVal Index As Integer = 0, _
                                      ByVal NewColor As Long)
Dim lngRGB As Long
Dim i As Integer
    
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case NewColor
          'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
          Case vbScrollBars To vbInfoBackground
            lngRGB = GetSysColor(NewColor And cstMask)
          Case &H0& To &HFFFFFF
            lngRGB = NewColor
          Case Else
            Exit Property
        End Select

        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, _
                                          MCSC_TEXT, ByVal lngRGB)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, _
                                      MCSC_TEXT, ByVal lngRGB)
          Case Else
        End Select
    Else
    End If
End Property

'-------------------------------------------------------------
'[前月/翌月]文字色
Public Property Get CalendarTrailingForeColor(Optional ByVal Index As Integer = 0) As Long
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETMCCOLOR, MCSC_TRAILINGTEXT, ByVal GDT_VALID)
            CalendarTrailingForeColor = CLng(lngSMResult)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETMCCOLOR, MCSC_TRAILINGTEXT, ByVal GDT_VALID)
            CalendarTrailingForeColor = CLng(lngSMResult)
          Case Else
            CalendarTrailingForeColor = -1
        End Select
    Else
        CalendarTrailingForeColor = -1
    End If
End Property

Public Property Let CalendarTrailingForeColor(Optional ByVal Index As Integer = 0, _
                                              ByVal NewColor As Long)
Dim lngRGB As Long
Dim i As Integer
    
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case NewColor
          'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
          Case vbScrollBars To vbInfoBackground
            lngRGB = GetSysColor(NewColor And cstMask)
          Case &H0& To &HFFFFFF
            lngRGB = NewColor
          Case Else
            Exit Property
        End Select

        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, _
                                          MCSC_TRAILINGTEXT, ByVal lngRGB)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, _
                                      MCSC_TRAILINGTEXT, ByVal lngRGB)
          Case Else
        End Select
    Else
    End If
End Property

'------------------------------------------------------------
'タイトルバー背景色/曜日文字色
Public Property Get CalendarTitleBackColor(Optional ByVal Index As Integer = 0) As Long
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETMCCOLOR, MCSC_TITLEBK, ByVal GDT_VALID)
            CalendarTitleBackColor = CLng(lngSMResult)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETMCCOLOR, MCSC_TITLEBK, ByVal GDT_VALID)
            CalendarTitleBackColor = CLng(lngSMResult)
          Case Else
            CalendarTitleBackColor = -1
        End Select
    Else
        CalendarTitleBackColor = -1
    End If
End Property

Public Property Let CalendarTitleBackColor(Optional ByVal Index As Integer = 0, _
                                           ByVal NewColor As Long)
Dim lngRGB As Long
Dim i As Integer
    
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case NewColor
          'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
          Case vbScrollBars To vbInfoBackground
            lngRGB = GetSysColor(NewColor And cstMask)
          Case &H0& To &HFFFFFF
            lngRGB = NewColor
          Case Else
            Exit Property
        End Select

        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, _
                                          MCSC_TITLEBK, ByVal lngRGB)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, _
                                      MCSC_TITLEBK, ByVal lngRGB)
          Case Else
        End Select
    Else
    End If
End Property

'-------------------------------------------------------------
'タイトルバー文字色
Public Property Get CalendarTitleForeColor(Optional ByVal Index As Integer = 0) As Long
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case Index
          Case 0            '[省略/0]→『１番』
            lngSMResult = SendMessage(lnghwndDTP(1), _
                                      DTM_GETMCCOLOR, MCSC_TITLETEXT, ByVal GDT_VALID)
            CalendarTitleForeColor = CLng(lngSMResult)
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), _
                                      DTM_GETMCCOLOR, MCSC_TITLETEXT, ByVal GDT_VALID)
            CalendarTitleForeColor = CLng(lngSMResult)
          Case Else
            CalendarTitleForeColor = -1
        End Select
    Else
        CalendarTitleForeColor = -1
    End If
End Property

Public Property Let CalendarTitleForeColor(Optional ByVal Index As Integer = 0, _
                                           ByVal NewColor As Long)
Dim lngRGB As Long
Dim i As Integer
    
#If VBA7 And Win64 Then
    Dim lngSMResult As LongPtr
#Else
    Dim lngSMResult As Long
#End If
    
    If (blnCreate = True) Then
        Select Case NewColor
          'vbScrollBars(&H80000000)～vbInfoBackground(&H80000018)
          Case vbScrollBars To vbInfoBackground
            lngRGB = GetSysColor(NewColor And cstMask)
          Case &H0& To &HFFFFFF
            lngRGB = NewColor
          Case Else
            Exit Property
        End Select

        Select Case Index
          Case 0            '[省略/0]→『全て』
            For i = 1 To colBaseCtrl.Count
                lngSMResult = SendMessage(lnghwndDTP(i), DTM_SETMCCOLOR, _
                                          MCSC_TITLETEXT, ByVal lngRGB)
            Next i
          Case 1 To colBaseCtrl.Count
            lngSMResult = SendMessage(lnghwndDTP(Index), DTM_SETMCCOLOR, _
                                      MCSC_TITLETEXT, ByVal lngRGB)
          Case Else
        End Select
    Else
    End If
End Property

'-----------------------------<< End Of Source >>-------------------------------


