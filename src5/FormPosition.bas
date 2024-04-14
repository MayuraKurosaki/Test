Attribute VB_Name = "FormPosition"
Option Explicit
Option Private Module
'===================================================================================================
#Const g_cnsAdjust = 1          ' ← 0=下端・右端制御しない、1=下端・右端制御する
'---------------------------------------------------------------------------------------------------
Public Const g_cnsTitle As String = "ユーザーフォーム位置制御テスト"
'---------------------------------------------------------------------------------------------------
' フォーム位置制御関連
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const SM_CYSCREEN As Long = 1
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79
Private Const SPI_GETWORKAREA As Long = 48
'---------------------------------------------------------------------------------------------------
' GetWindowRect用ユーザー定義
Private Type g_typRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' 64ビット版判定
#If VBA7 Then
' ■GetDC(API)
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
' ■ReleaseDC(API)
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" _
    (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
' ■GetDeviceCaps(API)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" _
    (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
' ■GetSystemMetrics(API)
Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
' ■SystemParametersInfo(API)
Private Declare PtrSafe Function SystemParametersInfo Lib "user32.dll" _
    Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As g_typRect, _
    ByVal fuWinIni As Long) As Long
#Else
' ■GetDC(API)
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
' ■ReleaseDC(API)
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
' ■GetDeviceCaps(API)
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
' ■GetSystemMetrics(API)
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
' ■SystemParametersInfo(API)
Private Declare Function SystemParametersInfo Lib "user32.dll" _
    Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As g_typRect, _
    ByVal fuWinIni As Long) As Long
#End If

'***************************************************************************************************
'　■■■ ワークシートからの呼び出し処理 ■■■
'***************************************************************************************************
'* 処理名　：ShowFormFromRange
'* 機能　　：セル(Range)から表示させる
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = セル(Object) ※単一セル又は結合したセル
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Public Sub ShowFormFromRange(ByRef objRange As Range)
    '-----------------------------------------------------------------------------------------------
    Dim lngLeft As Long                                             ' 横位置
    Dim lngTop As Long                                              ' 縦位置
    ' 非結合のセル範囲を選択している時は処理しない
    If objRange.Count > 1 Then
        ' 単一結合セルはOK とする
        If objRange.Address <> objRange.Cells(1).MergeArea.Address Then Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------
    ' ユーザーフォーム表示位置取得
    Call FP_GetFormPosition(objRange, TestForm.Width, TestForm.Height, lngLeft, lngTop)
    '-----------------------------------------------------------------------------------------------
    ' テストフォーム
    With TestForm
        ' フォーム表示位置の確認
        If ((lngLeft <> 0) Or (lngTop <> 0)) Then
            ' 指定がある場合はマニュアル指定
            .StartUpPosition = 0
            .Left = lngLeft
            .Top = lngTop
        Else
            ' 指定がない場合はスクリーンの中央
            .StartUpPosition = 2
        End If
        ' テストフォームを表示
        .Show
    End With
End Sub

'***************************************************************************************************
'　■■■ サブ処理 ■■■
'***************************************************************************************************
'* 処理名　：FP_GetFormPosition
'* 機能　　：ユーザーフォーム表示位置取得
'---------------------------------------------------------------------------------------------------
'* 返り値　：処理成否(Boolean)
'* 引数　　：Arg1 = 対象セル(Object)
'* 　　　　　Arg2 = ユーザーフォームの幅(Long)
'* 　　　　　Arg3 = ユーザーフォームの高さ(Long)
'* 　　　　　Arg4 = スクリーン上の横位置(Long)          ※Ref参照
'* 　　　　　Arg5 = スクリーン上の縦位置(Long)          ※Ref参照
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Function FP_GetFormPosition(ByRef objRange As Range, _
                                    ByVal lngFormWidth As Long, _
                                    ByVal lngFormHeight As Long, _
                                    ByRef lngFormLeft As Long, _
                                    ByRef lngFormTop As Long) As Boolean
    '-----------------------------------------------------------------------------------------------
    Dim objTarget As Range                                          ' 対象セル(先頭セル)
    Dim objAW As Window                                             ' ActiveWindow
    Dim lngPaneIx As Long                                           ' PaneIndex(0〜4)
    Dim lngIx As Long                                               ' ループ用INDEX(Work)
    Dim lngR1C1Left As Long                                         ' 起点セル左端位置
    Dim lngR1C1Top As Long                                          ' 起点セル上端位置
    Dim lngTargetLeft As Long                                       ' 対象セル左端位置
    Dim lngTargetTop As Long                                        ' 対象セル上端位置
    Dim lngScreenRight As Long                                      ' スクリーン右端位置
    Dim lngScreenBottom As Long                                     ' スクリーン下端位置
    Dim lngDPIX As Long                                             ' Dots Per Inch(水平)
    Dim lngDPIY As Long                                             ' Dots Per Inch(垂直)
    Dim lngPPI As Long                                              ' Pixels Per Inch
    FP_GetFormPosition = False
    lngFormLeft = 0
    lngFormTop = 0
    lngPaneIx = 0
    Set objTarget = objRange.Cells(1).MergeArea
    Set objAW = ActiveWindow
    '-----------------------------------------------------------------------------------------------
    ' ウィンドウ分割無しか
    If Not objAW.FreezePanes And Not objAW.Split Then
        ' 表示域外は無視
        If Intersect(objAW.VisibleRange, objTarget) Is Nothing Then Exit Function
    Else            ' 分割あり
        ' ウィンドウ枠固定か
        If objAW.FreezePanes Then
            ' どのウィンドウに属するか判定
            For lngIx = 1 To objAW.Panes.Count
                ' 発見？
                If Not Intersect(objAW.Panes(lngIx).VisibleRange, objTarget) Is Nothing Then
                    lngPaneIx = lngIx
                    Exit For
                End If
            Next lngIx
            ' 見つからないか
            If lngPaneIx = 0 Then Exit Function
        Else
            ' ウィンドウ分割はアクティブペインのみ判定
            If Not Intersect(objAW.ActivePane.VisibleRange, objTarget) Is Nothing Then
                lngPaneIx = objAW.ActivePane.index
            Else
                Exit Function
            End If
        End If
    End If
    '-----------------------------------------------------------------------------------------------
    ' ※以下はExcel2003以前では動作しない
    lngDPIX = FP_GetDPIX
    lngDPIY = FP_GetDPIY
    lngPPI = FP_GetPPI
    ' ウィンドウ分割無しか
    If lngPaneIx = 0 Then
        lngR1C1Left = objAW.PointsToScreenPixelsX(0)
        lngR1C1Top = objAW.PointsToScreenPixelsY(0)
    Else
        lngR1C1Left = objAW.Panes(lngPaneIx).PointsToScreenPixelsX(0)
        lngR1C1Top = objAW.Panes(lngPaneIx).PointsToScreenPixelsY(0)
    End If
    lngTargetLeft = ((objTarget.Left * (lngDPIX / lngPPI)) * (objAW.Zoom / 100)) + lngR1C1Left
    lngTargetTop = (((objTarget.Top + objTarget.Height) * (lngDPIY / lngPPI)) * (objAW.Zoom / 100)) + lngR1C1Top
    lngFormLeft = lngTargetLeft * (lngPPI / lngDPIX)
    lngFormTop = lngTargetTop * (lngPPI / lngDPIY)
    '-----------------------------------------------------------------------------------------------
    ' 下端・右端制御しない時は終了
#If g_cnsAdjust <> 1 Then
    FP_GetFormPosition = True
    Exit Function
#End If
    ' スクリーンサイズ位置の取得
    Call GP_GetScreenPos(0, 0, lngScreenRight, lngScreenBottom)
    '-----------------------------------------------------------------------------------------------
    ' ユーザーフォームがスクリーンからはみ出すか(横)
    If (lngFormLeft + lngFormWidth) * (lngDPIX / lngPPI) > lngScreenRight Then
        ' スクリーン右端に移動(+3は誤差？)
        lngFormLeft = lngScreenRight * (lngPPI / lngDPIX) - lngFormWidth + 3
    End If
    ' ユーザーフォームがスクリーンからはみ出すか(縦)
    If (lngFormTop + lngFormHeight) * (lngDPIY / lngPPI) > lngScreenBottom Then
        ' セル上端に移動
        lngFormTop = lngFormTop - (objRange.Height + lngFormHeight)
    End If
    FP_GetFormPosition = True
End Function

'***************************************************************************************************
'* 処理名　：FP_GetPPI
'* 機能　　：PPI(Pixels Per Inch)取得
'---------------------------------------------------------------------------------------------------
'* 返り値　：PPI値(Long)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Function FP_GetPPI() As Long
    '-----------------------------------------------------------------------------------------------
    FP_GetPPI = Application.InchesToPoints(1)
End Function

'***************************************************************************************************
'* 処理名　：FP_GetDPIX
'* 機能　　：DPI(Dots Per Inch)取得(水平方向)
'---------------------------------------------------------------------------------------------------
'* 返り値　：DPI値(Long)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Function FP_GetDPIX() As Long
    '-----------------------------------------------------------------------------------------------
    FP_GetDPIX = FP_GetDPI(LOGPIXELSX)
End Function

'***************************************************************************************************
'* 処理名　：FP_GetDPIY
'* 機能　　：DPI(Dots Per Inch)取得(垂直方向)
'---------------------------------------------------------------------------------------------------
'* 返り値　：DPI値(Long)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Function FP_GetDPIY() As Long
    '-----------------------------------------------------------------------------------------------
    FP_GetDPIY = FP_GetDPI(LOGPIXELSY)
End Function

'***************************************************************************************************
'* 処理名　：FP_GetDPI
'* 機能　　：DPI(Dots Per Inch)取得(API)
'---------------------------------------------------------------------------------------------------
'* 返り値　：DPI値(Long)
'* 引数　　：Arg1 = nFlag(Long)
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Function FP_GetDPI(ByVal nFlag As Long) As Long
    '-----------------------------------------------------------------------------------------------
#If VBA7 Then
    Dim lngHdc As LongPtr                                           ' ウィンドウハンドルのDC
#Else
    Dim lngHdc As Long                                              ' ウィンドウハンドルのDC
#End If
    lngHdc = GetDC(Application.hWnd)
    FP_GetDPI = GetDeviceCaps(lngHdc, nFlag)
    Call ReleaseDC(&H0, lngHdc)
End Function

'***************************************************************************************************
'* 処理名　：GP_GetScreenPos
'* 機能　　：スクリーン位置の取得
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = スクリーン左端位置(Long)              ※Ref参照
'* 　　　　　Arg2 = スクリーン上端位置(Long)              ※Ref参照
'* 　　　　　Arg3 = スクリーン右端位置(Long)              ※Ref参照
'* 　　　　　Arg4 = スクリーン下端位置(Long)              ※Ref参照
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
Private Sub GP_GetScreenPos(ByRef lngScreenLeft As Long, _
                            ByRef lngScreenTop As Long, _
                            ByRef lngScreenRight As Long, _
                            ByRef lngScreenBottom As Long)
    '-----------------------------------------------------------------------------------------------
    Dim lngWidth As Long                                            ' スクリーンの幅
    Dim lngHeight As Long                                           ' スクリーンの高さ�@
    Dim lngHeight2 As Long                                          ' スクリーンの高さ�A
    Dim lngHeight3 As Long                                          ' スクリーンの高さ�B
    Dim objRect As g_typRect                                        ' Rect
    ' スクリーンの左端､上端､幅､高さの取得(複数スクリーン対応)
    lngScreenLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)         ' 左端
    lngScreenTop = GetSystemMetrics(SM_YVIRTUALSCREEN)          ' 上端
    lngWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)             ' 幅(仮想スクリーン域)
    lngHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)            ' 高さ(仮想スクリーン域)
    lngHeight2 = GetSystemMetrics(SM_CYSCREEN)                  ' 高さ(メインスクリーンのみ)
    ' タスクバーを除くスクリーンの大きさ取得(メインスクリーンのみ)
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, objRect, 0)
    lngHeight3 = objRect.Bottom - objRect.Top                   ' 高さ(メインのタスクバー以外の分)
    ' タスクバーがメインスクリーンの下端にあるものとし、この高さを差し引く
    lngHeight = lngHeight - (lngHeight2 - lngHeight3)
    ' 右端の算出
    lngScreenRight = lngWidth - lngScreenLeft
    ' 下端の算出
    lngScreenBottom = lngHeight - lngScreenTop
End Sub

