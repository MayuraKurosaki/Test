Attribute VB_Name = "ModScreenCoordinates"
Option Explicit

'【TODO】
' - ActiveWindowが解像度が違うモニタ(デスクトップ)間にまたがっている場合にはUserFormの表示や位置が崩れてしまう(画面端にスナップした場合等にも発生)

' 【覚書】
' - 環境によってはActiveCellがActiveWindowの右の方にあるとUserFormの表示位置が左によってきてしまう場合がある（PointsToScreenPixelsXが正しい値を返さない模様）
'   → Microsoft 365のエクセルにて「ファイル＞オプション＞設定＞全般＞ユーザー インターフェイスのオプション＞複数ディスプレイを使用する場合」で
'      ○ 表示を優先した最適化 (アプリケーションの再起動が必要)
'   　 を選択していると、マルチディスプレイ環境にて発生する場合があることが判明（https://twitter.com/furyutei/status/1645413942582448129）
'      この場合には
'      ○  互換性に対応した最適化
'      に変更後、再起動すれば、改善される模様

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function WindowFromObject Lib "oleacc" Alias "WindowFromAccessibleObject" (ByVal pacc As Object, phwnd As LongPtr) As LongPtr

Type ScreenPosition
    x As Double
    y As Double
End Type

Type DotsPerPoint
    x As Double
    y As Double
End Type

Type CoordinateFactor
    x As Double
    y As Double
End Type

Function ConvertToScreenPosition(TargetTop As Double, TargetLeft As Double, Optional TargetWindow As Window) As ScreenPosition
    If TargetWindow Is Nothing Then Set TargetWindow = ActiveWindow
    Dim Result As ScreenPosition
    Dim PaneIndex As Long, WorkPane As Pane, WorkRange As Range
    For PaneIndex = 1 To TargetWindow.Panes.Count
        Set WorkPane = TargetWindow.Panes(PaneIndex)
        With WorkPane.VisibleRange
            If _
                ((.Top <= TargetTop) And (TargetTop < .Top + .Height)) And _
                ((.Left <= TargetLeft) And (TargetLeft < .Left + .Width)) _
            Then
                Result.x = WorkPane.PointsToScreenPixelsX(TargetLeft)
                Result.y = WorkPane.PointsToScreenPixelsY(TargetTop)
                Exit For
            End If
        End With
    Next
    ConvertToScreenPosition = Result
End Function

Function GetDisplayDotsPerPoint(Optional TargetWindow As Window) As DotsPerPoint
    If TargetWindow Is Nothing Then Set TargetWindow = ActiveWindow
    Dim Result As DotsPerPoint
    Dim WindowRect As RECT
    If Application.Version < 15# Then
        ' Excel 2010以前のバージョンだとWindowオブジェクトにはhWndプロパティがない（SDIのためと思われる）
        ' →代わりにApplication.hWndを使用
        Call GetWindowRect(Application.hWnd, WindowRect)
    Else
        Call GetWindowRect(TargetWindow.hWnd, WindowRect)
    End If
    With WindowRect
        Result.x = (.Bottom - .Top) / TargetWindow.Height
        Result.y = (.Right - .Left) / TargetWindow.Width
    End With
    GetDisplayDotsPerPoint = Result
End Function

Function SetUserFormPosition(TargetForm, Top As Double, Left As Double, Optional Calibration As Boolean = True) As CoordinateFactor
    '【覚書】
    '   UserFormの表示位置(.Top/.Left)は画面左上を基点として(ポイント数で)指定するはずだが、何故か想定する位置より右下にずれてしまうことがある
    '   その場合は経験上、表示したい位置(ポイント)の値にある係数（高さと幅(.Height/.Width)の設定値と実測値の比率・例:14/15=0.9333…）を掛けるとちょうどよい位置となる模様
    Dim Result As CoordinateFactor
    With TargetForm
        .StartUpPosition = 0 ' 0:Manual, 1:CenterOwner, 2:CenterScreen, 3:WindowsDefault
        .Top = Top
        .Left = Left
        If Not Calibration Then
            Result.y = 1#: Result.x = 1#
            GoTo CLEANUP
        End If
        Dim Height As Double: Height = .Height
        Dim Width As Double: Width = .Width
        Dim FormhWnd As LongPtr: Call WindowFromObject(TargetForm, FormhWnd)
        Dim FormRect As RECT: Call GetWindowRect(FormhWnd, FormRect)
        Dim TargetDisplayDotsPerPoint As DotsPerPoint: TargetDisplayDotsPerPoint = GetDisplayDotsPerPoint()
        Dim ActualHeight As Double: ActualHeight = (FormRect.Bottom - FormRect.Top) / TargetDisplayDotsPerPoint.y
        Dim ActualWidth As Double: ActualWidth = (FormRect.Right - FormRect.Left) / TargetDisplayDotsPerPoint.x
        Dim FormCoordinateFactorY As Double: FormCoordinateFactorY = Height / ActualHeight
        Dim FormCoordinateFactorX As Double: FormCoordinateFactorX = Width / ActualWidth
        .Top = Top * FormCoordinateFactorY
        .Left = Left * FormCoordinateFactorX
        Result.y = FormCoordinateFactorY
        Result.x = FormCoordinateFactorX
    End With

CLEANUP:
    SetUserFormPosition = Result
End Function


