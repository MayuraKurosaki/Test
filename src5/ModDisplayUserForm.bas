Attribute VB_Name = "ModDisplayUserForm"
Option Explicit

' ■ セルの右横にユーザーフォームを表示する例

Sub DisplayUserFormOnRightSideOfCell(Target As Range, Optional Calibration As Boolean = True)
    Dim TargetWindow As Window: Set TargetWindow = ActiveWindow
    
    ' ディスプレイ座標系上の対象セル(Target)の位置を取得(TargetDisplayPositionの.x/.yは共にドット(ピクセル)単位)
    Dim TargetDisplayPosition As ScreenPosition: TargetDisplayPosition = ConvertToScreenPosition(Target.Top, Target.Left, TargetWindow)
    
    ' ディスプレイ座標系上での1ポイントあたりのドット(ピクセル)数を取得
    Dim TargetDisplayDotsPerPoint As DotsPerPoint: TargetDisplayDotsPerPoint = GetDisplayDotsPerPoint(TargetWindow)
    
    ' ユーザーフォーム表示位置のオフセット計算（左上を対象セルの左上からどれほどずらすか）
    ' ※単位はディスプレイ座標系上でのポイント数であるため、ワークシート座標系の拡大率で乗じる必要あり
    Dim WindowScale As Double: WindowScale = TargetWindow.Zoom / 100# ' ワークシート座標系の拡大率
    Dim FormTopOffset As Double: FormTopOffset = 0 * WindowScale ' Y方向はずらさない
    Dim FormLeftOffset As Double: FormLeftOffset = Target.MergeArea.Width * WindowScale ' X方向は対象セルの幅分右にずらす
    
    ' ユーザーフォームのディスプレイ座標系上での表示位置を決定（単位：ディスプレイ座標系上でのポイント数)
    Dim FormTop As Double: FormTop = (TargetDisplayPosition.y / TargetDisplayDotsPerPoint.y) + FormTopOffset
    Dim FormLeft As Double: FormLeft = (TargetDisplayPosition.x / TargetDisplayDotsPerPoint.x) + FormLeftOffset
    
    ' ユーザーフォームを指定位置に設定
    Dim FormCoordinateFactor As CoordinateFactor: FormCoordinateFactor = SetUserFormPosition(UserForm1, FormTop, FormLeft, Calibration:=Calibration)
    
    ' 座標／サイズ情報をTextBox1に出力(デバッグ用)
    With UserForm1
        .TextBox1.Text = Join(Array( _
            "[Cell] Height=" & Format(Target.Height, "0.00") & " Width=" & Format(Target.Width, "0.00"), _
            "  Top=" & Format(Target.Top, "0.00") & " Left=" & Format(Target.Left, "0.00"), _
            "  DPPY=" & Format(TargetDisplayDotsPerPoint.y, "0.00") & " DPPX=" & Format(TargetDisplayDotsPerPoint.x, "0.00"), _
            "  PosY=" & TargetDisplayPosition.y & " PosX=" & TargetDisplayPosition.x, _
            "", _
            "[Form] Height=" & Format(.Height, "0.00") & " Width=" & Format(.Width, "0.00"), _
            " Calibration: " & IIf(Calibration, "On", "Off"), _
            " Specified:", _
            "  Top=" & Format(FormTop, "0.00") & " Left=" & Format(FormLeft, "0.00"), _
            " Coordinate Factors:", _
            "  CFY=" & Format(FormCoordinateFactor.y, "0.00") & " CFX=" & Format(FormCoordinateFactor.x, "0.00"), _
            " Actual:", _
            "  Top=" & Format(.Top, "0.00") & " Left=" & Format(.Left, "0.00") _
        ), vbLf)
    End With
    
    ' ユーザーフォームの表示
    Call UserForm1.Show(vbModal)
End Sub
