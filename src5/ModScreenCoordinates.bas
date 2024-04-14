Attribute VB_Name = "ModScreenCoordinates"
Option Explicit

'�yTODO�z
' - ActiveWindow���𑜓x���Ⴄ���j�^(�f�X�N�g�b�v)�Ԃɂ܂������Ă���ꍇ�ɂ�UserForm�̕\����ʒu������Ă��܂�(��ʒ[�ɃX�i�b�v�����ꍇ���ɂ�����)

' �y�o���z
' - ���ɂ���Ă�ActiveCell��ActiveWindow�̉E�̕��ɂ����UserForm�̕\���ʒu�����ɂ���Ă��Ă��܂��ꍇ������iPointsToScreenPixelsX���������l��Ԃ��Ȃ��͗l�j
'   �� Microsoft 365�̃G�N�Z���ɂāu�t�@�C�����I�v�V�������ݒ聄�S�ʁ����[�U�[ �C���^�[�t�F�C�X�̃I�v�V�����������f�B�X�v���C���g�p����ꍇ�v��
'      �� �\����D�悵���œK�� (�A�v���P�[�V�����̍ċN�����K�v)
'   �@ ��I�����Ă���ƁA�}���`�f�B�X�v���C���ɂĔ�������ꍇ�����邱�Ƃ������ihttps://twitter.com/furyutei/status/1645413942582448129�j
'      ���̏ꍇ�ɂ�
'      ��  �݊����ɑΉ������œK��
'      �ɕύX��A�ċN������΁A���P�����͗l

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
        ' Excel 2010�ȑO�̃o�[�W��������Window�I�u�W�F�N�g�ɂ�hWnd�v���p�e�B���Ȃ��iSDI�̂��߂Ǝv����j
        ' �������Application.hWnd���g�p
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
    '�y�o���z
    '   UserForm�̕\���ʒu(.Top/.Left)�͉�ʍ������_�Ƃ���(�|�C���g����)�w�肷��͂������A���̂��z�肷��ʒu���E���ɂ���Ă��܂����Ƃ�����
    '   ���̏ꍇ�͌o����A�\���������ʒu(�|�C���g)�̒l�ɂ���W���i�����ƕ�(.Height/.Width)�̐ݒ�l�Ǝ����l�̔䗦�E��:14/15=0.9333�c�j���|����Ƃ��傤�ǂ悢�ʒu�ƂȂ�͗l
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


