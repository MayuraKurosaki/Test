Attribute VB_Name = "ModDisplayUserForm"
Option Explicit

' �� �Z���̉E���Ƀ��[�U�[�t�H�[����\�������

Sub DisplayUserFormOnRightSideOfCell(Target As Range, Optional Calibration As Boolean = True)
    Dim TargetWindow As Window: Set TargetWindow = ActiveWindow
    
    ' �f�B�X�v���C���W�n��̑ΏۃZ��(Target)�̈ʒu���擾(TargetDisplayPosition��.x/.y�͋��Ƀh�b�g(�s�N�Z��)�P��)
    Dim TargetDisplayPosition As ScreenPosition: TargetDisplayPosition = ConvertToScreenPosition(Target.Top, Target.Left, TargetWindow)
    
    ' �f�B�X�v���C���W�n��ł�1�|�C���g������̃h�b�g(�s�N�Z��)�����擾
    Dim TargetDisplayDotsPerPoint As DotsPerPoint: TargetDisplayDotsPerPoint = GetDisplayDotsPerPoint(TargetWindow)
    
    ' ���[�U�[�t�H�[���\���ʒu�̃I�t�Z�b�g�v�Z�i�����ΏۃZ���̍��ォ��ǂ�قǂ��炷���j
    ' ���P�ʂ̓f�B�X�v���C���W�n��ł̃|�C���g���ł��邽�߁A���[�N�V�[�g���W�n�̊g�嗦�ŏ悶��K�v����
    Dim WindowScale As Double: WindowScale = TargetWindow.Zoom / 100# ' ���[�N�V�[�g���W�n�̊g�嗦
    Dim FormTopOffset As Double: FormTopOffset = 0 * WindowScale ' Y�����͂��炳�Ȃ�
    Dim FormLeftOffset As Double: FormLeftOffset = Target.MergeArea.Width * WindowScale ' X�����͑ΏۃZ���̕����E�ɂ��炷
    
    ' ���[�U�[�t�H�[���̃f�B�X�v���C���W�n��ł̕\���ʒu������i�P�ʁF�f�B�X�v���C���W�n��ł̃|�C���g��)
    Dim FormTop As Double: FormTop = (TargetDisplayPosition.y / TargetDisplayDotsPerPoint.y) + FormTopOffset
    Dim FormLeft As Double: FormLeft = (TargetDisplayPosition.x / TargetDisplayDotsPerPoint.x) + FormLeftOffset
    
    ' ���[�U�[�t�H�[�����w��ʒu�ɐݒ�
    Dim FormCoordinateFactor As CoordinateFactor: FormCoordinateFactor = SetUserFormPosition(UserForm1, FormTop, FormLeft, Calibration:=Calibration)
    
    ' ���W�^�T�C�Y����TextBox1�ɏo��(�f�o�b�O�p)
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
    
    ' ���[�U�[�t�H�[���̕\��
    Call UserForm1.Show(vbModal)
End Sub
