Attribute VB_Name = "Holiday"
Option Explicit

Private Type Field
    YearMin As Long
    YearMax As Long
End Type

'�����Œ�̏j�����
Private Type HolidayInfoMonthDay
    MonthDay As String
    BeginYear As Long
    EndYear As Long
    Name As String
End Type

'���T�j���Œ�̏j�����
Private Type HolidayInfoDayOfWeek
    Month As Long
    NthWeek As Long
    DayOfWeek As Long
    BeginYear As Long
    EndYear As Long
    Name As String
End Type

'�u�����̏j���Ɋւ���@���v�{�s�N����
Private Const BEGIN_DATE As Date = #7/20/1948#

'�u�U�֋x���v�{�s�N����
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE As Date = #1/1/2007#

'�u�����̋x���v�{�s�N����
Private Const NATIONAL_HOLIDAY_BEGIN_DATE As Date = #12/27/1985#

'�G���[�R�[�h�i�p�����[�^�ُ�j
Private Const ERROR_INVALID_PARAMETER As Long = &H57

'�j���i�[�p Dictionary
'Key  : �N�����iDateTime�^�j
'Item : �j����
Private HolidayList As Dictionary

Private This As Field

'�w������x�������肷��i�x���̏ꍇ�� HolidayName �ŋx������Ԃ��j
Public Function IsHoliday(ByVal TargetDate As Date, ByRef HolidayName As String) As Boolean
    HolidayName = ""

    '�����b�f�[�^��؂�̂Ă�
    TargetDate = VBA.Fix(TargetDate)

    If HolidayList Is Nothing Then
        Call MakeHolidayDictionary(2000, 2050, SheetList.ListObjects("T_�����Œ�x��"), SheetList.ListObjects("T_���T�j���Œ�x��"))
    End If
    
    If TargetDate < BEGIN_DATE Then
        Err.Raise ERROR_INVALID_PARAMETER, "IsHoliday", Format$(TargetDate, "yyyy/mm/dd") & "�́A�K�p�͈͊O�ł��B"
        Exit Function
    ElseIf VBA.Year(TargetDate) > This.YearMax Then
        Err.Raise ERROR_INVALID_PARAMETER, "IsHoliday", This.YearMax + 1 & "�N�ȍ~�́A�K�p�͈͊O�ł��B"
        Exit Function
    End If

    IsHoliday = HolidayList.Exists(TargetDate)
    If IsHoliday Then HolidayName = HolidayList(TargetDate)
End Function

'�j������ Dictionary �Ɋi�[����
' ���X�g�̊J�n�N�A�ŏI�N�A�����Œ胊�X�g�A�j���Œ胊�X�g���w�肷��
' �����Œ胊�X�g�͈ȉ��̍��ڂ���ׂ��e�[�u��
'   ����    �K�p�J�n�N  �K�p�I���N  ���O
' �j���Œ胊�X�g�͈ȉ��̍��ڂ���ׂ��e�[�u��
'   ��  �T  �j��    �K�p�J�n�N  �K�p�I���N  ���O
Public Sub MakeHolidayDictionary(ByVal YearMin As Long, ByVal YearMax As Long, HolidayInfoMonthDayList As ListObject, HolidayInfoDayOfWeekList As ListObject)
    Dim HolidayInfoMD() As HolidayInfoMonthDay
    Dim HolidayInfoDOW() As HolidayInfoDayOfWeek

    Set HolidayList = New Dictionary
    
    This.YearMin = YearMin
    This.YearMax = YearMax
    
    '�����Œ�̏j�����
    Call GetNationalHolidayInfoMD(HolidayInfoMD, HolidayInfoMonthDayList)

    '���T�j���Œ�̏j�����
    Call GetNationalHolidayInfoWN(HolidayInfoDOW, HolidayInfoDayOfWeekList)
    
    'Dictionary �֒ǉ�
    Call AddToDictionary(HolidayInfoMD, HolidayInfoDOW)
End Sub

'�����Œ�̏j����񐶐�
Private Sub GetNationalHolidayInfoMD(ByRef HolidayInfo() As HolidayInfoMonthDay, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
    
        Dim I As Long
        For I = 1 To .ListRows.Count
            HolidayInfo(I).MonthDay = .ListColumns("����").DataBodyRange(I)
            HolidayInfo(I).BeginYear = CLng(.ListColumns("�K�p�J�n�N").DataBodyRange(I))
            HolidayInfo(I).EndYear = CLng(.ListColumns("�K�p�I���N").DataBodyRange(I))
            HolidayInfo(I).Name = .ListColumns("���O").DataBodyRange(I)
        Next I
    End With
End Sub

'���T�j���Œ�̏j����񐶐�
Private Sub GetNationalHolidayInfoWN(ByRef HolidayInfo() As HolidayInfoDayOfWeek, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
        
        Dim I As Long
        For I = 1 To .ListRows.Count
            HolidayInfo(I).Month = CLng(.ListColumns("��").DataBodyRange(I))
            HolidayInfo(I).NthWeek = CLng(.ListColumns("�T").DataBodyRange(I))
            Select Case .ListColumns("�j��").DataBodyRange(I)
                Case "��": HolidayInfo(I).DayOfWeek = 1
                Case "��": HolidayInfo(I).DayOfWeek = 2
                Case "��": HolidayInfo(I).DayOfWeek = 3
                Case "��": HolidayInfo(I).DayOfWeek = 4
                Case "��": HolidayInfo(I).DayOfWeek = 5
                Case "��": HolidayInfo(I).DayOfWeek = 6
                Case "�y": HolidayInfo(I).DayOfWeek = 7
            End Select
            HolidayInfo(I).BeginYear = CLng(.ListColumns("�K�p�J�n�N").DataBodyRange(I))
            HolidayInfo(I).EndYear = CLng(.ListColumns("�K�p�I���N").DataBodyRange(I))
            HolidayInfo(I).Name = .ListColumns("���O").DataBodyRange(I)
        Next I
    End With
End Sub

'�j������Dictionary�֊i�[
Private Sub AddToDictionary(ByRef HolidayInfoMD() As HolidayInfoMonthDay, ByRef HolidayInfoDOW() As HolidayInfoDayOfWeek)
    Dim Holiday As Date
    Dim AddedDays As Long
    Dim DateArray() As Date
    Dim Year As Long
    Dim I As Long

    For Year = This.YearMin To This.YearMax
        '�N�Ԃ̏j���i�[�p�z��N���A
        AddedDays = 0
        ReDim DateArray(AddedDays)

        '�����Œ�̏j��
        For I = 0 To UBound(HolidayInfoMD)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If HolidayInfoMD(I).BeginYear <= Year And HolidayInfoMD(I).EndYear >= Year Then
                Holiday = CDate(CStr(Year) & "/" & HolidayInfoMD(I).MonthDay)

                Call HolidayList.Add(Holiday, HolidayInfoMD(I).Name)

                ReDim Preserve DateArray(AddedDays)
                DateArray(AddedDays) = Holiday
                AddedDays = AddedDays + 1
            End If
        Next I

        '���T�j���Œ�̏j��
        For I = 0 To UBound(HolidayInfoDOW)
            '�K�p���Ԃ݂̂�ΏۂƂ���
            If HolidayInfoDOW(I).BeginYear <= Year And HolidayInfoDOW(I).EndYear >= Year Then
                Holiday = GetNthWeeksDayOfWeek(Year, HolidayInfoDOW(I).Month, HolidayInfoDOW(I).NthWeek, HolidayInfoDOW(I).DayOfWeek)

                Call HolidayList.Add(Holiday, HolidayInfoDOW(I).Name)

                ReDim Preserve DateArray(AddedDays)
                DateArray(AddedDays) = Holiday
                AddedDays = AddedDays + 1
            End If
        Next I

        '�t���̓�
        Holiday = GetVernalEquinoxDay(Year)
        Call HolidayList.Add(Holiday, "�t���̓�")

        ReDim Preserve DateArray(AddedDays)
        DateArray(AddedDays) = Holiday
        AddedDays = AddedDays + 1

        '�H���̓�
        Holiday = GetAutumnalEquinoxDay(Year)
        Call HolidayList.Add(Holiday, "�H���̓�")

        ReDim Preserve DateArray(AddedDays)
        DateArray(AddedDays) = Holiday
        AddedDays = AddedDays + 1

        '�U�֋x��
        For I = 0 To AddedDays - 1
            If ExistsSubstituteHoliday(DateArray(I), Holiday) Then
                Call HolidayList.Add(Holiday, "�U�֋x��")
            End If
        Next I

        '�����̋x��
        For I = 0 To AddedDays - 1
            If ExistsNationalHoliday(DateArray(I), Holiday) Then
                Call HolidayList.Add(Holiday, "�����̋x��")
            End If
        Next I

        Erase DateArray
    Next Year
End Sub

'�U�֋x���̗L��
'�@�j���iTargetDate�j�ɑ΂���U�֋x���̗L���i����ꍇ�́ASubstituteHoliday �ɑ�������j
Private Function ExistsSubstituteHoliday(ByVal TargetDate As Date, ByRef SubstituteHoliday As Date) As Boolean
    Dim NextDay As Date

    ExistsSubstituteHoliday = False

    If HolidayList.Exists(TargetDate) = False Then
        'TargetDate ���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If TargetDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And TargetDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(TargetDate) = vbSunday Then
            '�j�������j���ł���΁A�����i���j���j���U�֋x��
            SubstituteHoliday = DateAdd("d", 1, TargetDate)

            ExistsSubstituteHoliday = True
        End If
    ElseIf TargetDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '�u�����̏j���v�����j���ɓ�����Ƃ��́A���̓���ɂ����Ă��̓��ɍł��߂��u�����̏j���v�łȂ������x���Ƃ���
        If Weekday(TargetDate) = vbSunday Then
            NextDay = DateAdd("d", 1, TargetDate)

            '���߂̏j���łȂ������擾
            Do Until HolidayList.Exists(NextDay) = False
                NextDay = DateAdd("d", 1, NextDay)
            Loop

            SubstituteHoliday = NextDay

            ExistsSubstituteHoliday = True
        End If
    End If
End Function

'�����̋x���̗L��
'�@�j���iTargetDate�j�ɑ΂������̋x���̗L���i����ꍇ�́ANationalHoliday �ɑ�������j
Private Function ExistsNationalHoliday(ByVal TargetDate As Date, ByRef NationalHoliday As Date) As Boolean
    Dim BaseDay As Date
    Dim NextDay As Date

    ExistsNationalHoliday = False

    If HolidayList.Exists(TargetDate) = False Then
        'TargetDate ���j���łȂ���ΏI��
        Exit Function
    End If

    '�K�p���Ԃ݂̂�ΏۂƂ���
    If TargetDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        BaseDay = DateAdd("d", 1, TargetDate)

        '���߂̏j���łȂ������擾
        Do Until HolidayList.Exists(BaseDay) = False
            BaseDay = DateAdd("d", 1, BaseDay)
        Loop

        '���j���ł���ΑΏۊO
        If Weekday(BaseDay) <> vbSunday Then
            NextDay = DateAdd("d", 1, BaseDay)

            '�������j���ł���ΑΏ�
            If HolidayList.Exists(NextDay) = True Then
                ExistsNationalHoliday = True

                NationalHoliday = BaseDay
            End If
        End If
    End If
End Function

'���̑�N W�j���̓������擾
Private Function GetNthWeeksDayOfWeek(ByVal Year As Long, ByVal Month As Long, ByVal Nth As Long, ByVal DayOfWeek As VbDayOfWeek) As Date
    Dim FirstDate As Date
    Dim DayOfWeekFirst As Long
    Dim Offset As Long

    '�w��N���̂P�����擾
    FirstDate = DateSerial(Year, Month, 1)

    '�P���̗j�����擾
    DayOfWeekFirst = Weekday(FirstDate)

    '�w����ւ̃I�t�Z�b�g���擾
    Offset = DayOfWeek - DayOfWeekFirst

    If DayOfWeekFirst > DayOfWeek Then
        Offset = Offset + 7
    End If

    Offset = Offset + 7 * (Nth - 1)

    GetNthWeeksDayOfWeek = DateAdd("d", Offset, FirstDate)
End Function

'�t���̓�
Private Function GetVernalEquinoxDay(ByVal Year As Long) As Date
    Dim Day As Long

    Day = Int(20.8431 + 0.242194 * (Year - 1980) - Int((Year - 1980) / 4))

    GetVernalEquinoxDay = DateSerial(Year, 3, Day)
End Function

'�H���̓�
Private Function GetAutumnalEquinoxDay(ByVal Year As Long) As Date
    Dim Day As Long

    Day = Int(23.2488 + 0.242194 * (Year - 1980) - Int((Year - 1980) / 4))

    GetAutumnalEquinoxDay = DateSerial(Year, 9, Day)
End Function

