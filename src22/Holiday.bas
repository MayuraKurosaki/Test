Attribute VB_Name = "Holiday"
Option Explicit

Private Type Field
    YearMin As Long
    YearMax As Long
End Type

'月日固定の祝日情報
Private Type HolidayInfoMonthDay
    MonthDay As String
    BeginYear As Long
    EndYear As Long
    Name As String
End Type

'月週曜日固定の祝日情報
Private Type HolidayInfoDayOfWeek
    Month As Long
    NthWeek As Long
    DayOfWeek As Long
    BeginYear As Long
    EndYear As Long
    Name As String
End Type

'「国民の祝日に関する法律」施行年月日
Private Const BEGIN_DATE As Date = #7/20/1948#

'「振替休日」施行年月日
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE As Date = #1/1/2007#

'「国民の休日」施行年月日
Private Const NATIONAL_HOLIDAY_BEGIN_DATE As Date = #12/27/1985#

'エラーコード（パラメータ異常）
Private Const ERROR_INVALID_PARAMETER As Long = &H57

'祝日格納用 Dictionary
'Key  : 年月日（DateTime型）
'Item : 祝日名
Private HolidayList As Dictionary

Private This As Field

'指定日が休日か判定する（休日の場合は HolidayName で休日名を返す）
Public Function IsHoliday(ByVal TargetDate As Date, ByRef HolidayName As String) As Boolean
    HolidayName = ""

    '時分秒データを切り捨てる
    TargetDate = VBA.Fix(TargetDate)

    If HolidayList Is Nothing Then
        Call MakeHolidayDictionary(2000, 2050, SheetList.ListObjects("T_月日固定休日"), SheetList.ListObjects("T_月週曜日固定休日"))
    End If
    
    If TargetDate < BEGIN_DATE Then
        Err.Raise ERROR_INVALID_PARAMETER, "IsHoliday", Format$(TargetDate, "yyyy/mm/dd") & "は、適用範囲外です。"
        Exit Function
    ElseIf VBA.Year(TargetDate) > This.YearMax Then
        Err.Raise ERROR_INVALID_PARAMETER, "IsHoliday", This.YearMax + 1 & "年以降は、適用範囲外です。"
        Exit Function
    End If

    IsHoliday = HolidayList.Exists(TargetDate)
    If IsHoliday Then HolidayName = HolidayList(TargetDate)
End Function

'祝日情報を Dictionary に格納する
' リストの開始年、最終年、月日固定リスト、曜日固定リストを指定する
' 月日固定リストは以下の項目を並べたテーブル
'   月日    適用開始年  適用終了年  名前
' 曜日固定リストは以下の項目を並べたテーブル
'   月  週  曜日    適用開始年  適用終了年  名前
Public Sub MakeHolidayDictionary(ByVal YearMin As Long, ByVal YearMax As Long, HolidayInfoMonthDayList As ListObject, HolidayInfoDayOfWeekList As ListObject)
    Dim HolidayInfoMD() As HolidayInfoMonthDay
    Dim HolidayInfoDOW() As HolidayInfoDayOfWeek

    Set HolidayList = New Dictionary
    
    This.YearMin = YearMin
    This.YearMax = YearMax
    
    '月日固定の祝日情報
    Call GetNationalHolidayInfoMD(HolidayInfoMD, HolidayInfoMonthDayList)

    '月週曜日固定の祝日情報
    Call GetNationalHolidayInfoWN(HolidayInfoDOW, HolidayInfoDayOfWeekList)
    
    'Dictionary へ追加
    Call AddToDictionary(HolidayInfoMD, HolidayInfoDOW)
End Sub

'月日固定の祝日情報生成
Private Sub GetNationalHolidayInfoMD(ByRef HolidayInfo() As HolidayInfoMonthDay, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
    
        Dim I As Long
        For I = 1 To .ListRows.Count
            HolidayInfo(I).MonthDay = .ListColumns("月日").DataBodyRange(I)
            HolidayInfo(I).BeginYear = CLng(.ListColumns("適用開始年").DataBodyRange(I))
            HolidayInfo(I).EndYear = CLng(.ListColumns("適用終了年").DataBodyRange(I))
            HolidayInfo(I).Name = .ListColumns("名前").DataBodyRange(I)
        Next I
    End With
End Sub

'月週曜日固定の祝日情報生成
Private Sub GetNationalHolidayInfoWN(ByRef HolidayInfo() As HolidayInfoDayOfWeek, Table As ListObject)
    With Table
        ReDim HolidayInfo(.ListRows.Count)
        
        Dim I As Long
        For I = 1 To .ListRows.Count
            HolidayInfo(I).Month = CLng(.ListColumns("月").DataBodyRange(I))
            HolidayInfo(I).NthWeek = CLng(.ListColumns("週").DataBodyRange(I))
            Select Case .ListColumns("曜日").DataBodyRange(I)
                Case "日": HolidayInfo(I).DayOfWeek = 1
                Case "月": HolidayInfo(I).DayOfWeek = 2
                Case "火": HolidayInfo(I).DayOfWeek = 3
                Case "水": HolidayInfo(I).DayOfWeek = 4
                Case "木": HolidayInfo(I).DayOfWeek = 5
                Case "金": HolidayInfo(I).DayOfWeek = 6
                Case "土": HolidayInfo(I).DayOfWeek = 7
            End Select
            HolidayInfo(I).BeginYear = CLng(.ListColumns("適用開始年").DataBodyRange(I))
            HolidayInfo(I).EndYear = CLng(.ListColumns("適用終了年").DataBodyRange(I))
            HolidayInfo(I).Name = .ListColumns("名前").DataBodyRange(I)
        Next I
    End With
End Sub

'祝日情報をDictionaryへ格納
Private Sub AddToDictionary(ByRef HolidayInfoMD() As HolidayInfoMonthDay, ByRef HolidayInfoDOW() As HolidayInfoDayOfWeek)
    Dim Holiday As Date
    Dim AddedDays As Long
    Dim DateArray() As Date
    Dim Year As Long
    Dim I As Long

    For Year = This.YearMin To This.YearMax
        '年間の祝日格納用配列クリア
        AddedDays = 0
        ReDim DateArray(AddedDays)

        '月日固定の祝日
        For I = 0 To UBound(HolidayInfoMD)
            '適用期間のみを対象とする
            If HolidayInfoMD(I).BeginYear <= Year And HolidayInfoMD(I).EndYear >= Year Then
                Holiday = CDate(CStr(Year) & "/" & HolidayInfoMD(I).MonthDay)

                Call HolidayList.Add(Holiday, HolidayInfoMD(I).Name)

                ReDim Preserve DateArray(AddedDays)
                DateArray(AddedDays) = Holiday
                AddedDays = AddedDays + 1
            End If
        Next I

        '月週曜日固定の祝日
        For I = 0 To UBound(HolidayInfoDOW)
            '適用期間のみを対象とする
            If HolidayInfoDOW(I).BeginYear <= Year And HolidayInfoDOW(I).EndYear >= Year Then
                Holiday = GetNthWeeksDayOfWeek(Year, HolidayInfoDOW(I).Month, HolidayInfoDOW(I).NthWeek, HolidayInfoDOW(I).DayOfWeek)

                Call HolidayList.Add(Holiday, HolidayInfoDOW(I).Name)

                ReDim Preserve DateArray(AddedDays)
                DateArray(AddedDays) = Holiday
                AddedDays = AddedDays + 1
            End If
        Next I

        '春分の日
        Holiday = GetVernalEquinoxDay(Year)
        Call HolidayList.Add(Holiday, "春分の日")

        ReDim Preserve DateArray(AddedDays)
        DateArray(AddedDays) = Holiday
        AddedDays = AddedDays + 1

        '秋分の日
        Holiday = GetAutumnalEquinoxDay(Year)
        Call HolidayList.Add(Holiday, "秋分の日")

        ReDim Preserve DateArray(AddedDays)
        DateArray(AddedDays) = Holiday
        AddedDays = AddedDays + 1

        '振替休日
        For I = 0 To AddedDays - 1
            If ExistsSubstituteHoliday(DateArray(I), Holiday) Then
                Call HolidayList.Add(Holiday, "振替休日")
            End If
        Next I

        '国民の休日
        For I = 0 To AddedDays - 1
            If ExistsNationalHoliday(DateArray(I), Holiday) Then
                Call HolidayList.Add(Holiday, "国民の休日")
            End If
        Next I

        Erase DateArray
    Next Year
End Sub

'振替休日の有無
'　祝日（TargetDate）に対する振替休日の有無（ある場合は、SubstituteHoliday に代入される）
Private Function ExistsSubstituteHoliday(ByVal TargetDate As Date, ByRef SubstituteHoliday As Date) As Boolean
    Dim NextDay As Date

    ExistsSubstituteHoliday = False

    If HolidayList.Exists(TargetDate) = False Then
        'TargetDate が祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If TargetDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And TargetDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(TargetDate) = vbSunday Then
            '祝日が日曜日であれば、翌日（月曜日）が振替休日
            SubstituteHoliday = DateAdd("d", 1, TargetDate)

            ExistsSubstituteHoliday = True
        End If
    ElseIf TargetDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '「国民の祝日」が日曜日に当たるときは、その日後においてその日に最も近い「国民の祝日」でない日を休日とする
        If Weekday(TargetDate) = vbSunday Then
            NextDay = DateAdd("d", 1, TargetDate)

            '直近の祝日でない日を取得
            Do Until HolidayList.Exists(NextDay) = False
                NextDay = DateAdd("d", 1, NextDay)
            Loop

            SubstituteHoliday = NextDay

            ExistsSubstituteHoliday = True
        End If
    End If
End Function

'国民の休日の有無
'　祝日（TargetDate）に対す国民の休日の有無（ある場合は、NationalHoliday に代入される）
Private Function ExistsNationalHoliday(ByVal TargetDate As Date, ByRef NationalHoliday As Date) As Boolean
    Dim BaseDay As Date
    Dim NextDay As Date

    ExistsNationalHoliday = False

    If HolidayList.Exists(TargetDate) = False Then
        'TargetDate が祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If TargetDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        BaseDay = DateAdd("d", 1, TargetDate)

        '直近の祝日でない日を取得
        Do Until HolidayList.Exists(BaseDay) = False
            BaseDay = DateAdd("d", 1, BaseDay)
        Loop

        '日曜日であれば対象外
        If Weekday(BaseDay) <> vbSunday Then
            NextDay = DateAdd("d", 1, BaseDay)

            '翌日が祝日であれば対象
            If HolidayList.Exists(NextDay) = True Then
                ExistsNationalHoliday = True

                NationalHoliday = BaseDay
            End If
        End If
    End If
End Function

'月の第N W曜日の日時を取得
Private Function GetNthWeeksDayOfWeek(ByVal Year As Long, ByVal Month As Long, ByVal Nth As Long, ByVal DayOfWeek As VbDayOfWeek) As Date
    Dim FirstDate As Date
    Dim DayOfWeekFirst As Long
    Dim Offset As Long

    '指定年月の１日を取得
    FirstDate = DateSerial(Year, Month, 1)

    '１日の曜日を取得
    DayOfWeekFirst = Weekday(FirstDate)

    '指定日へのオフセットを取得
    Offset = DayOfWeek - DayOfWeekFirst

    If DayOfWeekFirst > DayOfWeek Then
        Offset = Offset + 7
    End If

    Offset = Offset + 7 * (Nth - 1)

    GetNthWeeksDayOfWeek = DateAdd("d", Offset, FirstDate)
End Function

'春分の日
Private Function GetVernalEquinoxDay(ByVal Year As Long) As Date
    Dim Day As Long

    Day = Int(20.8431 + 0.242194 * (Year - 1980) - Int((Year - 1980) / 4))

    GetVernalEquinoxDay = DateSerial(Year, 3, Day)
End Function

'秋分の日
Private Function GetAutumnalEquinoxDay(ByVal Year As Long) As Date
    Dim Day As Long

    Day = Int(23.2488 + 0.242194 * (Year - 1980) - Int((Year - 1980) / 4))

    GetAutumnalEquinoxDay = DateSerial(Year, 9, Day)
End Function

