VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'textbox to update
Dim dateControl As Control

'current date of the calendar
Dim currentDate As Date

'creates the caledar for the month / year
Sub buildCalendar(Optional iYear As Integer, Optional iMonth As Integer)

    Dim startOfMonth As Date: Dim trackingDate As Date
    Dim iStartofMonthDay As Integer: Dim cDay As Control
    
    'if the year or month isn't passed in, then assume current year and month
    If iYear = 0 Or iMonth = 0 Then
        iYear = VBA.Year(Now())
        iMonth = VBA.Month(Now())
    End If
    
    'set the month and year in the calendar
    Me.Controls("lblmonth").Caption = VBA.MonthName(iMonth, True)
    Me.Controls("lblyear").Caption = iYear
    
    'calcuate the start of the month and the start of the calendar (top left day)
    startOfMonth = VBA.DateSerial(iYear, iMonth, 1)
    currentDate = startOfMonth
    iStartofMonthDay = VBA.Weekday(startOfMonth, vbSunday)
    trackingDate = DateAdd("d", -iStartofMonthDay + 1, startOfMonth)
    
    'loop through all the day controls
    For i = 1 To 42
        
        'get and set the day controls
        Set cDay = Me.Controls("day" & i)
        cDay.Caption = VBA.Day(trackingDate)
        cDay.Tag = trackingDate
        
        'make the days not in the current month gray
        If VBA.Month(trackingDate) <> iMonth Then
            cDay.ForeColor = 8421504
        Else
            cDay.ForeColor = 0
        End If
        
        'move to the next day
        trackingDate = VBA.DateAdd("d", 1, trackingDate)
    
    Next i

End Sub

'all the click events call into here with the control id
Sub dayClick(i As Integer)
    dateControl.Text = Me.Controls("day" & i).Tag
    frameDatePicker.Visible = False
End Sub

'move calendar to the previous month
Private Sub btnLastMonth_Click()
    currentDate = DateAdd("m", -1, currentDate)
    buildCalendar VBA.Year(currentDate), VBA.Month(currentDate)
End Sub

'move calendar to the next month
Private Sub btnNextMonth_Click()
    currentDate = DateAdd("m", 1, currentDate)
    buildCalendar VBA.Year(currentDate), VBA.Month(currentDate)
End Sub

'close the form on the ok button
Private Sub btnOK_Click()
    Unload Me
End Sub

'date button for textbox 1
Private Sub dateButton1_Click()
    toggleDatePicker TextBox1
End Sub

'date button for textbox 2
Private Sub dateButton2_Click()
    toggleDatePicker TextBox2
End Sub

'shows / hides date picker
Private Sub toggleDatePicker(oControl As Control)
    
    'check if its visible
    If frameDatePicker.Visible Then
        
        'already there, so toggle it off
        frameDatePicker.Visible = False
        
    Else
        
        'check if there is already a date in the text box
        If VBA.IsDate(oControl.Text) Then
        
            'load the calendar with that date
            buildCalendar VBA.Year(oControl.Text), VBA.Month(oControl.Text)
            
        Else
            
            'no date
            buildCalendar
            
        End If
        
        'set the global control so when the calendar gets clicked
        'we know what textbox to update
        Set dateControl = oControl
        
        'set the position of the calendar
        frameDatePicker.Top = dateControl.Top
        
        'show the calendar
        frameDatePicker.Visible = True
        
    End If
End Sub

'click events for the days
Private Sub day1_Click(): dayClick (1): End Sub
Private Sub day2_Click(): dayClick (2): End Sub
Private Sub day3_Click(): dayClick (3): End Sub
Private Sub day4_Click(): dayClick (4): End Sub
Private Sub day5_Click(): dayClick (5): End Sub
Private Sub day6_Click(): dayClick (6): End Sub
Private Sub day7_Click(): dayClick (7): End Sub
Private Sub day8_Click(): dayClick (8): End Sub
Private Sub day9_Click(): dayClick (9): End Sub
Private Sub day10_Click(): dayClick (10): End Sub
Private Sub day11_Click(): dayClick (11): End Sub
Private Sub day12_Click(): dayClick (12): End Sub
Private Sub day13_Click(): dayClick (13): End Sub
Private Sub day14_Click(): dayClick (14): End Sub
Private Sub day15_Click(): dayClick (15): End Sub
Private Sub day16_Click(): dayClick (16): End Sub
Private Sub day17_Click(): dayClick (17): End Sub
Private Sub day18_Click(): dayClick (18): End Sub
Private Sub day19_Click(): dayClick (19): End Sub
Private Sub day20_Click(): dayClick (20): End Sub
Private Sub day21_Click(): dayClick (21): End Sub
Private Sub day22_Click(): dayClick (22): End Sub
Private Sub day23_Click(): dayClick (23): End Sub
Private Sub day24_Click(): dayClick (24): End Sub
Private Sub day25_Click(): dayClick (25): End Sub
Private Sub day26_Click(): dayClick (26): End Sub
Private Sub day27_Click(): dayClick (27): End Sub
Private Sub day28_Click(): dayClick (28): End Sub
Private Sub day29_Click(): dayClick (29): End Sub
Private Sub day30_Click(): dayClick (30): End Sub
Private Sub day31_Click(): dayClick (31): End Sub
Private Sub day32_Click(): dayClick (32): End Sub
Private Sub day33_Click(): dayClick (33): End Sub
Private Sub day34_Click(): dayClick (34): End Sub
Private Sub day35_Click(): dayClick (35): End Sub
Private Sub day36_Click(): dayClick (36): End Sub
Private Sub day37_Click(): dayClick (37): End Sub
Private Sub day38_Click(): dayClick (38): End Sub
Private Sub day39_Click(): dayClick (39): End Sub
Private Sub day40_Click(): dayClick (40): End Sub
Private Sub day41_Click(): dayClick (41): End Sub
Private Sub day42_Click(): dayClick (42): End Sub

'init the date picker
Private Sub UserForm_Initialize()
    buildCalendar
End Sub
