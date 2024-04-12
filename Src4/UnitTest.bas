Attribute VB_Name = "UnitTest"
Option Explicit

Public Sub CalendarTest()
    Call DatePicker.Init
    Debug.Print Format(DatePicker.SelectionDate, "YYYY/MM/DD")
End Sub
