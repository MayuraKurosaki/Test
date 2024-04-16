Attribute VB_Name = "MouseEventTest"
Option Explicit

Sub test1_EventForm()
    MouseEventTestForm1.Show vbModeless
End Sub

Sub test2_EventForm()
    MouseEventTestForm1.StartUpPosition = 2
    MouseEventTestForm1.Show vbModeless
    With MouseEventTestForm2
        .StartUpPosition = 0
        .Top = MouseEventTestForm1.Top
        .Left = MouseEventTestForm1.Left + MouseEventTestForm1.Width
        .Show vbModeless
    End With
End Sub
