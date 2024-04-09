VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If VBA.IsDate(Target.Value) Or VBA.IsDate(Target.Offset(-1)) Or InStr(1, UCase$(Target.Offset(-1).Value), "DATE", vbTextCompare) > 0 Then
        ActiveSheet.Shapes("icon").Top = Target.Top
        ActiveSheet.Shapes("icon").Left = Target.Left + Target.Width + 8
        ActiveSheet.Shapes("icon").Visible = True
    Else
        ActiveSheet.Shapes("icon").Visible = False
    
    End If
End Sub
