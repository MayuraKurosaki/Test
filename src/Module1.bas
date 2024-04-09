Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.Unprotect
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("E8").Select
End Sub
