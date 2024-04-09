VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'hooks into Excel objects
Private WithEvents c1App As Application
Attribute c1App.VB_VarHelpID = -1

'local vars
Private m_PickerVisible As Boolean

'**************************************************************
'CLASS EVENTS
'**************************************************************

'INIT
Private Sub Class_Initialize()
    On Error Resume Next

    'get the application object
    Set c1App = Application
    
    'set the defaults
    m_PickerVisible = False
    
    On Error GoTo 0
End Sub

'TERMINATE
Private Sub Class_Terminate()
    clearAllUI
End Sub

'**************************************************************
'PROPERTIES
'**************************************************************

'Date Picker Visible Tracking
Public Property Get PickerVisible() As Variant
    PickerVisible = m_PickerVisible
End Property
Public Property Let PickerVisible(ByVal vNewValue As Variant)
    m_PickerVisible = vNewValue
End Property

'**************************************************************
'SHARED UPDATES
'**************************************************************

'sheet changed
Private Sub DoSheetChange()
    clearAllUI
    processSelectionChange
End Sub

'workbook changed
Private Sub DoWorkbookChange()
    DoSheetChange
End Sub

'clear all UI
Private Sub clearAllUI()
    closeDatePicker
    killGridDP
End Sub

'process the cell changing
Private Sub processSelectionChange()
    'clear the ui
    clearAllUI
    
    'see if we should show the DP icon
    If checkForGridDP Then
        'show the date picker
        createGridDP
    End If
End Sub

'called to see if we should show the dp entry point in the grid
Private Function checkForGridDP() As Boolean
    On Error GoTo check_err
    
    'ref to cell above
    Dim cellAbove As Range
    
    'default to false
    checkForGridDP = False
    
    'be sure the flag is enabled
    If (fShowDPInGrid) Then
        'check to see if active cell is a range
        If TypeName(ActiveCell) = "Range" Then
            'check the current cell's value
            If VBA.IsDate(ActiveCell) Or ActiveCell.NumberFormat Like "*[Dd]*" Then
                'show the shape dp
                checkForGridDP = True
            Else
                'exit out if its populated but not a date
                If ActiveCell <> "" Then Exit Function
                
                'make sure we aren't in the top row
                If ActiveCell.Row <> 1 Then
                    'get the cell above
                    Set cellAbove = ActiveCell.Offset(-1, 0)
                    
                    'check to see if the cell above has the word date in it, or if its a date
                    If InStr(1, cellAbove, "date", vbTextCompare) > 0 Or VBA.IsDate(cellAbove) Then
                        'show the shape dp
                        checkForGridDP = True
                    End If
                End If
            End If
        End If
    End If
    
check_err:
End Function

'**************************************************************
'APP EVENTS
'**************************************************************

Private Sub c1App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    DoSheetChange
End Sub

'App New workbook created
Private Sub c1App_NewWorkbook(ByVal Wb As Workbook)
    DoWorkbookChange
End Sub

'App before print
Private Sub c1App_WorkbookBeforePrint(ByVal Wb As Workbook, Cancel As Boolean)
    clearAllUI
End Sub

'App before save
Private Sub c1App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    clearAllUI
End Sub

'App window resize
Private Sub c1App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    closeDatePicker
End Sub

'App sheet activated
Private Sub c1App_SheetActivate(ByVal Sh As Object)
    DoSheetChange
End Sub

'App workbook activated
Private Sub c1App_WorkbookActivate(ByVal Wb As Workbook)
    DoWorkbookChange
End Sub
