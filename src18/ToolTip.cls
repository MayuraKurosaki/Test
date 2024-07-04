VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Sub ShowListToolTip(CtrlEvt As ControlEvent, ByVal X As Single, ByVal Y As Single)
    Dim tipListIndex As Long
    Dim tipText As String
    Dim itemHeight As Single
    
    With Me.ListBoxEdit
        itemHeight = .Font.size * 1.4
        tipListIndex = Fix(Y / itemHeight + .TopIndex)
        
        If tipListIndex > .ListCount - 1 Then tipListIndex = .ListCount - 1
        
        If .ColumnHeads Then tipListIndex = tipListIndex - 1
        If tipListIndex < 0 Then CloseListToolTip: Exit Sub
        tipText = ListBoxHeaderText(1) & ":" & .List(tipListIndex, 1) & vbLf & _
                  ListBoxHeaderText(2) & ":" & .List(tipListIndex, 2) & vbLf & _
                  ListBoxHeaderText(3) & ":" & .List(tipListIndex, 3)
    End With
    Me.LabelToolTip.Caption = tipText
    Call AutoFitControl(Me.LabelToolTip)
'    Me.LabelToolTip.AutoSize = True
    
    
    Dim tipTop As Single, tipLeft As Single
    tipTop = Me.MultiPageSwitchMode.Top + CtrlEvt.Control.Top + Y + itemHeight
    tipLeft = Me.MultiPageSwitchMode.Left + CtrlEvt.Control.Left + X
    
    With Me.FrameToolTip
        .Height = Me.LabelToolTip.Height
        .Width = Me.LabelToolTip.Width
        If tipTop + .Height > .Parent.InsideHeight Then tipTop = Me.MultiPageSwitchMode.Top + CtrlEvt.Control.Top + Y - itemHeight - .Height
        If tipLeft + .Width > .Parent.InsideWidth Then tipLeft = tipLeft - .Width
        .Top = tipTop
        .Left = tipLeft
        
        If Not isShowToolTip Then
            .Visible = True
            isShowToolTip = True
        End If
    End With
End Sub

Public Sub CloseToolTip()
    With Me.FrameToolTip
        .Caption = ""
        .Visible = False
    End With
    
    isShowToolTip = False
End Sub

