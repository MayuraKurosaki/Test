Attribute VB_Name = "ListViewEditModule"
Option Explicit

Public Sub ListViewEditTest()
    ListViewEditView.Show
End Sub

Public Function Redirect(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id&, ByVal lv As EditableListView) As LongPtr
    Redirect = lv.WndProc(hWnd, uMsg, wParam, lParam)
End Function

Public Sub ShowMessage(prm As LongPtr)
    ListViewEditView.Label2.Caption = prm
End Sub

Public Sub ShowMessage2(prm As LongPtr)
    ListViewEditView.Label3.Caption = prm
End Sub

Public Sub ShowMessage3(prm As Long)
    ListViewEditView.Label4.Caption = prm
End Sub

