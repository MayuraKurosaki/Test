VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListViewEditView 
   Caption         =   "UserForm2"
   ClientHeight    =   5220
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   9200.001
   OleObjectBlob   =   "ListViewEditView.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ListViewEditView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ListViewEdit As EditableListView

Private Sub UserForm_Initialize()
    Set ListViewEdit = New EditableListView
    ListViewEdit.Init Me.FrameListView, Me.Font.Size
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub

Private Sub UserForm_Terminate()
    Set ListViewEdit = Nothing
End Sub
