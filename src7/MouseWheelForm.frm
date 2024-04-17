VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MouseWheelForm 
   Caption         =   "MouseWheelForm"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   OleObjectBlob   =   "MouseWheelForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MouseWheelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------
'Elemente:
'   ComboBox1
'   ListBox1
'   ListBox2
'-----------------------------

Private ActControl As Object

'---------
'User Form
'---------

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 10 To 30
        ListBox1.AddItem i & " - ListBox1"
        ListBox2.AddItem i & " - ListBox2"
        ComboBox1.AddItem i & " - ComboBox1"
    Next
    ComboBox1.ListIndex = 0
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    UnHookControl
End Sub

'----------------------------
'CombBox1, ListBox1, ListBox2
'----------------------------

Private Sub ComboBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    HookControl ComboBox1
End Sub

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    HookControl ListBox1
End Sub

Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    HookControl ListBox2
End Sub
