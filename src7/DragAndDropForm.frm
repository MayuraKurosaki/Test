VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DragAndDropForm 
   Caption         =   "DragAndDropForm"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DragAndDropForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DragAndDropForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame1_Click()
End Sub

Private Sub UserForm_Initialize()
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Label2.Caption = URL
    Cancel = True
End Sub
