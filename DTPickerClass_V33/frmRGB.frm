VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRGB 
   Caption         =   "Color ( BB GG RR )"
   ClientHeight    =   2565
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   4830
   OleObjectBlob   =   "frmRGB.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngColor As Long     'キャンセル時は [ -1 ]
Public strRGB As String     'BBGGRR

Private Sub UserForm_Initialize()
    lngColor = -1
    strRGB = ""
End Sub

Private Sub cmdOK_Click()
    lngColor = Frame1.BackColor
    strRGB = Right("000000" & Hex(lngColor), 6)
    Me.Hide
End Sub

Private Sub scrBlue_Change()
    lblBlue.Caption = Right("00" & Hex(scrBlue.Value), 2)
    Frame1.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
End Sub

Private Sub scrGreen_Change()
    lblGreen.Caption = Right("00" & Hex(scrGreen.Value), 2)
    Frame1.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
End Sub

Private Sub scrRed_Change()
    lblRed.Caption = Right("00" & Hex(scrRed.Value), 2)
    Frame1.BackColor = RGB(scrRed.Value, scrGreen.Value, scrBlue.Value)
End Sub

Public Sub SetRGB(ByVal argColor As Long)
Dim BB As Long
Dim GG As Long
Dim RR As Long

    BB = Int(argColor / 65536)
    GG = Int((argColor - (65536 * BB)) / 256)
    RR = argColor - (65536 * BB) - (256 * GG)

    scrBlue.Value = BB
    scrGreen.Value = GG
    scrRed.Value = RR
End Sub
