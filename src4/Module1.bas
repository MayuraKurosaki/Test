Attribute VB_Name = "Module1"
Option Explicit

'Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Declare PtrSafe Function DllGetVersion Lib "shell32.dll" (pdwVersion As DLLVERSIONINFO) As Long
Declare PtrSafe Function ComCtlDllGetVersion Lib "comctl32.dll" Alias "DllGetVersion" (pdwVersion As DLLVERSIONINFO) As Long

' DLLVERSIONINFO構造体の定義
Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Sub GetShellVersion()
    Dim dvi As DLLVERSIONINFO
    Dim ret As Long
    
    ' 構造体のサイズを設定する
    dvi.cbSize = Len(dvi)
    
    ' DllGetVersionを呼び出す
    ret = DllGetVersion(dvi)
    
    ' 戻り値が0 (S_OK) であれば成功
    If ret = 0 Then
        MsgBox "Shell32.dll のバージョン情報:" & vbCrLf & _
               "Major Version: " & dvi.dwMajorVersion & vbCrLf & _
               "Minor Version: " & dvi.dwMinorVersion & vbCrLf & _
               "Build Number: " & dvi.dwBuildNumber
    Else
        MsgBox "DllGetVersion の呼び出しに失敗しました。戻り値: " & ret
    End If
End Sub

Sub GetCommCtrlVersion()
    Dim dvi As DLLVERSIONINFO
    Dim ret As Long
    
    ' 構造体のサイズを設定する
    dvi.cbSize = Len(dvi)
    
    ' ComCtlDllGetVersionを呼び出す
    ret = ComCtlDllGetVersion(dvi)
    
    ' 戻り値が0 (S_OK) であれば成功
    If ret = 0 Then
        MsgBox "comctl32.dll のバージョン情報:" & vbCrLf & _
               "Major Version: " & dvi.dwMajorVersion & vbCrLf & _
               "Minor Version: " & dvi.dwMinorVersion & vbCrLf & _
               "Build Number: " & dvi.dwBuildNumber
    Else
        MsgBox "ComCtlDllGetVersion の呼び出しに失敗しました。戻り値: " & ret
    End If
End Sub

Sub Main()
    UserForm1.Show 'vbModeless
End Sub
