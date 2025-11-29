Attribute VB_Name = "Module1"
Option Explicit

' Private Window Messages Start Here:
Const WM_USER = &H400

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

Type CHOOSE_COLORA
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Declare PtrSafe Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As CHOOSE_COLORA) As Long

Const CC_RGBINIT = &H1
Const CC_FULLOPEN = &H2
Const CC_PREVENTFULLOPEN = &H4
Const CC_SHOWHELP = &H8
Const CC_ENABLEHOOK = &H10
Const CC_ENABLETEMPLATE = &H20
Const CC_ENABLETEMPLATEHANDLE = &H40
Const CC_SOLIDCOLOR = &H80
Const CC_ANYCOLOR = &H100

Type CHOOSE_FONTA
    lStructSize As Long
    hwndOwner As LongPtr       '  caller's window handle
    hdc As LongPtr             '  printer DC/IC or NULL
    lpLogFont As LongPtr       '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As LongPtr       '  data passed to hook fn.
    lpfnHook As LongPtr        '  ptr. to hook function
    lpTemplateName As String   '  custom template name
    hInstance As LongPtr       '  instance handle of.EXE that
                               '    contains cust. dlg. template
    lpszStyle As String        '  return the style field here
                               '  must be LF_FACESIZE or bigger
    nFontType As Integer       '  same value reported to the EnumFonts
                               '    call back with the extra FONTTYPE_
                               '    bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
                               '    CF_LIMITSIZE is used
End Type

Declare PtrSafe Function ChooseFontA Lib "comdlg32.dll" (pChoosefont As CHOOSE_FONTA) As Long

Public Const CF_SCREENFONTS = &H1
Const CF_PRINTERFONTS = &H2
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_SHOWHELP = &H4&
Const CF_ENABLEHOOK = &H8&
Const CF_ENABLETEMPLATE = &H10&
Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Const CF_APPLY = &H200&
Const CF_ANSIONLY = &H400&
Const CF_SCRIPTSONLY = CF_ANSIONLY
Const CF_NOVECTORFONTS = &H800&
Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Const CF_NOSIMULATIONS = &H1000&
Const CF_LIMITSIZE = &H2000&
Const CF_FIXEDPITCHONLY = &H4000&
Const CF_WYSIWYG = &H8000& '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Const CF_FORCEFONTEXIST = &H10000
Const CF_SCALABLEONLY = &H20000
Const CF_TTONLY = &H40000
Const CF_NOFACESEL = &H80000
Const CF_NOSTYLESEL = &H100000
Const CF_NOSIZESEL = &H200000
Const CF_SELECTSCRIPT = &H400000
Const CF_NOSCRIPTSEL = &H800000
Const CF_NOVERTFONTS = &H1000000

Const SIMULATED_FONTTYPE = &H8000
Const PRINTER_FONTTYPE = &H4000
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)

Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONTA) As LongPtr

' Logical Font
Const LF_FACESIZE = 32
Const LF_FULLFACESIZE = 64

Type LOGFONTA
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
'    lfFaceName(1 To LF_FACESIZE) As Byte
    lfFaceName As String * 32
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

Function ChooseColorDialog(ByRef lColor As Long, Optional ByVal hWnd As LongPtr = 0) As Boolean
    Static lCustColors(0 To 15) As Long 'ｶｽﾀﾑｶﾗｰを保持するため静的変数
    Static flgInit As Boolean
    
    Dim tChooseColor As CHOOSE_COLORA
    Dim lFlg As Long
    Dim i As Long
    
    '初期化ﾌﾗｸﾞの設定(複合の場合は[Or])
    lFlg = CC_RGBINIT Or CC_FULLOPEN
    
    '初回起動時のみ[作成した色]をすべて白色にｾｯﾄ
    If flgInit = False Then
        For i = 0 To UBound(lCustColors)
            lCustColors(i) = RGB(255, 255, 255)
        Next
        flgInit = True
    End If
    
    'ﾀﾞｲｱﾛｸﾞ初期設定
    With tChooseColor
        .lStructSize = LenB(tChooseColor)       '構造体のｻｲｽﾞ
        .hwndOwner = hWnd                       'ｵｰﾅｰ(親)ｳｨﾝﾄﾞｳのﾊﾝﾄﾞﾙ
        .rgbResult = lColor                     '初期選択ｶﾗｰ(&ﾕｰｻﾞｰ選択ｶﾗｰの受け取り)
        .lpCustColors = VarPtr(lCustColors(0))  '[作成した色]のｶﾗｰｺｰﾄﾞを格納した配列ﾎﾟｲﾝﾀ
        .flags = lFlg                           'ﾀﾞｲｱﾛｸﾞの初期化ﾌﾗｸﾞ
    End With
    If ChooseColorA(tChooseColor) = 0 Then
        ChooseColorDialog = False
    Else
        lColor = tChooseColor.rgbResult
        ChooseColorDialog = True
    End If
End Function

Function ChooseFontDialog(ByRef cf As CHOOSE_FONTA) As Boolean
    ChooseFontDialog = False
    Debug.Print "ChooseFontDialog"
    If Not ChooseFontA(cf) Then Exit Function
    ChooseFontDialog = True
End Function

Sub Main()
    UserForm1.Show 'vbModeless
End Sub
