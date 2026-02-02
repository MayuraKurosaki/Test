Attribute VB_Name = "Module1"
Option Explicit

' Private Window Messages Start Here:
'Const WM_USER = &H400

'Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Declare PtrSafe Function DllGetVersion Lib "shell32.dll" (pdwVersion As DLLVERSIONINFO) As Long
Declare PtrSafe Function ComCtlDllGetVersion Lib "comctl32.dll" Alias "DllGetVersion" (pdwVersion As DLLVERSIONINFO) As Long

Declare PtrSafe Function FormatMessageW Lib "kernel32" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, Arguments As Long) As Long

'システム定義エラーのメッセージを取得する
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

'%1などの挿入シーケンスは無視され、変更されずに出力バッファーに渡される
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

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
    Flags As Long
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

'Type CHOOSE_FONTA
'    lStructSize As Long
'    hwndOwner As LongPtr       '  caller's window handle
'    hDC As LongPtr             '  printer DC/IC or NULL
'    lpLogFont As LongPtr       '  ptr. to a LOGFONT struct
'    iPointSize As Long         '  10 * size in points of selected font
'    Flags As Long              '  enum. type flags
'    rgbColors As Long          '  returned text color
'    lCustData As LongPtr       '  data passed to hook fn.
'    lpfnHook As LongPtr        '  ptr. to hook function
'    lpTemplateName As String   '  custom template name
'    hInstance As LongPtr       '  instance handle of.EXE that
'                               '    contains cust. dlg. template
'    lpszStyle As String        '  return the style field here
'                               '  must be LF_FACESIZE or bigger
'    nFontType As Integer       '  same value reported to the EnumFonts
'                               '    call back with the extra FONTTYPE_
'                               '    bits added
'    MISSING_ALIGNMENT As Integer
'    nSizeMin As Long           '  minimum pt size allowed &
'    nSizeMax As Long           '  max pt size allowed if
'                               '    CF_LIMITSIZE is used
'End Type
'typedef struct tagCHOOSEFONTA {
'   DWORD           lStructSize;
'   HWND            hwndOwner;          // caller's window handle
'   HDC             hDC;                // printer DC/IC or NULL
'   LPLOGFONTA      lpLogFont;          // ptr. to a LOGFONT struct
'   INT             iPointSize;         // 10 * size in points of selected font
'   DWORD           Flags;              // enum. type flags
'   COLORREF        rgbColors;          // returned text color
'   LPARAM          lCustData;          // data passed to hook fn.
'   LPCFHOOKPROC    lpfnHook;           // ptr. to hook function
'   LPCSTR          lpTemplateName;     // custom template name
'   HINSTANCE       hInstance;          // instance handle of.EXE that
'                                       //   contains cust. dlg. template
'   LPSTR           lpszStyle;          // return the style field here
'                                       // must be LF_FACESIZE or bigger
'   WORD            nFontType;          // same value reported to the EnumFonts
'                                       //   call back with the extra FONTTYPE_
'                                       //   bits added
'   WORD            ___MISSING_ALIGNMENT__;
'   INT             nSizeMin;           // minimum pt size allowed &
'   INT             nSizeMax;           // max pt size allowed if
'                                       //   CF_LIMITSIZE is used
'} CHOOSEFONTA;

'typedef struct tagCHOOSEFONTW {
'   DWORD           lStructSize;
'   HWND            hwndOwner;          // caller's window handle
'   HDC             hDC;                // printer DC/IC or NULL
'   LPLOGFONTW      lpLogFont;          // ptr. to a LOGFONT struct
'   INT             iPointSize;         // 10 * size in points of selected font
'   DWORD           Flags;              // enum. type flags
'   COLORREF        rgbColors;          // returned text color
'   LPARAM          lCustData;          // data passed to hook fn.
'   LPCFHOOKPROC    lpfnHook;           // ptr. to hook function
'   LPCWSTR         lpTemplateName;     // custom template name
'   HINSTANCE       hInstance;          // instance handle of.EXE that
'                                       //   contains cust. dlg. template
'   LPWSTR          lpszStyle;          // return the style field here
'                                       // must be LF_FACESIZE or bigger
'   WORD            nFontType;          // same value reported to the EnumFonts
'                                       //   call back with the extra FONTTYPE_
'                                       //   bits added
'   WORD            ___MISSING_ALIGNMENT__;
'   INT             nSizeMin;           // minimum pt size allowed &
'   INT             nSizeMax;           // max pt size allowed if
'                                       //   CF_LIMITSIZE is used
'} CHOOSEFONTW;

'Declare PtrSafe Function ChooseFontA Lib "comdlg32.dll" (pChoosefont As CHOOSE_FONTA) As Long

'Public Const CF_SCREENFONTS = &H1
'Const CF_PRINTERFONTS = &H2
'Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Const CF_SHOWHELP = &H4&
'Const CF_ENABLEHOOK = &H8&
'Const CF_ENABLETEMPLATE = &H10&
'Const CF_ENABLETEMPLATEHANDLE = &H20&
'Public Const CF_INITTOLOGFONTSTRUCT = &H40&
'Const CF_USESTYLE = &H80&
'Public Const CF_EFFECTS = &H100&
'Const CF_APPLY = &H200&
'Const CF_ANSIONLY = &H400&
'Const CF_SCRIPTSONLY = CF_ANSIONLY
'Const CF_NOVECTORFONTS = &H800&
'Const CF_NOOEMFONTS = CF_NOVECTORFONTS
'Const CF_NOSIMULATIONS = &H1000&
'Const CF_LIMITSIZE = &H2000&
'Const CF_FIXEDPITCHONLY = &H4000&
'Const CF_WYSIWYG = &H8000& '  must also have CF_SCREENFONTS CF_PRINTERFONTS
'Const CF_FORCEFONTEXIST = &H10000
'Const CF_SCALABLEONLY = &H20000
'Const CF_TTONLY = &H40000
'Const CF_NOFACESEL = &H80000
'Const CF_NOSTYLESEL = &H100000
'Const CF_NOSIZESEL = &H200000
'Const CF_SELECTSCRIPT = &H400000
'Const CF_NOSCRIPTSEL = &H800000
'Const CF_NOVERTFONTS = &H1000000
'
'Const SIMULATED_FONTTYPE = &H8000
'Const PRINTER_FONTTYPE = &H4000
'Const SCREEN_FONTTYPE = &H2000
'Const BOLD_FONTTYPE = &H100
'Const ITALIC_FONTTYPE = &H200
'Public Const REGULAR_FONTTYPE = &H400
'
'Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
'Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
'Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)

'Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
'Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONTA) As LongPtr
Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

'' Logical Font
'Const LF_FACESIZE = 32
'Const LF_FULLFACESIZE = 64

'Type LOGFONTA
'    lfHeight As Long
'    lfWidth As Long
'    lfEscapement As Long
'    lfOrientation As Long
'    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
''    lfFaceName(1 To LF_FACESIZE) As Byte
'    lfFaceName As String * 32
'End Type
'typedef struct tagLOGFONTA
'{
'    LONG      lfHeight;
'    LONG      lfWidth;
'    LONG      lfEscapement;
'    LONG      lfOrientation;
'    LONG      lfWeight;
'    BYTE      lfItalic;
'    BYTE      lfUnderline;
'    BYTE      lfStrikeOut;
'    BYTE      lfCharSet;
'    BYTE      lfOutPrecision;
'    BYTE      lfClipPrecision;
'    BYTE      lfQuality;
'    BYTE      lfPitchAndFamily;
'    CHAR      lfFaceName[LF_FACESIZE];
'} LOGFONTA, *PLOGFONTA, NEAR *NPLOGFONTA, FAR *LPLOGFONTA;

'typedef struct tagLOGFONTW
'{
'    LONG      lfHeight;
'    LONG      lfWidth;
'    LONG      lfEscapement;
'    LONG      lfOrientation;
'    LONG      lfWeight;
'    BYTE      lfItalic;
'    BYTE      lfUnderline;
'    BYTE      lfStrikeOut;
'    BYTE      lfCharSet;
'    BYTE      lfOutPrecision;
'    BYTE      lfClipPrecision;
'    BYTE      lfQuality;
'    BYTE      lfPitchAndFamily;
'    WCHAR     lfFaceName[LF_FACESIZE];
'} LOGFONTW, *PLOGFONTW, NEAR *NPLOGFONTW, FAR *LPLOGFONTW;

Sub GetShellVersion()
    Dim dvi As DLLVERSIONINFO
    Dim Ret As Long
    
    ' 構造体のサイズを設定する
    dvi.cbSize = Len(dvi)
    
    ' DllGetVersionを呼び出す
    Ret = DllGetVersion(dvi)
    
    ' 戻り値が0 (S_OK) であれば成功
    If Ret = 0 Then
        MsgBox "Shell32.dll のバージョン情報:" & vbCrLf & _
               "Major Version: " & dvi.dwMajorVersion & vbCrLf & _
               "Minor Version: " & dvi.dwMinorVersion & vbCrLf & _
               "Build Number: " & dvi.dwBuildNumber
    Else
        MsgBox "DllGetVersion の呼び出しに失敗しました。戻り値: " & Ret
    End If
End Sub

Sub GetCommCtrlVersion()
    Dim dvi As DLLVERSIONINFO
    Dim Ret As Long
    
    ' 構造体のサイズを設定する
    dvi.cbSize = Len(dvi)
    
    ' ComCtlDllGetVersionを呼び出す
    Ret = ComCtlDllGetVersion(dvi)
    
    ' 戻り値が0 (S_OK) であれば成功
    If Ret = 0 Then
        MsgBox "comctl32.dll のバージョン情報:" & vbCrLf & _
               "Major Version: " & dvi.dwMajorVersion & vbCrLf & _
               "Minor Version: " & dvi.dwMinorVersion & vbCrLf & _
               "Build Number: " & dvi.dwBuildNumber
    Else
        MsgBox "ComCtlDllGetVersion の呼び出しに失敗しました。戻り値: " & Ret
    End If
End Sub

Function ChooseColorDialog(ByRef lColor As Long, Optional ByVal hwnd As LongPtr = 0) As Boolean
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
        .hwndOwner = hwnd                       'ｵｰﾅｰ(親)ｳｨﾝﾄﾞｳのﾊﾝﾄﾞﾙ
        .rgbResult = lColor                     '初期選択ｶﾗｰ(&ﾕｰｻﾞｰ選択ｶﾗｰの受け取り)
        .lpCustColors = VarPtr(lCustColors(0))  '[作成した色]のｶﾗｰｺｰﾄﾞを格納した配列ﾎﾟｲﾝﾀ
        .Flags = lFlg                           'ﾀﾞｲｱﾛｸﾞの初期化ﾌﾗｸﾞ
    End With
    If ChooseColorA(tChooseColor) = 0 Then
        ChooseColorDialog = False
    Else
        lColor = tChooseColor.rgbResult
        ChooseColorDialog = True
    End If
End Function

'Function ChooseFontDialog(ByRef cf As CHOOSE_FONTA) As Boolean
'    ChooseFontDialog = False
'    Debug.Print "ChooseFontDialog"
'    If Not ChooseFontA(cf) Then Exit Function
'    ChooseFontDialog = True
'End Function

Public Function SelectFile() As String
    SelectFile = ""
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .InitialFileName = Path2 & "\"
        .AllowMultiSelect = False
        If .Show = True Then
            SelectFile = .SelectedItems(1)
        End If
    End With
End Function

Public Function GetAPIError(ByVal ErrorCode As Long) As String
    Dim errorBuffer As String
    errorBuffer = String$(256, vbNullChar)
    
    Dim result As Long
    result = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                            ByVal 0&, _
                            ErrorCode, _
                            0&, _
                            StrPtr(errorBuffer), _
                            Len(errorBuffer), _
                            0&)
    
    If result = 0 Then
        MsgBox "FormatMessageW Fail.", , "Error"
    Else
        Dim errorMessage As String
        errorMessage = Left$(errorBuffer, InStr(errorBuffer, vbNullChar) - 1)
        GetAPIError = errorMessage
    End If
End Function

Public Property Get Path2()
    Dim sPath As String
    sPath = ThisWorkbook.Path
    If Not sPath Like "http*" Then
        Path2 = sPath
        Exit Property
    End If

    Dim oneD As String: oneD = Environ("OneDrive")
    
    Dim sTemp As String
    sTemp = Replace(sPath, "/", "_", , 3)
    Path2 = oneD & "\" & Mid(sTemp, InStr(sTemp, "/") + 1)
End Property

Sub Main()
    UserForm1.Show 'vbModeless
End Sub

Public Function Redirect(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, _
                    ByVal lParam As LongPtr, ByVal id As Long, ByVal lv As ListView) As LongPtr
    Redirect = lv.WndProc(hwnd, uMsg, wParam, lParam)
End Function


