Attribute VB_Name = "Util"
Option Explicit
Option Private Module

'ScriptingFileSystemObject
Public Enum FSO_OpenMode
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Private Type CONV_INT
    i As Integer
End Type

Private Type CONV_SINGLE
    s As Single
End Type

Private Type CONV_LONG
    l As Long
End Type

Private Type BYTE_ARRAY2
    b(1) As Byte
End Type

Private Type BYTE_ARRAY4
    b(3) As Byte
End Type

Private Type BYTE_ARRAY9
    b(8) As Byte
End Type

' API の定義
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Public Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
Public Declare PtrSafe Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As LongPtr, ByVal pszPath As String, ByVal Dx As Long) As Long
Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Public Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long
Public Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr

Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'Private Declare PtrSafe Sub ReleaseCapture Lib "user32" ()
'Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#If Win64 Then
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'Windows API定数
Public Const MF_BYCOMMAND      As Long = &H0&          '定数の設定

' System Menu Command Values
Const SC_SIZE = &HF000&
Const SC_MOVE = &HF010&
Const SC_MINIMIZE = &HF020&
Const SC_MAXIMIZE = &HF030&
Const SC_NEXTWINDOW = &HF040&
Const SC_PREVWINDOW = &HF050&
Const SC_CLOSE = &HF060&
Const SC_VSCROLL = &HF070&
Const SC_HSCROLL = &HF080&
Const SC_MOUSEMENU = &HF090&
Const SC_KEYMENU = &HF100&
Const SC_ARRANGE = &HF110&
Const SC_RESTORE = &HF120&
Const SC_TASKLIST = &HF130&
Const SC_SCREENSAVE = &HF140&
Const SC_HOTKEY = &HF150&

' Window field offsets for GetWindowLong() and GetWindowWord()
Const GWL_WNDPROC = (-4&)
Const GWL_HINSTANCE = (-6&)
Const GWL_HWNDPARENT = (-8&)
Const GWL_STYLE = (-16&)
Const GWL_EXSTYLE = (-20&)
Const GWL_USERDATA = (-21&)
Const GWL_ID = (-12&)

' Window Styles
Const WS_OVERLAPPED = &H0&
Const WS_POPUP = &H80000000
Const WS_CHILD = &H40000000
Const WS_MINIMIZE = &H20000000
Const WS_VISIBLE = &H10000000
Const WS_DISABLED = &H8000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_MAXIMIZE = &H1000000
Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_GROUP = &H20000
Const WS_TABSTOP = &H10000

Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000

Const WS_TILED = WS_OVERLAPPED
Const WS_ICONIC = WS_MINIMIZE
Const WS_SIZEBOX = WS_THICKFRAME
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

' Extended Window Styles
Const WS_EX_DLGMODALFRAME = &H1&
Const WS_EX_NOPARENTNOTIFY = &H4&
Const WS_EX_TOPMOST = &H8&
Const WS_EX_ACCEPTFILES = &H10&
Const WS_EX_TRANSPARENT = &H20&

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Window Messages
Public Const WM_COMMAND = &H111

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type POINTF
    X As Single
    Y As Single
End Type

Public Enum ControlCorner
    TopLeft = 0
    TopRight = 1
    BottomLeft = 2
    BottomRight = 3
End Enum

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, _
        ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, _
        ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
        ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
    
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
'Private Declare PtrSafe Function EnumDisplayMonitors Lib "user32" (ByVal hdc As LongPtr, ByRef lprcClip As Any, ByVal lpfnEnum As LongPtr, ByVal dwData As LongPtr) As Boolean
'Private Declare PtrSafe Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpmi As MonitorInfoEx) As Boolean

''support finding the primary monitor
'Public Const CCHDEVICENAME = 32
'Public Const MONITORINFOF_PRIMARY = &H1
'
'Public Type MonitorInfoEx
'    cbSize As Long
'    rcMonitor As RECT
'    rcWork As RECT
'    dwFlags As Long
'    szDevice As String * CCHDEVICENAME
'End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400

'Private MonitorId() As String

'GetDeviceCapsのnIndex設定値
Private Const LOGPIXELSX As Long = 88   ' Logical pixels/inch in X
Private Const LOGPIXELSY As Long = 90   ' Logical pixels/inch in Y
Private Const HORZRES As Long = 8       ' Horizontal width in pixels
Private Const VERTRES As Long = 10      ' Vertical width in pixels

'ポイントをピクセルに変換
Public Function PointToPixcel(Point As POINTF) As POINTAPI
    PointToPixcel.X = (Point.X / 72 * LogicalPixcel.X)
    PointToPixcel.Y = (Point.Y / 72 * LogicalPixcel.Y)
End Function

'ピクセルをポイントに変換
Public Function PixcelToPoint(Pixcel As POINTAPI) As POINTF
    PixcelToPoint.X = (Pixcel.X * 72 / LogicalPixcel.X)
    PixcelToPoint.Y = (Pixcel.Y * 72 / LogicalPixcel.Y)
End Function

'DPIを取得：ディスプレイの拡大率込
Public Function LogicalPixcel() As POINTAPI
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)

    LogicalPixcel.X = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    LogicalPixcel.Y = GetDeviceCaps(hDCDesk, LOGPIXELSY)

    Call ReleaseDC(hWndDesk, hDCDesk)
End Function

Public Function GetDisplaySize() As POINTAPI
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)

    GetDisplaySize.X = GetDeviceCaps(hDCDesk, HORZRES)
    GetDisplaySize.Y = GetDeviceCaps(hDCDesk, VERTRES)

    Call ReleaseDC(hWndDesk, hDCDesk)
End Function

Function GetControlPosition(ctrl As MSForms.Control, Corner As ControlCorner) As POINTF
    Dim parent As Object
    Set parent = ctrl.parent
    
    Dim clientCoordinate As POINTF
    clientCoordinate.X = ctrl.Left + (Corner Mod 2) * ctrl.Width
    clientCoordinate.Y = ctrl.Top + (Corner \ 2) * ctrl.height
    
    On Error GoTo ERROR_HANDLER:
    Dim hWnd As LongPtr
    Call WindowFromAccessibleObject(parent, hWnd)
    
    Dim screenCoordinate As POINTAPI
    screenCoordinate = PointToPixcel(clientCoordinate)
    
    Call ClientToScreen(hWnd, screenCoordinate)
    
    GetControlPosition = PixcelToPoint(screenCoordinate)
    
    Exit Function
ERROR_HANDLER:
    MsgBox "エラー:" & ERR.Number & " : " & ERR.Description
End Function

Public Function MeasureTextSize( _
        target_text As String, _
        FONT_NAME As String, _
        Optional font_height As Long = 10) As POINTF
    
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)
    
'    Dim hWholeScreenDC As LongPtr: hWholeScreenDC _
'        = GetDC(0&)
    
'    Dim hVirtualDC As LongPtr: hVirtualDC _
        = CreateCompatibleDC(hWholeScreenDC)
    Dim hVirtualDC As LongPtr: hVirtualDC _
        = CreateCompatibleDC(hDCDesk)

    Dim hFont As LongPtr: hFont _
        = CreateFont(font_height, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, FONT_NAME)
            
    Call SelectObject(hVirtualDC, hFont)
    
    Dim DrawAreaRectangle As RECT
    Call DrawText(hVirtualDC, target_text, -1, DrawAreaRectangle, DT_CALCRECT)
    
    Call DeleteObject(hFont)
    Call DeleteObject(hVirtualDC)
'    Call ReleaseDC(0&, hWholeScreenDC)
    Call ReleaseDC(hWndDesk, hDCDesk)
'    MeasureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left
    Dim TextSizePixel As POINTAPI
    TextSizePixel.X = DrawAreaRectangle.Right - DrawAreaRectangle.Left
    TextSizePixel.Y = DrawAreaRectangle.Bottom - DrawAreaRectangle.Top
'    MeasureTextSize = PixcelToPoint(TextSizePixel)
    MeasureTextSize.X = TextSizePixel.X
    MeasureTextSize.Y = TextSizePixel.Y
End Function

Public Function TableSearch(ByVal Table As ListObject, ByVal KeyColumn As String, ByVal SearchKey As Variant, ByVal ColumnName As String)
    With Table.ListColumns(KeyColumn)
        Dim i As Long
        For i = 1 To .DataBodyRange.count
            If .DataBodyRange(i).value = SearchKey Then
                Exit For
            End If
        Next
    End With
    
    TableSearch = Table.ListColumns(ColumnName).DataBodyRange(i).value
End Function

''フォーム再描写
'Public Sub Redraw()
'    If Form Is Nothing Then Exit Sub
'
'    Style_ = GetWindowLongPtr(hWnd_, GWL_STYLE)
'
'    If Minimize Then Style_ = Style_ Or WS_MINIMIZEBOX
'    If Maximize Then Style_ = Style_ Or WS_MAXIMIZEBOX
'    If Resize Then Style_ = Style_ Or WS_THICKFRAME
'
'    Call SetWindowLongPtr(hWnd_, GWL_STYLE, Style_)
'
'    If Not CloseButton Then
'        Dim hMenu_ As LongPtr
'        hMenu_ = GetSystemMenu(hWnd_, 0&)
'        Call DeleteMenu(hMenu_, SC_CLOSE, MF_BYCOMMAND)
'    End If
'
'    Call DrawMenuBar(hWnd_)
'
'    If Menu Then
'        MenuHandle = MenuInitialize
'        Call SetMenu(hWnd_, MenuHandle)
'    End If
'End Sub

Public Function FormNonCaption(ByVal UserForm As Object, Optional ByVal Flat As Boolean) As LongPtr
    Dim hWnd As LongPtr
    Dim ih As Single, iw As Single
    ih = UserForm.InsideHeight
    iw = UserForm.InsideWidth
    
    Call WindowFromAccessibleObject(UserForm, hWnd)
    If Flat Then Call SetWindowLongPtr(hWnd, GWL_EXSTYLE, GetWindowLongPtr(hWnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME)
    
    FormNonCaption = SetWindowLongPtr(hWnd, GWL_STYLE, GetWindowLongPtr(hWnd, GWL_STYLE) And Not WS_CAPTION)
    Call DrawMenuBar(hWnd)
    
    UserForm.height = ih
    UserForm.Width = iw
End Function

' 与えられたパス文字列からフルパス文字列を返す
' (Root指定のない場合は本WorkbookのパスをRootとする)
Public Function GetFullPath(ByVal filePath As String) As String
    GetFullPath = filePath
    If GetFolderName(filePath) = "" Or GetDriveName(filePath) = "" Then
        With CreateObject("Scripting.FileSystemObject")
            GetFullPath = .BuildPath(ThisWorkbook.Path, filePath)
        End With
    End If
End Function

' ファイルの存在確認(ワイルドカード可)
Public Function FileExists(ByVal filePath As String, Optional ByRef FoundFile As String) As Boolean
    FileExists = False
    
    If filePath = "" Then Exit Function
    
    Dim parentFolderName As String
    Dim fileName_ As String
    
    With VBA.CreateObject("Scripting.FileSystemObject")
        parentFolderName = .GetParentFolderName(filePath)
        fileName_ = .GetFileName(filePath)
        Dim file_ As Object
        For Each file_ In .GetFolder(parentFolderName).Files
            If file_.Name Like fileName_ Then
                FileExists = True
                'パターン一致したもの
                If Not VBA.IsMissing(FoundFile) Then FoundFile = parentFolderName & "\" & file_.Name
                Exit For
            End If
        Next file_
        Set file_ = Nothing
    End With
End Function

' フォルダの存在確認(ワイルドカード可)
Public Function FolderExists(ByVal filePath As String, Optional ByRef FoundFolder As String) As Boolean
    FolderExists = False
    If filePath = "" Then Exit Function
    
    Dim parentFolderName As String
    Dim folderName_ As String
    
    With VBA.CreateObject("Scripting.FileSystemObject")
        parentFolderName = .GetParentFolderName(filePath)
        folderName_ = .GetFileName(filePath)
    
        Dim folder_ As Object
        For Each folder_ In .GetFolder(parentFolderName).SubFolders
            If folder_.Name Like folderName_ Then
                FolderExists = True
                'パターン一致したもの
                If Not VBA.IsMissing(FoundFolder) Then FoundFolder = parentFolderName & "\" & folder_.Name
                Exit For
            End If
        Next folder_
        Set folder_ = Nothing
    End With
End Function

Public Function CleanString(aInput As String) As String
    CleanString = aInput
    
    Dim ZeroPos As Long
    ZeroPos = VBA.InStr(1, aInput, VBA.Chr$(0))
    If ZeroPos > 0 Then CleanString = VBA.Left$(aInput, ZeroPos - 1)
End Function

' 与えられたパス文字列から指定文字数以下での短縮形式のパスを取得
Public Function GetCompactPath(ByVal aPath As String, ByVal aLength As Long) As String
    GetCompactPath = ""
    
    Dim strTmp As String
    strTmp = VBA.String(255, 0)
    If PathCompactPathEx(strTmp, aPath, aLength, 0) <> 0 Then GetCompactPath = CleanString(strTmp)
End Function

'' 与えられたパス文字列から指定ピクセル数以下での短縮形式のパスを取得
'Public Function GetCompactPathPixel(ByVal aPath As String, ByVal Dx As Long, Optional ByVal aHDC As LongPtr = 0) As String
'    GetCompactPathPixel = ""
'
'    Dim strTmp As String
'    strTmp = aPath
'    If PathCompactPath(aHDC, strTmp, Dx) <> 0 Then GetCompactPathPixel = CleanString(strTmp)
'End Function

' 指定した名前のワークシートの存在を確認
Public Function ExistsWorksheet(ByVal SheetName As String)
    Dim WS As Worksheet
    
    For Each WS In Sheets
        If WS.Name = SheetName Then
            ' 存在する
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    
    ' 存在しない
    ExistsWorksheet = False
End Function

Public Sub wait(ByVal Milisecond As Long)
    Dim T As Long
    T = timeGetTime + Milisecond
    While T > timeGetTime
        DoEvents
    Wend
End Sub

Public Sub MyWait(ByVal milliSec As Long)
    Dim startTime As Single
 
    startTime = Timer
    Do While Timer < startTime + milliSec / 1000
        DoEvents
    Loop
End Sub

Public Static Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function

Public Function Byte2Int(ByRef b() As Byte, ByVal Index As Long, Optional ByVal size As Integer = 2) As Integer
    Dim byteArray As BYTE_ARRAY2
    Dim i As Integer
    
    For i = 0 To 1
        byteArray.b(i) = 0
    Next i
    
    For i = 0 To size - 1
        byteArray.b(i) = b(Index + i)
    Next i
    
    Dim ConvInt As CONV_INT
    LSet ConvInt = byteArray
    Byte2Int = ConvInt.i
End Function

Public Function Byte2Long(ByRef b() As Byte, ByVal Index As Long, Optional ByVal size As Integer = 4) As Long
    Dim byteArray As BYTE_ARRAY4
    Dim i As Integer
    
    For i = 0 To 3
        byteArray.b(i) = 0
    Next i
    
    For i = 0 To size - 1
        byteArray.b(i) = b(Index + i)
    Next i
    
    Dim ConvLong As CONV_LONG
    LSet ConvLong = byteArray
    Byte2Long = ConvLong.l
End Function

Public Function Byte2Single(ByRef b() As Byte, ByVal Index As Long, Optional ByVal size As Integer = 4) As Variant
    Dim byteArray As BYTE_ARRAY4
    Dim i As Integer
    
    For i = 0 To 3
        byteArray.b(i) = 0
    Next i
    
    For i = 0 To size - 1
        byteArray.b(i) = b(Index + i)
    Next i
    
    Dim ConvSingle As CONV_SINGLE
    LSet ConvSingle = byteArray
    Byte2Single = CDec(ConvSingle.s)
End Function

' 符号あり10進数→2進数文字列 (16bit)
Public Function Dec2Bin(ByVal argDec As Integer) As String
    Dim Binary As String
    Binary = ""
    
    Dim i As Integer
    For i = 0 To 15
        If BitTest(argDec, i) Then
            Binary = "1" & Binary
        Else
            Binary = "0" & Binary
        End If
    Next i
    
    Dec2Bin = Binary
End Function

' ビットテスト (16bitサイズ,LSB:bit0 とする)
Public Function FlagCount(ByVal argBitField As Integer) As Integer
    Dim calcTmp As Long
    calcTmp = CLng(argBitField + 2 ^ 16)
    
    Dim count As Integer
    Dim i As Integer
    For i = 0 To 15
        If (calcTmp \ (2 ^ i)) And 1 Then count = count + 1
    Next i
    
    FlagCount = count
End Function

' ビットテスト (16bitサイズ,LSB:bit0 とする)
Public Function BitTest(ByVal argBitField As Integer, ByVal argBit As Integer) As Boolean
    Dim calcTmp As Long
    calcTmp = CLng(argBitField + 2 ^ 16)
    
    BitTest = (calcTmp \ (2 ^ argBit)) And 1
End Function

' 文字列を1つ以上のヌル文字(0x00)で切り分ける
Public Function SplitByNull(ByVal Str As String) As String()
    Dim strTmp As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\x00+"
        .ignorecase = False '大文字と小文字の区別
        .Global = True      '文字列全体の検索
        strTmp = .Replace(Str, ",")
    End With

    SplitByNull = Split(strTmp, ",")
End Function

' 文字列を指定の文字リストで切り分ける(連続する場合は一つの区切りとする)
Public Function SplitByChar(ByVal Str As String, ByRef delim() As Variant) As String()
    Dim patternString As String
    patternString = Join(delim, "|")
    patternString = "(" & patternString & ")+"
    
    Dim strTmp As String
    With CreateObject("VBScript.RegExp")
        .Pattern = patternString
        .ignorecase = False '大文字と小文字の区別
        .Global = True      '文字列全体の検索
        strTmp = .Replace(Str, ",")
        ' 先頭と末尾の","を削除
        .Pattern = "(^,|,$)"
        strTmp = .Replace(strTmp, "")
    End With
        
    SplitByChar = Split(strTmp, ",")
End Function

' 文字列から指定の文字リストを削除する
Public Function ReplaceChar(ByVal Str As String, ByRef repChar() As Variant) As String
    Dim patternString As String
    patternString = Join(repChar, "|")
    patternString = "(" & patternString & ")+"
    
    With CreateObject("VBScript.RegExp")
        .Pattern = patternString
        .ignorecase = False '大文字と小文字の区別
        .Global = True      '文字列全体の検索
        ReplaceChar = .Replace(Str, "")
    End With
End Function

' 文字列から数値のみを取り出す
Public Function ExtractNumber(ByVal Str As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\D"
        .Global = True
        ExtractNumber = .Replace(Str, "")
    End With
End Function

' 文字列から数値以外を取り出す
Public Function ExtractWithoutNumber(ByVal Str As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\d"
        .Global = True
        ExtractWithoutNumber = .Replace(Str, "")
    End With
End Function

Public Function MakeZip(ByVal SrcPath As String, argZipPath As String) As Boolean
    Dim sh      As Object
    Dim ex      As Object
    Dim commandLine    As String
    
    Set sh = CreateObject("Script.Shell")
    
    SrcPath = Replace(SrcPath, " ", "` ")
    argZipPath = Replace(argZipPath, " ", "` ")
    
    '// Compress-Archive：圧縮コマンド
    '// -Path：フォルダパスまたはファイルパスを指定する。
    '// -DestinationPath：生成ファイルパスを指定する。
    '// -Force：生成ファイルが既に存在している場合は上書きする
    commandLine = "Compress-Archive -Path " & SrcPath & " -DestinationPath " & argZipPath & " -Force"
    
    Set ex = sh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & commandLine)
    
    If ex.Status = 2 Then
        MakeZip = False
        Exit Function
    End If
    
    Do While ex.Status = 0
        DoEvents
    Loop
    
    MakeZip = True
End Function

Public Function UnZip(ByVal SrcZipPath As String, argExpandPath As String) As Boolean
    Dim sh  As Object
    Dim ex  As Object
    Dim commandLine As String
    
    Set sh = CreateObject("Script.Shell")
    
    SrcZipPath = Replace(SrcZipPath, " ", "` ")
    argExpandPath = Replace(argExpandPath, " ", "` ")
    
    '// Expand-Archive：解凍コマンド
    '// -Path：フォルダパスまたはファイルパスを指定する。
    '// -DestinationPath：生成ファイルパスを指定する。
    '// -Force：生成ファイルが既に存在している場合は上書きする
    commandLine = "Expand-Archive -Path " & SrcZipPath & " -DestinationPath " & argExpandPath & " -Force"
    
    Set ex = sh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & commandLine)
    
    If ex.Status = 2 Then
        UnZip = False
        Exit Function
    End If
    
    Do While ex.Status = 0
        DoEvents
    Loop
    
    UnZip = True
End Function

' 与えられたパス文字列のファイル拡張子を指定の文字列に置き換えたファイル名を返す
Public Function ChangeExtension(ByVal aFilePath As String, ByVal aExtension As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        ChangeExtension = .GetParentFolderName(aFilePath) & "\" & _
                          .GetBaseName(aFilePath) & "." & aExtension
    End With
End Function

' 与えられたパス文字列から拡張子無しのベースファイル名を返す
Public Function GetBaseFileName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetBaseFileName = .GetBaseName(aFilePath)
    End With
End Function

' 与えられたパス文字列からファイル名を返す
Public Function GetFileName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetFileName = .GetFileName(aFilePath)
    End With
End Function

' 与えられたパス文字列から拡張子を返す
Public Function GetExtension(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetExtension = .GetExtensionName(aFilePath)
    End With
End Function

' 与えられたパス文字列からフォルダ名を返す
Public Function GetFolderName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetFolderName = .GetParentFolderName(aFilePath)
    End With
End Function

' 与えられたパス文字列からドライブ名を返す
Public Function GetDriveName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetDriveName = .GetDriveName(aFilePath)
    End With
End Function

