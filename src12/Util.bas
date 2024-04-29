Attribute VB_Name = "Util"
Option Explicit
Option Private Module

'ScriptingFileSystemObject
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8

Enum ErrCode
    'No.0�`512:System Reserved
    'No.513�`514:�N���X�^���R���X�g���N�^�Ăяo���p
    InitTwice = 513
    NotReady = 514
    'No.515�`5--:�X�N���v�g�֘A
    ScriptRunTime = 515
    VarNotExists = 516
    NumCannotConvert = 517
    ScriptFileNotExists = 518
    ScriptSyntax = 519
End Enum

Private Type CONV_INT
    i As Integer
End Type

Private Type CONV_SINGLE
    S As Single
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

' API �̒�`
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Public Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
Public Declare PtrSafe Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As LongPtr, ByVal pszPath As String, ByVal Dx As Long) As Long
Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal pUnk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Public Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long
'Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As Object, phwnd As Long) As Long
Public Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr

Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long

'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare PtrSafe Sub ReleaseCapture Lib "user32" ()
'Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#If Win64 Then
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'Windows API�萔
Public Const MF_BYCOMMAND      As Long = &H0&          '�萔�̐ݒ�

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
    x As Long
    y As Long
End Type

'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
'Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hgdiobj As Long) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
'Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

'Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
    ByVal nWidth As Long, _
    ByVal nEscapement As Long, _
    ByVal nOrientation As Long, _
    ByVal fnWeight As Long, _
    ByVal IfdwItalic As Long, _
    ByVal fdwUnderline As Long, _
    ByVal fdwStrikeOut As Long, _
    ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, _
    ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, _
        ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal IfdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, _
        ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
        ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
    
'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400

'GetDeviceCaps��nIndex�ݒ�l
Private Const LOGPIXELSX As Long = 88   ' Logical pixels/inch in X
Private Const LOGPIXELSY As Long = 90   ' Logical pixels/inch in Y

'�|�C���g���s�N�Z���ɕϊ�
Public Function PointToPixcel(ByVal aPoint As Single) As Long
    PointToPixcel = aPoint / 72 * LogicalPixcel
End Function

'�s�N�Z�����|�C���g�ɕϊ�
Public Function PixcelToPoint(ByVal aPixcel As Long) As Single
    PixcelToPoint = aPixcel * 72 / LogicalPixcel
End Function

'DPI���擾�F�f�B�X�v���C�̊g�嗦��
Public Function LogicalPixcel() As Long
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)

    LogicalPixcel = GetDeviceCaps(hDCDesk, LOGPIXELSX)

    Call ReleaseDC(hWndDesk, hDCDesk)
End Function

Function GetControlPosition(ctrl As MSForms.Control) As POINTAPI
    Dim parent As Object
    Set parent = ctrl.parent
    Dim clientX As Single, clientY As Single
    
    clientX = ctrl.Left
    clientY = ctrl.Top
    Debug.Print "1:(" & clientX & "," & clientY & ")"
    Do Until (TypeName(parent) = "MainForm")
'        Debug.Print "1:" & TypeName(parent)
        If TypeName(parent) = "Page" Then Set parent = parent.parent
'        Debug.Print "2:" & TypeName(parent)
        clientX = clientX + parent.Left
        clientY = clientY + parent.Top
        Debug.Print "2:(" & clientX & "," & clientY & ")"
        Set parent = parent.parent
        Debug.Print "3:" & TypeName(parent)
    Loop
    
    Debug.Print "3:(" & clientX & "," & clientY & ")"

    On Error GoTo ERR:
    Dim hwnd As LongPtr
    Debug.Print "4:" & TypeName(parent)
'    Call WindowFromAccessibleObject(parent, hwnd)
    Debug.Print "4:(" & clientX & "," & clientY & ")"
    With GetControlPosition
'        .x = parent.Left + PointToPixcel(clientY)
'        .y = parent.Top + PointToPixcel(clientY)
        .x = PointToPixcel(clientY)
        .y = PointToPixcel(clientY)
        Debug.Print "(" & .x & "," & .y & ")"
    End With
    
'    GetControlPosition.x = GetControlPosition.x
'    GetControlPosition.y = GetControlPosition.y
'    Call ClientToScreen(hwnd, GetControlPosition)
    Exit Function
ERR:
    MsgBox "ERR"
End Function

Function MeasureTextWidth( _
        target_text As String, _
        FONT_NAME As String, _
        Optional font_height As Long = 10) As Long
    
    Dim hWholeScreenDC As Long: hWholeScreenDC _
        = GetDC(0&)
    
    Dim hVirtualDC As Long: hVirtualDC _
        = CreateCompatibleDC(hWholeScreenDC)

    Dim hFont As Long: hFont _
        = CreateFont(font_height, 0, 0, 0, FW_NORMAL, _
            0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
            CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
            DEFAULT_PITCH Or FF_SCRIPT, FONT_NAME)
            
    Call SelectObject(hVirtualDC, hFont)
    
    Dim DrawAreaRectangle As RECT
    Call DrawText(hVirtualDC, target_text, -1, DrawAreaRectangle, DT_CALCRECT)
    
    Call DeleteObject(hFont)
    Call DeleteObject(hVirtualDC)
    Call ReleaseDC(0&, hWholeScreenDC)
    MeasureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left
End Function

'Public Const HTCAPTION = 2
'Public Const WM_NCLBUTTONDOWN = &HA1

'Public Sub moverForm(Form As Object, obj As Object, Button As Integer)
'    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
'    If Val(Application.Version) < 9 Then
'        lngMyHandle = FindWindow("ThunderXFrame", Form.Caption)
'    Else
'        lngMyHandle = FindWindow("ThunderDFrame", Form.Caption)
'    End If
'
'    If Button = 1 Then
'        With obj
'            Call ReleaseCapture
'            Call SendMessage(lngMyHandle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'        End With
'    End If
'End Sub

Sub Autofit_Listbox1()
    Dim WS As Worksheet
    Dim LS, LastColumn, i As Long
    Dim objek As String
    
    Set WS = ThisWorkbook.Sheets("Data")
    LS = WS.Range("A" & Rows.count).End(xlUp).Row
    objek = "userform1.listbox1_"
    WS.Cells.EntireColumn.AutoFit
    
        With UserForm1.ListBox1
            .ColumnCount = 13
            .ColumnWidths = ""
            For i = 1 To 15
                If i > 20 Then
                    .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & 0
                Else
                    .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & WS.Cells(1, i).Width
                End If
            Next i
        End With
    
    UserForm1.ListBox1.RowSource = WS.Range("A2:E" & LS + 1).Address(External:=True)
End Sub

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

'�t�H�[���ĕ`��
Public Sub Redraw()
    If form Is Nothing Then Exit Sub
    
    Style_ = GetWindowLongPtr(hWnd_, GWL_STYLE)
    
    If Minimize Then Style_ = Style_ Or WS_MINIMIZEBOX
    If Maximize Then Style_ = Style_ Or WS_MAXIMIZEBOX
    If Resize Then Style_ = Style_ Or WS_THICKFRAME
    
    Call SetWindowLongPtr(hWnd_, GWL_STYLE, Style_)
    
    If Not CloseButton Then
        Dim hMenu_ As LongPtr
        hMenu_ = GetSystemMenu(hWnd_, 0&)
        Call DeleteMenu(hMenu_, SC_CLOSE, MF_BYCOMMAND)
    End If
    
    Call DrawMenuBar(hWnd_)

    If Menu Then
        MenuHandle = MenuInitialize
        Call SetMenu(hWnd_, MenuHandle)
    End If
End Sub

Public Function FormNonCaption(ByVal UserForm As Object, Optional ByVal Flat As Boolean) As LongPtr
    Dim hwnd As LongPtr
    Dim ih As Single, iw As Single
    ih = UserForm.InsideHeight
    iw = UserForm.InsideWidth
    Call WindowFromAccessibleObject(UserForm, hwnd)
    If Flat Then Call SetWindowLongPtr(hwnd, GWL_EXSTYLE, GetWindowLongPtr(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME)
    FormNonCaption = SetWindowLongPtr(hwnd, GWL_STYLE, GetWindowLongPtr(hwnd, GWL_STYLE) And Not WS_CAPTION)
    Call DrawMenuBar(hwnd)
    UserForm.Height = ih
    UserForm.Width = iw
End Function

' �^����ꂽ�p�X�����񂩂�t���p�X�������Ԃ�
' (Root�w��̂Ȃ��ꍇ�͖{Workbook�̃p�X��Root�Ƃ���)
Public Function GetFullPath(ByVal filePath As String) As String
    GetFullPath = filePath
    If GetFolderString(filePath) = "" Or GetDriveString(filePath) = "" Then
        With CreateObject("Scripting.FileSystemObject")
            GetFullPath = .BuildPath(ThisWorkbook.Path, filePath)
        End With
    End If
End Function

' �t�@�C���̑��݊m�F(���C���h�J�[�h��)
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
                '�p�^�[����v��������
                If Not VBA.IsMissing(FoundFile) Then FoundFile = parentFolderName & "\" & file_.Name
                Exit For
            End If
        Next file_
        Set file_ = Nothing
    End With
End Function

' �t�H���_�̑��݊m�F(���C���h�J�[�h��)
Public Function FolderExists(ByVal filePath As String, Optional ByRef FoundFolder As String) As Boolean
    FolderExists = False
    If filePath = "" Then Exit Function
    
    Dim parentFolderName As String
    Dim folderName_ As String
    
    With VBA.CreateObject("Scripting.FileSystemObject")
        parentFolderName = .GetParentFolderName(aFilePath)
        folderName_ = .GetFileName(aFilePath)
    
        Dim folder_ As Object
        For Each folder_ In .GetFolder(parentFolderName).SubFolders
            If folder_.Name Like folderName_ Then
                FolderExists = True
                '�p�^�[����v��������
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

' �^����ꂽ�p�X�����񂩂�w�蕶�����ȉ��ł̒Z�k�`���̃p�X���擾
Public Function GetCompactPath(ByVal aPath As String, ByVal aLength As Long) As String
    GetCompactPath = ""
    
    Dim strTmp As String
    strTmp = VBA.String(255, 0)
    If PathCompactPathEx(strTmp, aPath, aLength, 0) <> 0 Then GetCompactPath = CleanString(strTmp)
End Function

'' �^����ꂽ�p�X�����񂩂�w��s�N�Z�����ȉ��ł̒Z�k�`���̃p�X���擾
'Public Function GetCompactPathPixel(ByVal aPath As String, ByVal Dx As Long, Optional ByVal aHDC As LongPtr = 0) As String
'    GetCompactPathPixel = ""
'
'    Dim strTmp As String
'    strTmp = aPath
'    If PathCompactPath(aHDC, strTmp, Dx) <> 0 Then GetCompactPathPixel = CleanString(strTmp)
'End Function

' �w�肵�����O�̃��[�N�V�[�g�̑��݂��m�F
Public Function ExistsWorksheet(ByVal SheetName As String)
    Dim WS As Worksheet
    
    For Each WS In Sheets
        If WS.Name = SheetName Then
            ' ���݂���
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    
    ' ���݂��Ȃ�
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

Public Static Function Log10(x)
    Log10 = Log(x) / Log(10#)
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
    Byte2Single = CDec(ConvSingle.S)
End Function

' ��������10�i����2�i�������� (16bit)
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

' �r�b�g�e�X�g (16bit�T�C�Y,LSB:bit0 �Ƃ���)
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

' �r�b�g�e�X�g (16bit�T�C�Y,LSB:bit0 �Ƃ���)
Public Function BitTest(ByVal argBitField As Integer, ByVal argBit As Integer) As Boolean
    Dim calcTmp As Long
    calcTmp = CLng(argBitField + 2 ^ 16)
    
    BitTest = (calcTmp \ (2 ^ argBit)) And 1
End Function

' �������1�ȏ�̃k������(0x00)�Ő؂蕪����
Public Function SplitByNull(ByVal Str As String) As String()
    With CreateObject("VBScript.RegExp")
        .Pattern = "\x00+"
        .ignorecase = False '�啶���Ə������̋��
        .Global = True      '������S�̂̌���
    End With

    Dim strTmp As String
    strTmp = reg.Replace(Str, ",")

    SplitByNull = Split(strTmp, ",")
End Function

' ��������w��̕������X�g�Ő؂蕪����(�A������ꍇ�͈�̋�؂�Ƃ���)
Public Function SplitByChar(ByVal Str As String, ByRef delim() As Variant) As String()
    Dim patternString As String
    patternString = Join(delim, "|")
    patternString = "(" & patternString & ")+"
    
    Dim strTmp As String
    With CreateObject("VBScript.RegExp")
        .Pattern = patternString
        .ignorecase = False '�啶���Ə������̋��
        .Global = True      '������S�̂̌���
        strTmp = .Replace(Str, ",")
        ' �擪�Ɩ�����","���폜
        .Pattern = "(^,|,$)"
        strTmp = .Replace(strTmp, "")
    End With
        
    SplitByChar = Split(strTmp, ",")
End Function

' �����񂩂�w��̕������X�g���폜����
Public Function ReplaceChar(ByVal Str As String, ByRef repChar() As Variant) As String
    Dim patternString As String
    patternString = Join(repChar, "|")
    patternString = "(" & patternString & ")+"
    
    With CreateObject("VBScript.RegExp")
        .Pattern = patternString
        .ignorecase = False '�啶���Ə������̋��
        .Global = True      '������S�̂̌���
        ReplaceChar = .Replace(Str, "")
    End With
End Function

' �����񂩂琔�l�݂̂����o��
Public Function ExtractNumber(ByVal Str As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\D"
        .Global = True
        ExtractNumber = .Replace(Str, "")
    End With
End Function

' �����񂩂琔�l�ȊO�����o��
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
    
    '// Compress-Archive�F���k�R�}���h
    '// -Path�F�t�H���_�p�X�܂��̓t�@�C���p�X���w�肷��B
    '// -DestinationPath�F�����t�@�C���p�X���w�肷��B
    '// -Force�F�����t�@�C�������ɑ��݂��Ă���ꍇ�͏㏑������
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
    
    '// Expand-Archive�F�𓀃R�}���h
    '// -Path�F�t�H���_�p�X�܂��̓t�@�C���p�X���w�肷��B
    '// -DestinationPath�F�����t�@�C���p�X���w�肷��B
    '// -Force�F�����t�@�C�������ɑ��݂��Ă���ꍇ�͏㏑������
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

' �^����ꂽ�p�X������̃t�@�C���g���q���w��̕�����ɒu���������t�@�C������Ԃ�
Public Function ChangeExtension(ByVal aFilePath As String, ByVal aExtension As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        ChangeExtension = .GetParentFolderName(aFilePath) & "\" & _
                          .GetBaseName(aFilePath) & "." & aExtension
    End With
End Function

' �^����ꂽ�p�X�����񂩂�g���q�����̃x�[�X�t�@�C������Ԃ�
Public Function GetBaseFileName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetBaseFileName = .GetBaseName(aFilePath)
    End With
End Function

' �^����ꂽ�p�X�����񂩂�t�@�C������Ԃ�
Public Function GetFileName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetFileName = .GetFileName(aFilePath)
    End With
End Function

' �^����ꂽ�p�X�����񂩂�g���q��Ԃ�
Public Function GetExtension(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetExtension = .GetExtensionName(aFilePath)
    End With
End Function

' �^����ꂽ�p�X�����񂩂�t�H���_����Ԃ�
Public Function GetFolderName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetFolderName = .GetParentFolderName(aFilePath)
    End With
End Function

' �^����ꂽ�p�X�����񂩂�h���C�u����Ԃ�
Public Function GetDriveName(ByVal aFilePath As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetDriveName = .GetDriveName(aFilePath)
    End With
End Function

