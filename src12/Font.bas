Attribute VB_Name = "Font"
Option Explicit

'Private Const GMEM_MOVEABLE = &H2
'Private Const GMEM_ZEROINIT = &H40
'Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)




' Private Window Messages Start Here:
Private Const WM_USER = &H400

'Type CHOOSE_FONTA
'    lStructSize As Long
'    hwndOwner As LongPtr       '  caller's window handle
'    hdc As LongPtr             '  printer DC/IC or NULL
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
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_OEMTEXT = 7
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS

Const SIMULATED_FONTTYPE = &H8000
Const PRINTER_FONTTYPE = &H4000
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)

'Public Const LOGPIXELSY = 90
' Logical Font
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Public Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
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
    lfFaceName(LF_FACESIZE) As Byte
End Type
Public Type LOGFONTA
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
    lfFaceName(1 To LF_FACESIZE) As Byte
'    lfFaceName As String * 32
End Type

'Private Type FONTSTRUC
'    lStructSize As Long
'    hwnd As LongPtr
'    hDC As LongPtr
'    lpLogFont As LongPtr
'    iPointSize As Long
'    Flags As Long
'    rgbColors As Long
'    lCustData As LongPtr
'    lpfnHook As LongPtr
'    lpTemplateName As String
'    hInstance As LongPtr
'    lpszStyle As String
'    nFontType As Integer
'    MISSING_ALIGNMENT As Integer
'    nSizeMin As Long
'    nSizeMax As Long
'End Type
Private Type CHOOSE_FONTA
    lStructSize As Long
    hwndOwner As LongPtr       '  caller's window handle
    hDC As LongPtr             '  printer DC/IC or NULL
    lpLogFont As LongPtr       '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    Flags As Long              '  enum. type flags
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

Public Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONTA) As Long


'Private Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
'(pChoosefont As FONTSTRUC) As Long
'Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
'Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" _
'  (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
'  (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
'Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
'
'
'Public Declare PtrSafe Function WindowFromAccessibleObject Lib "Oleacc" (ByVal pacc As Object, phwnd As LongPtr) As Long

'Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
'Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONTA) As LongPtr
'Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' Logical Font
'Private Const LF_FACESIZE = 32
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

' Font Families
Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.

'/* Font Weights */
Public Const FW_DONTCARE As Long = 0
Public Const FW_THIN  As Long = 100
Public Const FW_EXTRALIGHT  As Long = 200
Public Const FW_LIGHT  As Long = 300
Public Const FW_NORMAL As Long = 400
Public Const FW_MEDIUM  As Long = 500
Public Const FW_SEMIBOLD As Long = 600
Public Const FW_BOLD  As Long = 700
Public Const FW_EXTRABOLD  As Long = 800
Public Const FW_HEAVY  As Long = 900

Public Const FW_ULTRALIGHT  As Long = FW_EXTRALIGHT
Public Const FW_REGULAR  As Long = FW_NORMAL
Public Const FW_DEMIBOLD  As Long = FW_SEMIBOLD
Public Const FW_ULTRABOLD  As Long = FW_EXTRABOLD
Public Const FW_BLACK  As Long = FW_HEAVY

Public Const OUT_DEFAULT_PRECIS  As Long = 0
Public Const OUT_STRING_PRECIS  As Long = 1
Public Const OUT_CHARACTER_PRECIS  As Long = 2
Public Const OUT_STROKE_PRECIS  As Long = 3
Public Const OUT_TT_PRECIS  As Long = 4
Public Const OUT_DEVICE_PRECIS  As Long = 5
Public Const OUT_RASTER_PRECIS  As Long = 6
Public Const OUT_TT_ONLY_PRECIS  As Long = 7
Public Const OUT_OUTLINE_PRECIS  As Long = 8
Public Const OUT_SCREEN_OUTLINE_PRECIS  As Long = 9
Public Const OUT_PS_ONLY_PRECIS  As Long = 10

Public Const CLIP_DEFAULT_PRECIS  As Long = 0
Public Const CLIP_CHARACTER_PRECIS  As Long = 1
Public Const CLIP_STROKE_PRECIS  As Long = 2
Public Const CLIP_MASK  As Long = &HF
Public Const CLIP_LH_ANGLES  As Long = &H10
Public Const CLIP_TT_ALWAYS  As Long = &H20
Public Const CLIP_DFA_DISABLE  As Long = &H40
Public Const CLIP_EMBEDDED  As Long = &H80

Public Const DEFAULT_QUALITY  As Long = 0
Public Const DRAFT_QUALITY  As Long = 1
Public Const PROOF_QUALITY  As Long = 2
Public Const NONANTIALIASED_QUALITY  As Long = 3
Public Const ANTIALIASED_QUALITY  As Long = 4
Public Const CLEARTYPE_QUALITY  As Long = 5
Public Const CLEARTYPE_NATURAL_QUALITY  As Long = 6

Public Const DEFAULT_PITCH  As Long = 0
Public Const FIXED_PITCH  As Long = 1
Public Const VARIABLE_PITCH  As Long = 2
Public Const MONO_FONT  As Long = 8

Public Const ANSI_CHARSET  As Long = 0
Public Const DEFAULT_CHARSET  As Long = 1
Public Const SYMBOL_CHARSET  As Long = 2
Public Const SHIFTJIS_CHARSET  As Long = 128
Public Const HANGEUL_CHARSET  As Long = 129
Public Const HANGUL_CHARSET  As Long = 129
Public Const GB2312_CHARSET  As Long = 134
Public Const CHINESEBIG5_CHARSET  As Long = 136
Public Const OEM_CHARSET  As Long = 255
Public Const JOHAB_CHARSET  As Long = 130
Public Const HEBREW_CHARSET  As Long = 177
Public Const ARABIC_CHARSET  As Long = 178
Public Const GREEK_CHARSET  As Long = 161
Public Const TURKISH_CHARSET  As Long = 162
Public Const VIETNAMESE_CHARSET  As Long = 163
Public Const THAI_CHARSET  As Long = 222
Public Const EASTEUROPE_CHARSET  As Long = 238
Public Const RUSSIAN_CHARSET  As Long = 204

Public Const MAC_CHARSET  As Long = 77
Public Const BALTIC_CHARSET  As Long = 186

'' Logical Font
'Public Const LF_FACESIZE = 32
'Public Const LF_FULLFACESIZE = 64

'Function ChooseFontDialog(ByRef cf As CHOOSE_FONTA) As Boolean
'    ChooseFontDialog = False
'    Debug.Print "ChooseFontDialog"
'    If Not ChooseFontA(cf) Then Exit Function
'    ChooseFontDialog = True
'End Function

Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
    Dim lngTemp As Long
    On Error GoTo MulDiv_err
    If In3 <> 0 Then
        lngTemp = In1 * In2
        lngTemp = lngTemp / In3
    Else
        lngTemp = -1
    End If
MulDiv_end:
    MulDiv = lngTemp
    Exit Function
MulDiv_err:
    lngTemp = -1
    Resume MulDiv_err
End Function

Private Function ByteToString(aBytes() As Byte) As String
    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes)
        dwByteVal = aBytes(dwBytePoint)
        If dwByteVal = 0 Then
            ByteToString = szOut
            Exit Function
        Else
            szOut = szOut & Chr$(dwByteVal)
        End If
        dwBytePoint = dwBytePoint + 1
    Wend
    ByteToString = szOut
End Function

Private Sub StringToByte(InString As String, ByteArray() As Byte)
    Dim intLbound As Integer
    Dim intUbound As Integer
    Dim intLen As Integer
    Dim intX As Integer
    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
    ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
    Next
End Sub


Public Function DialogFont(ByRef f As FormFontInfo) As Boolean
    Dim LF As LOGFONTA, FS As CHOOSE_FONTA
    Dim lLogFontAddress As LongPtr, lMemHandle As LongPtr
    
    LF.lfWeight = f.Weight
    LF.lfItalic = f.Italic * -1
    LF.lfUnderline = f.UnderLine * -1
    LF.lfHeight = -MulDiv(CLng(f.Height), GetDeviceCaps(GetDC(UserForm1.hWndFrame), LOGPIXELSY), 72)
    Call StringToByte(f.Name, LF.lfFaceName())
    FS.rgbColors = f.Color
    FS.lStructSize = LenB(FS)
    
    ' To be modal must be valid Hwnd
    FS.hwndOwner = UserForm1.hWndFrame
      
    lMemHandle = GlobalAlloc(GHND, LenB(LF))
    If lMemHandle = 0 Then
        DialogFont = False
        Exit Function
    End If
    
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        DialogFont = False
        Exit Function
    End If

    CopyMemory ByVal lLogFontAddress, LF, LenB(LF)
    FS.lpLogFont = lLogFontAddress
    FS.Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    Dim apiRetVal As Long
    apiRetVal = ChooseFont(FS)
    If apiRetVal = 1 Then
        CopyMemory LF, ByVal lLogFontAddress, LenB(LF)
        f.Weight = LF.lfWeight
        f.Italic = CBool(LF.lfItalic)
        f.UnderLine = CBool(LF.lfUnderline)
'        Debug.Print "Name:" & LF.lfFaceName
        f.Name = ByteToString(LF.lfFaceName())
        f.Height = CLng(FS.iPointSize / 10)
        f.Color = FS.rgbColors
        
        Debug.Print "CharSet:" & LF.lfCharSet
        
        DialogFont = True
    Else
        DialogFont = False
    End If
End Function

Function test_DialogFont(ctl As Control) As Boolean
    Dim f As FormFontInfo
    With f
        .Color = 0
        .Height = 12
        .Weight = 700
        .Italic = False
        .UnderLine = False
        .Name = "Arial"
    End With
    
'    Dim hWndFrame As LongPtr
'    hWndFrame = ctl.[_GethWnd]

    Call DialogFont(f)
    With f
        Debug.Print "Font Name: "; .Name
        Debug.Print "Font Size: "; .Height
        Debug.Print "Font Weight: "; .Weight
        Debug.Print "Font Italics: "; .Italic
        Debug.Print "Font Underline: "; .UnderLine
        Debug.Print "Font COlor: "; .Color
        

        ctl.Font.Name = .Name
        ctl.Font.SIZE = .Height
        ctl.Font.Weight = .Weight
        ctl.Font.Italic = .Italic
        ctl.Font.UnderLine = .UnderLine
        ctl.Caption = .Name & " - Size:" & .Height
    End With
    test_DialogFont = True
End Function
