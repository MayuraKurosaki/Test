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
'Type FormFontInfo2
'    sName As String
'    iWeight As Integer
'    iHeight As Integer
'    bUnderLine As Boolean
'    bItalic As Boolean
'    bStrikeOut As Boolean
'    iCharSet As Integer
'    lColor As Long
'End Type

'Private Type LOGFONT
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
'    lfFaceName(LF_FACESIZE) As Byte
'End Type
Public Type LOGFONTW
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
    lfFaceName(1 To LF_FACESIZE) As Integer
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
Private Type CHOOSEFONTW
    lStructSize As Long
    hwndOwner As LongPtr
    hdc As LongPtr
    lpLogFont As LongPtr
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    hInstance As LongPtr
    lpszStyle As LongPtr
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
Private Type CHOOSE_FONTA
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

Public Declare PtrSafe Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONTA) As Long
Private Declare PtrSafe Function CHOOSEFONTW Lib "comdlg32.dll" Alias "ChooseFontW" (ByVal lpcf As LongPtr) As Long


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

' ---------------------------------------------------------------------------------------------------------------------
' Purpose: The initial (and return) settings for the Font dialog
' ---------------------------------------------------------------------------------------------------------------------
Type FormFontInfo2
    sName As String
    iWeight As Integer
    iHeight As Integer
    bUnderLine As Boolean
    bItalic As Boolean
    bStrikeOut As Boolean
    iCharSet As Integer
    lColor As Long
End Type


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Call from a UserForm (eg click handler of a Button) ... the 'Font' dialog will be shown populated with Font
'     details from the TextBox ... if the user clicks 'OK' in the dialog, the TextBox will be updated with the selected
'     Font and text colour
' Param txtBox: A TextBox on the UserForm
' Returns: True if the dialog was closed by the 'OK' button, False otherwise
' ---------------------------------------------------------------------------------------------------------------------
Function TestWithTextBox(txtBox As Object) As Boolean
    Dim tFormFontInfo As FormFontInfo2
    With tFormFontInfo
        .sName = txtBox.fontName
        .iHeight = txtBox.fontSize
        .iWeight = txtBox.FontWeight
        .bItalic = txtBox.FontItalic
        .bUnderLine = txtBox.FontUnderline
        .bStrikeOut = txtBox.FontStrikethru
        .lColor = txtBox.ForeColor
    End With
    
    Dim bWasCancelled  As Boolean
'    If TryShowFontDialog(tFormFontInfo, bWasCancelled, False) Then
    If TryShowFontDialog(tFormFontInfo, bWasCancelled, True) Then
        If bWasCancelled Then
            MsgBox "Cancelled!"
        Else
            With tFormFontInfo
                Debug.Print "FontName:" & .sName
                Debug.Print "Font Name: " & .sName
                Debug.Print "Font Size: " & .iHeight
                Debug.Print "Font Weight: " & .iWeight
                Debug.Print "Font Italics: " & .bItalic
                Debug.Print "Font Underline: " & .bUnderLine
                Debug.Print "Font Strikethru: " & .bStrikeOut
                Debug.Print "Font COlor: " & .lColor
                txtBox.fontName = .sName
                txtBox.fontSize = .iHeight
                txtBox.FontWeight = .iWeight
                txtBox.FontItalic = .bItalic
                txtBox.FontUnderline = .bUnderLine
                txtBox.FontStrikethru = .bStrikeOut
                txtBox.ForeColor = .lColor
            End With
            
            TestWithTextBox = True
        End If
    End If
End Function


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Try to show the 'Font' dialog
' Param outtFormFontInfo: A FormFontInfo which specifies the initial settings for the Font dialog and, if the user
'     clicks the 'OK' button in the dialog, will be updated with the values selected by the user in the dialog
' Param outbWasCancelled: True if the user clicks the 'Cancel' button, False otherwise
' Param bAllowScriptSelection: True to allow selection of the 'Script' (aka CharSet) in the Font dialog
' Param hWnd: Optionally, a window handle if the Font dialog should be shown modally
' Returns: True if the Font dialog could be shown (this will be True whether the user clicks 'OK' or 'Cancel'), False
'     if there was a memory allocation problem meaning that the dialog could not be shown
' ---------------------------------------------------------------------------------------------------------------------
Function TryShowFontDialog(ByRef outtFormFontInfo As FormFontInfo2, ByRef outbWasCancelled As Boolean, _
        bAllowScriptSelection As Boolean, Optional hWnd As LongPtr = 0) As Boolean
    Dim tLogFont As LOGFONT, tCHOOSEFONT As CHOOSEFONT
    Dim lLogFontAddress As LongPtr, lMemHandle As LongPtr
    
    With tLogFont
        .lfHeight = -GetFontHeightInLogicalUnits(CLng(outtFormFontInfo.iHeight), hWnd)
        .lfWeight = outtFormFontInfo.iWeight
        .lfItalic = outtFormFontInfo.bItalic * -1
        .lfUnderline = outtFormFontInfo.bUnderLine * -1
        .lfStrikeOut = outtFormFontInfo.bStrikeOut * -1
        .lfCharSet = outtFormFontInfo.iCharSet
        IntegerArrayFixedBoundsFromString outtFormFontInfo.sName, .lfFaceName
    End With
    
    With tCHOOSEFONT
        .lStructSize = LenB(tCHOOSEFONT)
        .rgbColors = outtFormFontInfo.lColor
        .hwndOwner = hWnd ' to be modal must be a valid Hwnd
        
        lMemHandle = GlobalAlloc(GHND, LenB(tLogFont))
        If lMemHandle = 0 Then Exit Function
    
        lLogFontAddress = GlobalLock(lMemHandle)
        If lLogFontAddress = 0 Then Exit Function
    
        CopyMemory ByVal lLogFontAddress, tLogFont, LenB(tLogFont)
        .lpLogFont = lLogFontAddress
        .flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT Or IIf(bAllowScriptSelection, 0, CF_NOSCRIPTSEL)
    End With
    
    TryShowFontDialog = True
    If CHOOSEFONTW(VarPtr(tCHOOSEFONT)) Then
        CopyMemory tLogFont, ByVal lLogFontAddress, LenB(tLogFont)
        With outtFormFontInfo
            .iWeight = tLogFont.lfWeight
            .bItalic = CBool(tLogFont.lfItalic)
            .bUnderLine = CBool(tLogFont.lfUnderline)
            .bStrikeOut = CBool(tLogFont.lfStrikeOut)
            .iCharSet = tLogFont.lfCharSet
            .sName = StringFromIntegerArray(tLogFont.lfFaceName)
            .iHeight = CLng(tCHOOSEFONT.iPointSize / 10)
            .lColor = tCHOOSEFONT.rgbColors
        End With
    Else
        outbWasCancelled = True
    End If
    
    GlobalUnlock lMemHandle
    GlobalFree lMemHandle
End Function


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Calculate font the height in logical units
' Param lFontHeightInPoints: The font height in points
' Param hWnd: The window handle
' ---------------------------------------------------------------------------------------------------------------------
Private Function GetFontHeightInLogicalUnits(lFontHeightInPoints As Long, hWnd As LongPtr) As Long
    Dim lPixelsPerInch As Long
    lPixelsPerInch = GetDeviceCapsValue(LOGPIXELSY, hWnd)
    ' for '72', see the docs for the lfHeight member of LOGFONT
    GetFontHeightInLogicalUnits = (lFontHeightInPoints * lPixelsPerInch) / 72
End Function


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Get information on device capabilities ("DeviceCaps" = device capabilities)
' Param lInformationType: The information type ... see
'     https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdevicecaps
' Param hWnd: Optionally, provide a handle to a window, or omit to get device capabilties for the entire screen
' ---------------------------------------------------------------------------------------------------------------------
Private Function GetDeviceCapsValue(lInformationType As Long, Optional hWnd As LongPtr = 0) As Long
    Dim hdc As LongPtr
    hdc = GetDC(hWnd)
    GetDeviceCapsValue = GetDeviceCaps(hdc, lInformationType)
    ReleaseDC hWnd, hdc
End Function


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Get a (Unicode) String from an array of Integers
' Param aInts: The Integer array which can be fixed- or variable-bounds
' Returns: The String
' ---------------------------------------------------------------------------------------------------------------------
Private Function StringFromIntegerArray(aInts() As Integer) As String
    Dim lInts As Long, sText As String
    lInts = UBound(aInts) - LBound(aInts) + 1
    If lInts > 0 Then
        sText = String$(lInts + 1, 0)
        CopyMemory ByVal StrPtr(sText), aInts(LBound(aInts)), lInts * 2
        sText = Left$(sText, InStr(1, sText, vbNullChar, vbBinaryCompare) - 1)
        StringFromIntegerArray = sText
    End If
End Function


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Get a fixed-bounds Integer array from a (Unicode) String
' Param sString: The String to copy to the Integer array
' Param outaInts: The Integer array which must be large enough to accommodate the String (2 bytes per character and 2
'     bytes per Integer)
' Returns: True if successful, False if the Integer array was not large enough or the String was zero-length
' ---------------------------------------------------------------------------------------------------------------------
Private Function IntegerArrayFixedBoundsFromString(sString As String, ByRef outaInts() As Integer) As Boolean
    Dim lInts As Long
    lInts = LenB(sString) / 2
    If lInts > 0 And lInts <= (UBound(outaInts) - LBound(outaInts) + 1) Then
        CopyMemory outaInts(LBound(outaInts)), ByVal StrPtr(sString), lInts * 2
        IntegerArrayFixedBoundsFromString = True
    End If
End Function



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
    Dim lf As LOGFONTA, FS As CHOOSE_FONTA
    Dim lLogFontAddress As LongPtr, lMemHandle As LongPtr
    
    lf.lfWeight = f.Weight
    lf.lfItalic = f.Italic * -1
    lf.lfUnderline = f.UnderLine * -1
    lf.lfHeight = -MulDiv(CLng(f.Height), GetDeviceCaps(GetDC(UserForm1.hWndFrame), LOGPIXELSY), 72)
    Call StringToByte(f.Name, lf.lfFaceName())
    FS.rgbColors = f.Color
    FS.lStructSize = LenB(FS)
    
    ' To be modal must be valid Hwnd
    FS.hwndOwner = UserForm1.hWndFrame
      
    lMemHandle = GlobalAlloc(GHND, LenB(lf))
    If lMemHandle = 0 Then
        DialogFont = False
        Exit Function
    End If
    
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        DialogFont = False
        Exit Function
    End If

    CopyMemory ByVal lLogFontAddress, lf, LenB(lf)
    FS.lpLogFont = lLogFontAddress
    FS.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    Dim apiRetVal As Long
    apiRetVal = CHOOSEFONT(FS)
    If apiRetVal = 1 Then
        CopyMemory lf, ByVal lLogFontAddress, LenB(lf)
        f.Weight = lf.lfWeight
        f.Italic = CBool(lf.lfItalic)
        f.UnderLine = CBool(lf.lfUnderline)
'        Debug.Print "Name:" & LF.lfFaceName
        f.Name = ByteToString(lf.lfFaceName())
        f.Height = CLng(FS.iPointSize / 10)
        f.Color = FS.rgbColors
        
        Debug.Print "CharSet:" & lf.lfCharSet
        
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
