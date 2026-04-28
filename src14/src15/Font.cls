VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400

Private Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Private Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)

'Private Const LOGPIXELSY = 90
' Logical Font
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

' Font Families
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Private Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Private Const FF_SCRIPT = 64     '  Cursive, etc.
Private Const FF_DECORATIVE = 80 '  Old English, etc.

'/* Font Weights */
Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN  As Long = 100
Private Const FW_EXTRALIGHT  As Long = 200
Private Const FW_LIGHT  As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_MEDIUM  As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_BOLD  As Long = 700
Private Const FW_EXTRABOLD  As Long = 800
Private Const FW_HEAVY  As Long = 900

Private Const FW_ULTRALIGHT  As Long = FW_EXTRALIGHT
Private Const FW_REGULAR  As Long = FW_NORMAL
Private Const FW_DEMIBOLD  As Long = FW_SEMIBOLD
Private Const FW_ULTRABOLD  As Long = FW_EXTRABOLD
Private Const FW_BLACK  As Long = FW_HEAVY

Private Const OUT_DEFAULT_PRECIS  As Long = 0
Private Const OUT_STRING_PRECIS  As Long = 1
Private Const OUT_CHARACTER_PRECIS  As Long = 2
Private Const OUT_STROKE_PRECIS  As Long = 3
Private Const OUT_TT_PRECIS  As Long = 4
Private Const OUT_DEVICE_PRECIS  As Long = 5
Private Const OUT_RASTER_PRECIS  As Long = 6
Private Const OUT_TT_ONLY_PRECIS  As Long = 7
Private Const OUT_OUTLINE_PRECIS  As Long = 8
Private Const OUT_SCREEN_OUTLINE_PRECIS  As Long = 9
Private Const OUT_PS_ONLY_PRECIS  As Long = 10

Private Const CLIP_DEFAULT_PRECIS  As Long = 0
Private Const CLIP_CHARACTER_PRECIS  As Long = 1
Private Const CLIP_STROKE_PRECIS  As Long = 2
Private Const CLIP_MASK  As Long = &HF
Private Const CLIP_LH_ANGLES  As Long = &H10
Private Const CLIP_TT_ALWAYS  As Long = &H20
Private Const CLIP_DFA_DISABLE  As Long = &H40
Private Const CLIP_EMBEDDED  As Long = &H80

Private Const DEFAULT_QUALITY  As Long = 0
Private Const DRAFT_QUALITY  As Long = 1
Private Const PROOF_QUALITY  As Long = 2
Private Const NONANTIALIASED_QUALITY  As Long = 3
Private Const ANTIALIASED_QUALITY  As Long = 4
Private Const CLEARTYPE_QUALITY  As Long = 5
Private Const CLEARTYPE_NATURAL_QUALITY  As Long = 6

Private Const DEFAULT_PITCH  As Long = 0
Private Const FIXED_PITCH  As Long = 1
Private Const VARIABLE_PITCH  As Long = 2
Private Const MONO_FONT  As Long = 8

Private Const ANSI_CHARSET  As Long = 0
Private Const DEFAULT_CHARSET  As Long = 1
Private Const SYMBOL_CHARSET  As Long = 2
Private Const SHIFTJIS_CHARSET  As Long = 128
Private Const HANGEUL_CHARSET  As Long = 129
Private Const HANGUL_CHARSET  As Long = 129
Private Const GB2312_CHARSET  As Long = 134
Private Const CHINESEBIG5_CHARSET  As Long = 136
Private Const OEM_CHARSET  As Long = 255
Private Const JOHAB_CHARSET  As Long = 130
Private Const HEBREW_CHARSET  As Long = 177
Private Const ARABIC_CHARSET  As Long = 178
Private Const GREEK_CHARSET  As Long = 161
Private Const TURKISH_CHARSET  As Long = 162
Private Const VIETNAMESE_CHARSET  As Long = 163
Private Const THAI_CHARSET  As Long = 222
Private Const EASTEUROPE_CHARSET  As Long = 238
Private Const RUSSIAN_CHARSET  As Long = 204

Private Const MAC_CHARSET  As Long = 77
Private Const BALTIC_CHARSET  As Long = 186

Private Type FormFontInfo
    Name As String
    Weight As Integer
    Height As Integer
    Underline As Boolean
    Italic As Boolean
    Color As Long
End Type

Private Type LOGFONTW
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
Private Type LOGFONTA
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

Private Type CHOOSE_FONTW
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

Private Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONTA) As Long
Private Declare PtrSafe Function ChooseFontW Lib "comdlg32.dll" (ByVal lpcf As LongPtr) As Long

Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONTA) As LongPtr
Private Declare PtrSafe Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONTW) As LongPtr
Private Declare PtrSafe Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
Private Declare PtrSafe Function GetTextMetricsW Lib "gdi32" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRICW) As Long
Private Declare PtrSafe Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare PtrSafe Function GetTextFaceW Lib "gdi32" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFacename As LongPtr) As Long

Private Type Fields
    hFont As LongPtr
    lf As LOGFONTW
    hWnd As LongPtr
    hdc As LongPtr
    
    Size As Currency
    Width As Long
    Weight As Integer
    Escapement As Long
    Orientation As Long
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    Strikethrough As Boolean
    OutPrecision As Byte
    ClipPrecision As Byte
    Quality As Byte
    lfPitchAndFamily As Byte
    CharSet As Integer
    Name As String
    Unicode As Boolean
End Type

Private This As Fields





' ---------------------------------------------------------------------------------------------------------------------
' Purpose: The initial (and return) settings for the Font dialog
' ---------------------------------------------------------------------------------------------------------------------
Private Type FormFontInfo2
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
        .sName = txtBox.FontName
        .iHeight = txtBox.FontSize
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
                txtBox.FontName = .sName
                txtBox.FontSize = .iHeight
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
Private Function TryShowFontDialog(ByRef outtFormFontInfo As FormFontInfo2, ByRef outbWasCancelled As Boolean, _
        bAllowScriptSelection As Boolean, Optional hWnd As LongPtr = 0) As Boolean
    Dim tLogFont As LOGFONT, tCHOOSEFONT As ChooseFont
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
    If ChooseFontW(VarPtr(tCHOOSEFONT)) Then
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


' ANCI
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


Private Function DialogFont(ByRef f As FormFontInfo) As Boolean
    Dim lf As LOGFONTA, FS As CHOOSE_FONTA
    Dim lLogFontAddress As LongPtr, lMemHandle As LongPtr
    
    lf.lfWeight = f.Weight
    lf.lfItalic = f.Italic * -1
    lf.lfUnderline = f.Underline * -1
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
    apiRetVal = ChooseFont(FS)
    If apiRetVal = 1 Then
        CopyMemory lf, ByVal lLogFontAddress, LenB(lf)
        f.Weight = lf.lfWeight
        f.Italic = CBool(lf.lfItalic)
        f.Underline = CBool(lf.lfUnderline)
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
        .Underline = False
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
        Debug.Print "Font Underline: "; .Underline
        Debug.Print "Font COlor: "; .Color
        

        ctl.Font.Name = .Name
        ctl.Font.Size = .Height
        ctl.Font.Weight = .Weight
        ctl.Font.Italic = .Italic
        ctl.Font.Underline = .Underline
        ctl.Caption = .Name & " - Size:" & .Height
    End With
    test_DialogFont = True
End Function

Public Property Get Name() As String
    Name = This.Name
End Property
Public Property Let Name(FontName As String)
    If IntegerArrayFixedBoundsFromString(FontName, This.lf.lfFaceName) Then
        This.Name = FontName
    End If
    Debug.Print "FaceName:" & StringFromIntegerArray(This.lf.lfFaceName)
End Property

Public Property Get Size() As Currency
    Size = This.Size
End Property
Public Property Let Size(PointSize As Currency)
    Dim hWndDesk As LongPtr
    hWndDesk = GetDesktopWindow()

    Dim hDCDesk As LongPtr
    hDCDesk = GetDC(hWndDesk)
    
    This.lf.lfHeight = (PointSize / 72 * GetDeviceCaps(hDCDesk, LOGPIXELSY))
    This.Size = PointSize
End Property

Public Property Get CharSet() As Integer
    CharSet = This.CharSet
End Property
Public Property Let CharSet(CharSet As Integer)
    This.lf.lfCharSet = CByte(CharSet)
    This.CharSet = CharSet
End Property


Public Function Create() As Boolean
    Debug.Print "Create FaceName:" & StringFromIntegerArray(This.lf.lfFaceName)
    Debug.Print "CharSet:" & This.lf.lfCharSet
    Debug.Print "Height:" & This.lf.lfHeight
    This.hFont = CreateFontIndirectW(This.lf)
    Debug.Print "CreateFont"
End Function

Public Function Reflect() As Boolean
    SendMessageW This.hWnd, WM_SETFONT, This.hFont, ByVal 1&
    Debug.Print "SetFont"
End Function

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Function Init(hWnd As LongPtr) As Boolean
    With This
        .hWnd = hWnd
        If IsWindowUnicode(hWnd) <> 0 Then
            .Unicode = True
        Else
            .Unicode = False
        End If
        .hFont = SendMessageW(.hWnd, WM_GETFONT, ByVal 0, ByVal 0)
        If .hFont <> 0 Then
            SendMessageW .hWnd, WM_SETFONT, .hFont, ByVal 1&
            Debug.Print "ResetFontComplete"
            
            GetObjectW .hFont, Len(.lf), VarPtr(.lf)
            Debug.Print "CharSet:" & .lf.lfCharSet
            Debug.Print "Height:" & .lf.lfHeight
        End If
    End With
End Function
