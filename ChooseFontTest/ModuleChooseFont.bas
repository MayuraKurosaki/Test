Attribute VB_Name = "ModuleChooseFont"
' ---------------------------------------------------------------------------------------------------------------------
' Purpose: Show the 'Font' dialog to allow the user to choose a Font including font name, point size, underline and
'     other attributes
' External dependencies: The Windows API
'
' Note that ...
'
' Original code by Terry Kreft
'
' Modified by Stephen Lebans (http://lebans.com/)
'
' Revised January 2019 by Philipp Stiefel (http://codekabinett.com) to make it run in x64 VBA-Applications
'
' Updated October 2025 by John Mallinson (https://www.thevbahelp.com/):
' * Now uses Unicode (not ANSI) Functions and Types
' * Includes setting / getting Strikeout and CharSet (aka Script)
' * The Device Context is released
' * Global memory is released
' ---------------------------------------------------------------------------------------------------------------------


Option Explicit


Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const CF_EFFECTS = &H100&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_SCREENFONTS = &H1

Private Const LOGPIXELSY = 90

Private Const LF_FACESIZE = 32

Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As LongPtr
    hDC As LongPtr
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
    lfFaceName(1 To LF_FACESIZE) As Integer
End Type

Private Declare PtrSafe Function ChooseFontW Lib "comdlg32.dll" (ByVal lpcf As LongPtr) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
        ByVal Length As LongPtr)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long


' ---------------------------------------------------------------------------------------------------------------------
' Purpose: The initial (and return) settings for the Font dialog
' ---------------------------------------------------------------------------------------------------------------------
Type FormFontInfo
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
    Dim tFormFontInfo As FormFontInfo
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
    If TryShowFontDialog(tFormFontInfo, bWasCancelled, False) Then
        If bWasCancelled Then
            MsgBox "Cancelled!"
        Else
            With tFormFontInfo
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
Function TryShowFontDialog(ByRef outtFormFontInfo As FormFontInfo, ByRef outbWasCancelled As Boolean, _
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
    Dim hDC As LongPtr
    hDC = GetDC(hWnd)
    GetDeviceCapsValue = GetDeviceCaps(hDC, lInformationType)
    ReleaseDC hWnd, hDC
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


