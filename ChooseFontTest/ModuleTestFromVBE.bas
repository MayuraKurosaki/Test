Attribute VB_Name = "ModuleTestFromVBE"
Option Explicit


' run to test the 'Font' dialog ... settings chosen in the dialog will be remembered from one 'run' to the next (until
' the VBA Project's state is reset)
Sub TestFromVBE()
    Static tFormFontInfo As FormFontInfo
    With tFormFontInfo
        .iHeight = IIf(.iHeight = 0, 12, .iHeight)
        .iWeight = IIf(.iWeight = 0, 700, .iWeight)
        .sName = IIf(Len(.sName) = 0, "Arial", .sName)
    End With
    
    Dim bWasCancelled As Boolean
    If ModuleChooseFont.TryShowFontDialog(tFormFontInfo, bWasCancelled, True) Then
        If bWasCancelled Then
            Debug.Print Now, "Dialog cancelled"
        Else
            With tFormFontInfo
                Debug.Print Now, "Font Name: '" & .sName & "'"
                Debug.Print Now, "Font Size: " & .iHeight
                Debug.Print Now, "Font Weight: " & .iWeight
                Debug.Print Now, "Font Italics: " & .bItalic
                Debug.Print Now, "Font Underline: " & .bUnderLine
                Debug.Print Now, "Font StrikeOut: " & .bStrikeOut
                Debug.Print Now, "Font Color: " & .lColor
                Debug.Print Now, "Font CharSet: " & .iCharSet
            End With
        End If
    End If
End Sub


