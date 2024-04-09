Attribute VB_Name = "dp_icon"
Option Explicit

Global pathToIcon As String

'creates the icon for the in cell date picker
Sub CreateDPIcon()
    
    'set the path to the icon
    pathToIcon = Environ("temp") & "\samrad3.bmp"
    
    'check if the bmp exists
    If Dir(pathToIcon) = "" Then
        'icon doesn't exist, need to create it
        
        Dim hex_val As String
        hex_val = ThisWorkbook.Sheets("_data12345").Range("iconInCell").Value
    
        Dim output() As String
        output = Split(hex_val, "|")
    
        Dim handle As Long
        handle = FreeFile
        
        Open pathToIcon For Binary As #handle
    
        Dim i As Long
        For i = LBound(output) To UBound(output)
            Put #handle, , CByte("&H" & output(i))
        Next i
    
        Close #handle
    End If
    
End Sub


