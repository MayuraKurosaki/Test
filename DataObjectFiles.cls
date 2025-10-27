ERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataObjectFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DataObjectFiles As Collection
Private Type FileEntry
    FilePath As String
    FileName As String

End Type
'--------------Compatible Properties----------------
'Returns the number of filenames in the Files collection of a DataObject object (vbCFFiles format only).
Public Property Get Count() As Long
    If m_DataObjectFiles Is Nothing Then
        Count = 0
    Else
        Count = m_DataObjectFiles.Count
    End If
End Property

'Returns a specific filename by index from the Files collection of a DataObject object (vbCFFiles format only).
Public Property Get Item(Index As Long) As String
    If m_DataObjectFiles Is Nothing Then
        Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
    Else
        If Index < 1 Or Index > m_DataObjectFiles.Count Then
            Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
            Dim FileEntryItem As FileEntry
            FileEntryItem = m_DataObjectFiles.Item(Index)
            Item = FileEntryItem.FilePath
        End If
    End If
End Property

'--------------Compatible Methods----------------
'Adds a filename to the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Add(ByVal FilePath As String)
    Dim NewFile As FileEntry
    NewFile.FilePath = FilePath
    NewFile.FileName = Dir(FilePath)
    'Add the new file to the collection.
    If m_DataObjectFiles Is Nothing Then
        Set m_DataObjectFiles = New Collection
    End If
    m_DataObjectFiles.Add NewFile
End Sub

'Clears all filenames stored in the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Clear()
    'Clears all files from the DataObjectFiles collection.
    If Not m_DataObjectFiles Is Nothing Then
        Set m_DataObjectFiles = Nothing
    End If
End Sub

Public Function _NewEnum() As IUnknown
    If m_DataObjectFiles Is Nothing Then
        Err.Raise vbObjectError + 1, "DataObjectFiles", "No files in collection"
    Else
        Set _NewEnum = m_DataObjectFiles._NewEnum
    End If
End Function

'Removes a filename from the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Remove(ByVal Index As Long)
    'Removes the file at the specified index from the DataObjectFiles collection.
    If Not m_DataObjectFiles Is Nothing Then
        If Index < 1 Or Index > m_DataObjectFiles.Count Then
            Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
            m_DataObjectFiles.Remove Index
        End If
    End If
End Sub



