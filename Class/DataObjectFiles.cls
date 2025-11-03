VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataObjectFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Collection of file paths for DataObject (vbCFFiles format)"
Option Explicit

'Collection of file paths for DataObject (vbCFFiles format)

'--------------Fields----------------
Private This As Collection 'Collection to hold file paths


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This = New Collection
End Sub

Private Sub Class_Terminate()
    Set This = Nothing
End Sub


'--------------Compatible Properties----------------
'Returns the number of filenames in the Files collection of a DataObject object (vbCFFiles format only).
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of filenames in the Files collection of a DataObject object (vbCFFiles format only)."
'Attribute Count.VB_Description = "Returns the number of filenames in the Files collection of a DataObject object (vbCFFiles format only)."
    If This Is Nothing Then
        Count = 0
    Else
        Count = This.Count
    End If
End Property

'Returns a specific filename by index from the Files collection of a DataObject object (vbCFFiles format only).
Public Property Get Item(Index As Long) As String
Attribute Item.VB_Description = "Returns a specific filename by index from the Files collection of a DataObject object (vbCFFiles format only)."
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_Description = "Returns a specific filename by index from the Files collection of a DataObject object (vbCFFiles format only)."
'Attribute Item.VB_UserMemId = 0
    If This Is Nothing Then
        Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
    Else
        If Index < 1 Or Index > This.Count Then
            Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
            Item = This.Item(Index)
        End If
    End If
End Property


'--------------Compatible Methods----------------
'Adds a filename to the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Add(ByVal FilePath As String)
Attribute Add.VB_Description = "Adds a filename to the Files collection of a DataObject object (vbCFFiles format only)."
'Attribute Add.VB_Description = "Adds a filename to the Files collection of a DataObject object (vbCFFiles format only)."
    'Add the new file to the collection.
    If This Is Nothing Then
        Set This = New Collection
    End If
    This.Add FilePath
End Sub

'Clears all filenames stored in the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
'Attribute Clear.VB_Description = "Removes all objects in a collection."
    'Clears all files from the DataObjectFiles collection.
    If Not This Is Nothing Then
        Set This = Nothing
    End If
End Sub

'Returns an enumerator for iterating through the Files collection of a DataObject object (vbCFFiles format only).
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
'Attribute NewEnum.VB_UserMemId = -4
    If This Is Nothing Then
        Err.Raise vbObjectError + 1, "DataObjectFiles", "No files in collection"
    Else
        Set NewEnum = This.[_NewEnum]
    End If
End Function

'Removes a filename from the Files collection of a DataObject object (vbCFFiles format only).
Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes a filename from the Files collection of a DataObject object (vbCFFiles format only)."
'Attribute Remove.VB_Description = "Removes a filename from the Files collection of a DataObject object (vbCFFiles format only)."
    'Removes the file at the specified index from the DataObjectFiles collection.
    If Not This Is Nothing Then
        If Index < 1 Or Index > This.Count Then
            Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
            This.Remove Index
        End If
    End If
End Sub



