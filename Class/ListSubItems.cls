VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSubItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ListView SubItem collection"
Option Explicit

'ListView SubItem collection

'--------------Fields----------------
Private This As Collection  'Collection of ListSubItem objects


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This = New Collection
End Sub

Private Sub Class_Terminate()
    Set This = Nothing
End Sub


'--------------Compatible Properties----------------
'Returns the number of SubItems in the collection.
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of SubItems in the collection."
'Attribute Count.VB_Description = "Returns the number of SubItems in the collection."
    If This Is Nothing Then
        Count = 0
    Else
        Count = This.Count
    End If
End Property

'Returns a specific item of a Collection object either by position or by key.
Public Property Get Item(Index As Variant) As ListSubItem
Attribute Item.VB_Description = "Returns a specific item of a Collection object either by position or by key."
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_Description = "Returns a specific item of a Collection object either by position or by key."
'Attribute Item.VB_UserMemId = 0
    If This Is Nothing Then
        Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
    Else
        If Index < 1 Or Index > This.Count Then
            Err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
            Set Item = This.Item(Index)
        End If
    End If
End Property


'--------------Compatible Methods----------------
'Adds a ListSubItem object to a ListSubItems collection only at run time.
Public Function Add(Optional _
Attribute Add.VB_Description = "Adds a ListSubItem object to a ListSubItems collection only at run time."
'Attribute Add.VB_Description = "Adds a ListSubItem object to a ListSubItems collection only at run time."
    ByVal Index As Variant, _
    Optional ByVal Key As Variant, _
    Optional ByVal Text As Variant, _
    Optional ByVal ReportIcon As Variant) As ListSubItem
    Dim NewSubItem As ListSubItem
    Set NewSubItem = New ListSubItem
    With NewSubItem
        .Index = Index
        .Key = Key
        .Text = Text
        .ReportIcon = ReportIcon
    End With
    This.Add NewSubItem
    Set Add = NewSubItem
End Function

'Removes all objects in a collection.
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
'Attribute Clear.VB_Description = "Removes all objects in a collection."
    This.Clear
End Sub

'Returns an enumerator for iterating through the ListSubItem collection.
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ListSubItem collection."
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ListSubItem collection."
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from a collection."
'Attribute Remove.VB_Description = "Removes a specific member from a collection."
    This.Remove Index
End Sub
