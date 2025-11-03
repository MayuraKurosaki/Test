VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ListView Column Header collection"
Option Explicit

'ListView Column Header collection

'--------------Fields----------------
Private This As Collection   'Collection of ColumnHeader objects


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This = New Collection
End Sub

Private Sub Class_Terminate()
    Set This = Nothing
End Sub


'--------------Compatible Properties----------------
'Returns the number of objects in a collection.
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of objects in a collection."
'Attribute Count.VB_Description = "Returns the number of objects in a collection."
    If This Is Nothing Then
        Count = 0
    Else
        Count = This.Count
    End If
End Property

'Returns a specific member of a Collection object either by position or by key.
Public Property Get Item(Index As Variant) As ColumnHeader
Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or by key."
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or by key."
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
'Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time.
Public Function Add(Optional _
Attribute Add.VB_Description = "Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time."
'Attribute Add.VB_Description = "Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time."
    ByVal Index As Variant, _
    Optional ByVal Key As Variant, _
    Optional ByVal Text As Variant, _
    Optional ByVal Icon As Variant, _
    Optional ByVal Width As Variant, _
    Optional ByVal Alignment As Variant) As ColumnHeader
    Dim NewHeader As ColumnHeader
    Set NewHeader = New ColumnHeader
    With NewHeader
        .Index = Index
        .Key = Key
        .Text = Text
        .Icon = Icon
        If Not IsMissing(Width) Then .Width = Width
        If Not IsMissing(Alignment) Then .Alignment = Alignment
    End With
    This.Add NewHeader
    Set Add = NewHeader
End Function

'Removes all objects in a collection.
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
'Attribute Clear.VB_Description = "Removes all objects in a collection."
    This.Clear
End Sub

'Returns an enumerator for iterating through the ColumnHeaders collection.
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from a collection."
'Attribute Remove.VB_Description = "Removes a specific member from a collection."
    This.Remove Index
End Sub