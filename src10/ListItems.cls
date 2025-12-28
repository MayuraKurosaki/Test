VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ListView Item collection"
Option Explicit

'ListView Item collection

'--------------Fields----------------
Private This As Collection  'Collection of ListItem objects


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This = New Collection
End Sub

Private Sub Class_Terminate()
    Call Clear
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

'Returns a specific item of a Collection object either by position or by key.
Public Property Get Item(Index As Long) As ListItem
Attribute Item.VB_Description = "Returns a specific item of a Collection object either by position or by key."
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_Description = "Returns a specific item of a Collection object either by position or by key."
'Attribute Item.VB_UserMemId = 0
    If This Is Nothing Then
        Err.Raise vbObjectError + 1, "ListItems", "Index out of bounds"
    Else
        If Index < 0 Or Index > This.Count Then
            Err.Raise vbObjectError + 1, "ListItems", "Index out of bounds"
        Else
            Set Item = This.Item(Index + 1)
        End If
    End If
End Property


'--------------Compatible Methods----------------
'Adds a ListItem object to a ListItems collection only at run time.
Public Function Add(Optional Index As Long, Optional Key As String, Optional Text As String, Optional Icon As Variant, Optional SmallIcon As Variant) As ListItem
Attribute Add.VB_Description = "Adds a ListItem object to a ListItems collection only at run time."
'Attribute Add.VB_Description = "Adds a ListItem object to a ListItems collection only at run time."
    Dim NewItem As ListItem
    Set NewItem = New ListItem
    With NewItem
        .Index = Index
        .Key = Key
        .Text = Text
        .Icon = Icon
        .SmallIcon = SmallIcon
    End With
    This.Add NewItem
    Set Add = NewItem
End Function

'Removes all objects in a collection.
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all objects in a collection."
'Attribute Clear.VB_Description = "Removes all objects in a collection."
    Set This = Nothing
'    This.Clear
End Sub

'Removes a specific member from a collection.
Public Sub Remove(Index As Long)
Attribute Remove.VB_Description = "Removes a specific member from a collection."
'Attribute Remove.VB_Description = "Removes a specific member from a collection."
    This.Remove Index
End Sub

'Returns an enumerator for iterating through the ListItems collection.
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ListItems collection."
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ListItems collection."
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function
