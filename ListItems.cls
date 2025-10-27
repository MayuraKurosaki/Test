VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ListView Item collection
Private m_ListItems As Collection

'--------------Compatible Properties----------------
' Returns the number of objects in a collection.
Public Property Get Count() As Long
    Count = m_ListItems.Count
End Property

'Returns a specific item of a Collection object either by position or by key.
Public Property Get Item(Index As Variant) As ListItem
    Set Item = m_ListItems.Item(Index)
End Property


'--------------Compatible Methods----------------
'Adds a ListItem object to a ListItems collection only at run time.
Public Function Add(Optional _
    ByVal Index As Variant, _
    Optional ByVal Key As Variant, _
    Optional ByVal Text As Variant, _
    Optional ByVal Icon As Variant, _
    Optional ByVal SmallIcon As Variant) As ListItem
    Dim NewItem As ListItem
    Set NewItem = New ListItem
    With NewItem
        .Index = Index
        .Key = Key
        .Text = Text
        .Icon = Icon
        .SmallIcon = SmallIcon
    End With
    m_ListItems.Add NewItem
    Set Add = NewItem
End Function

'Removes all objects in a collection.
Public Sub Clear()
    m_ListItems.Clear
End Sub

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
    m_ListItems.Remove Index
End Sub

Public Function _NewEnum() As IUnknown
    Set _NewEnum = m_ListItems._NewEnum
End Function
