VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSubItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ListView SubItem collection
Private m_ListSubItems As Collection
'--------------Compatible Properties----------------
'Returns the number of SubItems in the collection.
Public Property Get Count() As Long
    Count = m_ListSubItems.Count
End Property

'Returns a specific item of a Collection object either by position or by key.
Public Property Get Item(Index As Variant) As ListSubItem
    Set Item = m_ListSubItems.Item(Index)
End Property

'--------------Compatible Methods----------------
' Adds a ListSubItem object to a ListSubItems collection only at run time.
Public Function Add(Optional _
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
    m_ListSubItems.Add NewSubItem
    Set Add = NewSubItem
End Function

'Removes all objects in a collection.
Public Sub Clear()
    m_ListSubItems.Clear
End Sub

Public Function _NewEnum() As IUnknown
    Set _NewEnum = m_ListSubItems.[_NewEnum]
End Function

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
    m_ListSubItems.Remove Index
End Sub
