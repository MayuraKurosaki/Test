VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ListView Column Header collection
Private m_ColumnHeaders As Collection

'--------------Compatible Properties----------------
'Returns the number of objects in a collection.
Public Property Get Count() As Long
    Count = m_ColumnHeaders.Count
End Property

'Returns a specific member of a Collection object either by position or by key.
Public Property Get Item(Index As Variant) As ColumnHeader
    Set Item = m_ColumnHeaders.Item(Index)
End Property


'--------------Compatible Methods----------------
'Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time.
Public Function Add(Optional _
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
    m_ColumnHeaders.Add NewHeader
    Set Add = NewHeader
End Function

'Removes all objects in a collection.
Public Sub Clear()
    m_ColumnHeaders.Clear
End Sub

Public Function _NewEnum() As IUnknown
    Set _NewEnum = New EnumColumnHeaders
    Set _NewEnum.Collection = m_ColumnHeaders
End Function

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
    m_ColumnHeaders.Remove Index
End Sub