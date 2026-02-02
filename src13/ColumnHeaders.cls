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
'Private Type Field
'    Items As List
'    Keys As List
'    Count As Long
'    Capacity As Long
'End Type

'Private This As Field

'Private This As Collection   'Collection of ColumnHeader objects
'Private This As ObjectList   'Collection of ColumnHeader objects
Private This As List   'Collection of ColumnHeader objects


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
'    Set This = New Collection
'    Set This = New ObjectList
    Set This = New List
'    Set This.Items = New List
'    Set This.Keys = New List
'
'    This.Capacity = 0
'    ReDim This.Items(This.Capacity)
End Sub

Private Sub Class_Terminate()
'    Call Clear
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
        err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
    Else
        If Index < 0 Or Index > This.Count - 1 Then
            err.Raise vbObjectError + 1, "DataObjectFiles", "Index out of bounds"
        Else
'            Set Item = This.GetItmeOfIndex(CLng(Index))
            Set Item = This.Item(CLng(Index))
        End If
    End If
End Property

Public Property Get Items() As ColumnHeader()
    Items = This.Items
End Property

'--------------Compatible Methods----------------
'Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time.
Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Width As Single, Optional ByVal Alignment As ListColumnAlignmentConstants, Optional ByVal Icon As Variant) As ColumnHeader
Attribute Add.VB_Description = "Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time."
'Attribute Add.VB_Description = "Adds a ColumnHeader object to a ColumnHeaders collection at both design time and run time."
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
    
'    If This.Count > This.Capacity Then
'        This.Capacity = (This.Capacity + 1) * 2 - 1
'        ReDim Preserve This.Items(This.Capacity)
'    End If
'
'    Set This.Items(This.Count) = NewHeader
'    This.Count = This.Count + 1
    
    This.Add NewHeader, Key
    Set Add = NewHeader
End Function

''Removes all objects in a collection.
'Public Sub Clear()
''Attribute Clear.VB_Description = "Removes all objects in a collection."
''    Set This = Nothing
''    This.Clear
'    This.Count = 0
'    This.Capacity = 0
'    ReDim This.Items(This.Capacity)
'End Sub

'Returns an enumerator for iterating through the ColumnHeaders collection.
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_Description = "Returns an enumerator for iterating through the ColumnHeaders collection."
'Attribute NewEnum.VB_UserMemId = -4
'    Set NewEnum = This.[_NewEnum]
End Function

'Removes a specific member from a collection.
Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a specific member from a collection."
'Attribute Remove.VB_Description = "Removes a specific member from a collection."
    This.Remove Index
    
'    If This.Count = 0 Then Exit Function
'    If Index < LBound(This.Items) Or UBound(This.Items) < Index Then Exit Function
'    Dim i As Long
'    For i = Index To This.Count - 1
'        This.Items(i) = This.Items(i + 1)
'    Next i
'
'    This.Count = This.Count - 1
'    Set This.Items(This.Count) = Nothing
'
'    Set Remove = Me
End Sub
