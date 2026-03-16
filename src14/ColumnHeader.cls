VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ListView Column Header"
Option Explicit

'ListView Column Header

'--------------Enums----------------
'ColumnAlignment constants
Public Enum ListColumnAlignmentConstants
    lvwColumnLeft = 0
    lvwColumnRight = 1
    lvwColumnCenter = 2
End Enum


'--------------Fields----------------
Private Type Fields
    Index As Long
    Key As String
    Position As Integer
    Alignment As ListColumnAlignmentConstants
    Icon As Variant
    SubItemIndex As Long
    Left As Single
    Width As Single
    Text As String
    Tag As Variant
End Type
Private This As Fields


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
End Sub

'--------------Compatible Properties----------------
'Returns/sets the index of an object in a collection. Read-only at run time.
Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns/sets the index of an object in a collection. Read-only at run time."
'Attribute Index.VB_Description = "Returns/sets the index of an object in a collection. Read-only at run time."
    Index = This.Index
End Property
Public Property Let Index(ByVal Value As Long)
    This.Index = Value
End Property

'Returns/sets the unique string of an object in a collection.
Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets the unique string of an object in a collection."
'Attribute Key.VB_Description = "Returns/sets the unique string of an object in a collection."
    Key = This.Key
End Property
Public Property Let Key(ByVal Value As String)
    This.Key = Value
End Property

'Returns current position of column.
Public Property Get Position() As Integer
Attribute Position.VB_Description = "Returns current position of column."
'Attribute Position.VB_Description = "Returns current position of column."
    Position = This.Position
End Property
Public Property Let Position(ByVal Value As Integer)
    This.Position = Value
End Property

'Returns/sets a value that determines where an object is displayed on a form.
Public Property Get Alignment() As ListColumnAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
'Attribute Alignment.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
    Alignment = This.Alignment
End Property
Public Property Let Alignment(ByVal Value As ListColumnAlignmentConstants)
    This.Alignment = Value
End Property

'Returns/sets the index of an icon in an associated ImageList control.
Public Property Get Icon() As Variant
Attribute Icon.VB_Description = "Returns/sets the index of an icon in an associated ImageList control."
'Attribute Icon.VB_Description = "Returns/sets the index of an icon in an associated ImageList control."
    Icon = This.Icon
End Property
Public Property Let Icon(ByVal Value As Variant)
    This.Icon = Value
End Property

'Returns the index of the subitem associated with a ColumnHeader object in a ListView control.
Public Property Get SubItemIndex() As Long
Attribute SubItemIndex.VB_Description = "Returns the index of the subitem associated with a ColumnHeader object in a ListView control."
'Attribute SubItemIndex.VB_Description = "Returns the index of the subitem associated with a ColumnHeader object in a ListView control."
    SubItemIndex = This.SubItemIndex
End Property
Public Property Let SubItemIndex(ByVal Value As Long)
    This.SubItemIndex = Value
End Property

'Returns/sets the distance between the internal left edge of an object and the left edge of its container.
Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
'Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
    Left = This.Left
End Property
Public Property Let Left(ByVal Value As Single)
    This.Left = Value
End Property

'Returns/sets the width of an object.
Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
'Attribute Width.VB_Description = "Returns/sets the width of an object."
    Width = This.Width
End Property
Public Property Let Width(ByVal Value As Single)
    This.Width = Value
End Property

'Returns/sets the text to be displayed in a control.
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text to be displayed in a control."
Attribute Text.VB_UserMemId = 0
'Attribute Text.VB_Description = "Returns/sets the text to be displayed in a control."
'Attribute Text.VB_UserMemId = 0
    Text = This.Text
End Property
Public Property Let Text(ByVal Value As String)
    This.Text = Value
End Property

'Stores any extra data needed for your program.
Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
'Attribute Tag.VB_Description = "Stores any extra data needed for your program."
    Tag = This.Tag
End Property
Public Property Let Tag(ByVal Value As Variant)
    This.Tag = Value
End Property


'--------------Extended Properties----------------
Public Property Get Self() As ColumnHeader
    Set Self = Me
End Property

'--------------Compatible Methods----------------
