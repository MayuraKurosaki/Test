VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An item in a ListView control that contains the index of icons associated with it, text, and an array of strings representing subitems that are displayed in Report view."
Option Explicit

'An item in a ListView control that contains the index of icons associated with it, text, and an array of strings representing subitems that are displayed in Report view.

'--------------Fields----------------
Private Type Fields
    Index As Long
    Key As String
    Checked As Boolean
    Selected As Boolean
    ListSubItems As ListSubItems
    Ghosted As Boolean
    Icon As Variant
    SmallIcon As Variant
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    Bold As Boolean
    ForeColor As OLE_COLOR
    Text As String
    Tag As Variant
    ToolTipText As String
End Type
Private This As Fields


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This.ListSubItems = New ListSubItems
End Sub

Private Sub Class_Terminate()
    Set This.ListSubItems = Nothing
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

'Returns/sets a value that determines whether a ListSubItem object is checked.
Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets a value that determines whether a ListSubItem object is checked."
'Attribute Checked.VB_Description = "Returns/sets a value that determines whether a ListSubItem object is checked."
    Checked = This.Checked
End Property
Public Property Let Checked(ByVal Value As Boolean)
    This.Checked = Value
End Property

'Returns/sets a value which determines if a ListItem or Node object is selected.
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value which determines if a ListItem or Node object is selected."
'Attribute Selected.VB_Description = "Returns/sets a value which determines if a ListItem or Node object is selected."
    Selected = This.Selected
End Property
Public Property Let Selected(ByVal Value As Boolean)
    This.Selected = Value
End Property

'Returns/sets an array of strings representing the ListIitem's data.
Public Property Get SubItems(ByVal Index As Long) As String
Attribute SubItems.VB_Description = "Returns/sets an array of strings representing the ListItem's data."
'Attribute SubItems.VB_Description = "Returns/sets an array of strings representing the ListItem's data."
    SubItems = This.ListSubItems.Item(Index)
End Property
Public Property Let SubItems(ByVal Index As Long, ByVal Value As String)
    This.ListSubItems.Item(Index).Text = Value
End Property

'Returns/sets a collection of ListSubItems belonging to the ListItem.
Public Property Get ListSubItems() As ListSubItems
Attribute ListSubItems.VB_Description = "Returns/sets a collection of ListSubItems belonging to the ListItem."
'Attribute ListSubItems.VB_Description = "Returns/sets a collection of ListSubItems belonging to the ListItem."
    Set ListSubItems = This.ListSubItems
End Property
Public Property Set ListSubItems(ByVal Value As ListSubItems)
    Set This.ListSubItems = Value
End Property

'Returns/sets the state of a ListView control Object item.
Public Property Get Ghosted() As Boolean
Attribute Ghosted.VB_Description = "Returns/sets the state of a ListView control Object item."
'Attribute Ghosted.VB_Description = "Returns/sets the state of a ListView control Object item."
    Ghosted = This.Ghosted
End Property
Public Property Let Ghosted(ByVal Value As Boolean)
    This.Ghosted = Value
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

'Returns/sets the index of an small icon in an associated ImageList control.
Public Property Get SmallIcon() As Variant
Attribute SmallIcon.VB_Description = "Returns/sets the index of an small icon in an associated ImageList control."
'Attribute SmallIcon.VB_Description = "Returns/sets the index of an small icon in an associated ImageList control."
    SmallIcon = This.SmallIcon
End Property
Public Property Let SmallIcon(ByVal Value As Variant)
    This.SmallIcon = Value
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

'Returns/sets the distance between the internal top edge of an object and the top edge of its container.
Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
'Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
    Top = This.Top
End Property
Public Property Let Top(ByVal Value As Single)
    This.Top = Value
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

'Returns/sets the height of an object.
Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
'Attribute Height.VB_Description = "Returns/sets the height of an object."
    Height = This.Height
End Property
Public Property Let Height(ByVal Value As Single)
    This.Height = Value
End Property

'Returns/sets a value that determines whether the text for a ListItem object is displayed with a bold font.
Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text for a ListItem object is displayed with a bold font."
'Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text for a ListItem object is displayed with a bold font."
    Bold = This.Bold
End Property
Public Property Let Bold(ByVal Value As Boolean)
    This.Bold = Value
End Property

'Returns/sets the foreground color used to display text and graphics in an object.
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
'Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = This.ForeColor
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    This.ForeColor = Value
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

Public Property Get TextPtr() As LongPtr
    TextPtr = StrPtr(This.Text)
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

'Returns/sets the text for the subitem tooltip
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text for the subitem tooltip"
'Attribute ToolTipText.VB_Description = "Returns/sets the text for the subitem tooltip"
    ToolTipText = This.ToolTipText
End Property
Public Property Let ToolTipText(ByVal Value As String)
    This.ToolTipText = Value
End Property


'--------------Extended Properties----------------
Public Property Get Self() As ListItem
    Set Self = Me
End Property


'--------------Compatible Methods----------------
