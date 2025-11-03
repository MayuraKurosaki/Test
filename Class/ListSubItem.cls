VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSubItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ListView SubItem object"
Option Explicit

'ListView SubItem object

'--------------Fields----------------
Private Type Fields
    Index As Long
    Key As String
    ReportIcon As Variant
    Bold As Boolean
    ForeColor As OLE_COLOR
    Text As String
    Tag As Variant
    ToolTipText As String
End Type
Private This As Fields


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()

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

'Returns/sets the subitem's report view icon
Public Property Get ReportIcon() As Variant
Attribute ReportIcon.VB_Description = "Returns/sets the subitem's report view icon"
'Attribute ReportIcon.VB_Description = "Returns/sets the subitem's report view icon"
    ReportIcon = This.ReportIcon
End Property
Public Property Let ReportIcon(ByVal Value As Variant)
    This.ReportIcon = Value
End Property

'Returns/sets a value that determines whether the text for a ListSubItem object is displayed with a bold font.
Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text for a ListSubItem object is displayed with a bold font."
'Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text for a ListSubItem object is displayed with a bold font."
    Bold = This.Bold
End Property
Public Property Let Bold(ByVal Value As Boolean)
    This.Bold = Value
End Property

'Returns/sets the foreground color used to display text for a ListSubItem object.
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text for a ListSubItem object."
'Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text for a ListSubItem object."
    ForeColor = This.ForeColor
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    This.ForeColor = Value
End Property

'Returns/sets the text to display text for a ListSubItem object.
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text to display text for a ListSubItem object."
Attribute Text.VB_UserMemId = 0
'Attribute Text.VB_Description = "Returns/sets the text to display text for a ListSubItem object."
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

'Returns/sets the text for the subitem tooltip
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text for the subitem tooltip"
'Attribute ToolTipText.VB_Description = "Returns/sets the text for the subitem tooltip"
    ToolTipText = This.ToolTipText
End Property
Public Property Let ToolTipText(ByVal Value As String)
    This.ToolTipText = Value
End Property


'--------------Compatible Methods----------------

