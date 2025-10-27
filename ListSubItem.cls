VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSubItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ListView SubItem object
Private m_Index As Long
Private m_Key As String
Private m_Text As String
Private m_Icon As Variant
'--------------Compatible Properties----------------
'Returns/sets a value that determines whether the text for a ListSubItem object is displayed with a bold font.
Private m_Bold As Boolean
Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property
Public Property Let Bold(ByVal Value As Boolean)
    m_Bold = Value
End Property

'Returns/sets the foreground color used to display text for a ListSubItem object.
Private m_ForeColor As OLE_COLOR
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
End Property

'Returns/sets the index of an object in a collection. Read-only at run time.
Public Property Get Index() As Long
    Index = m_Index
End Property
Public Property Let Index(ByVal Value As Long)
    m_Index = Value
End Property

'Returns/sets the unique string of an object in a collection.
Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Key(ByVal Value As String)
    m_Key = Value
End Property

'Returns/sets the subitem's report view icon
Public Property Get ReportIcon() As Variant
    ReportIcon = m_Icon
End Property
Public Property Let ReportIcon(ByVal Value As Variant)
    m_Icon = Value
End Property

' Stores any extra data needed for your program.
Private m_Tag As Variant
Public Property Get Tag() As Variant
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal Value As Variant)
    m_Tag = Value
End Property

'Returns/sets the text to display text for a ListSubItem object.
Public Property Get Text() As String
    Text = m_Text
End Property
Public Property Let Text(ByVal Value As String)
    m_Text = Value
End Property

' Returns/sets the text for the subitem tooltip
Private m_ToolTipText As String
Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property
Public Property Let ToolTipText(ByVal Value As String)
    m_ToolTipText = Value
End Property
'--------------Compatible Methods----------------

