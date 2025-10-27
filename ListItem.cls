VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'An item in a ListView control that contains the index of icons associated with it, text, and an array of strings representing subitems that are displayed in Report view.

'--------------Compatible Properties----------------
'Returns/sets a value that determines whether the text for a ListItem object is displayed with a bold font.
Private m_Bold As Boolean
Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property
Public Property Let Bold(ByVal Value As Boolean)
    m_Bold = Value
End Property

'Returns/sets a value that determines whether a ListSubItem object is checked.
Private m_Checked As Boolean
Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property
Public Property Let Checked(ByVal Value As Boolean)
    m_Checked = Value
End Property

'Returns/sets the background color used to display text and graphics in an object.
Private m_BackColor As OLE_COLOR
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
End Property

'Returns/sets the state of a ListView control Object item.
Private m_State As Boolean
Public Property Get State() As Boolean
    State = m_State
End Property
Public Property Let State(ByVal Value As Boolean)
    m_State = Value
End Property

' Returns/sets the height of an object.
Private m_Height As Single
Public Property Get Height() As Single
    Height = m_Height
End Property
Public Property Let Height(ByVal Value As Single)
    m_Height = Value
End Property

'Returns/sets the index of an icon in an associated ImageList control.
Private m_Icon As Variant
Public Property Get Icon() As Variant
    Icon = m_Icon
End Property
Public Property Let Icon(ByVal Value As Variant)
    m_Icon = Value
End Property

'Returns/sets the index of an object in a collection. Read-only at run time.
Private m_Index As Long
Public Property Get Index() As Long
    Index = m_Index
End Property
Public Property Let Index(ByVal Value As Long)
    m_Index = Value
End Property

'Returns/sets the unique string of an object in a collection.
Private m_Key As String
Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Key(ByVal Value As String)
    m_Key = Value
End Property

'Returns/sets the distance between the internal left edge of an object and the left edge of its container.
Private m_Left As Single
Public Property Get Left() As Single
    Left = m_Left
End Property
Public Property Let Left(ByVal Value As Single)
    m_Left = Value
End Property








'--------------Compatible Methods----------------