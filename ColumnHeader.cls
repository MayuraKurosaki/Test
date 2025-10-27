ERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ListView Column Header

'--------------Enums----------------
' ColumnAlignment constants
Public Enum ListColumnAlignmentConstants
    lvwColumnLeft = 0
    lvwColumnRight = 1
    lvwColumnCenter = 2
End Enum


'--------------Compatible Properties----------------
'Returns/sets a value that determines where an object is displayed on a form.
Private m_ZOrder As Long

Public Property Get ZOrder() As Long
    ZOrder = m_ZOrder
End Property
Public Property Let ZOrder(ByVal Value As Long)
    m_ZOrder = Value
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

'Returns current position of column.
Private m_Position As Integer
Public Property Get Position() As Integer
    Position = m_Position
End Property
Public Property Let Position(ByVal Value As Integer)
    m_Position = Value
End Property

'Returns the index of the subitem associated with a ColumnHeader object in a ListView control.
Private m_SubItemIndex As Long
Public Property Get SubItemIndex() As Long
    SubItemIndex = m_SubItemIndex
End Property
Public Property Let SubItemIndex(ByVal Value As Long)
    m_SubItemIndex = Value
End Property

'Stores any extra data needed for your program.
Private m_Tag As Variant
Public Property Get Tag() As Variant
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal Value As Variant)
    m_Tag = Value
End Property

'Returns/sets the text to be displayed in a control.
Private m_Text As String
Public Property Get Text() As String
    Text = m_Text
End Property
Public Property Let Text(ByVal Value As String)
    m_Text = Value
End Property

'Returns/sets the width of an object.
Private m_Width As Single
Public Property Get Width() As Single
    Width = m_Width
End Property
Public Property Let Width(ByVal Value As Single)
    m_Width = Value
End Property


'--------------Compatible Methods----------------
