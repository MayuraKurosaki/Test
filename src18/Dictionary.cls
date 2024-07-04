VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Drop-in replacement for Scripting.Dictionary on Mac\r\n\r\nDictionary v1.4.0\r\n(c) Tim Hall - https://github.com/timhall/VBA-Dictionary\r\nAuthor: tim.hall.engr@gmail.com\r\nLicense: MIT (http://www.opensource.org/licenses/mit-license.php)\r\n"
Option Explicit

' ---------------------------------------------
' Constants and Private Variables
' ---------------------------------------------

Private This As Object  ' Scripting.Dictionary

' ---------------------------------------------
' Types
' ---------------------------------------------

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' ---------------------------------------------
' Properties
' ---------------------------------------------

Public Property Get CompareMode() As CompareMethod
Attribute CompareMode.VB_Description = "Set or get the string comparison method."
'Attribute CompareMode.VB_Description = "Set or get the string comparison method."
    CompareMode = This.CompareMode
End Property

Public Property Let CompareMode(RHS As CompareMethod)
    This.CompareMode = RHS
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
'Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
    Count = This.Count
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_Description = "Set or get the item for a given key."
'Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(This.Item(Key)) Then
        Set Item = This.Item(Key)
    Else
        Item = This.Item(Key)
    End If
End Property

Public Property Let Item(Key As Variant, RHS As Variant)
    If VBA.IsObject(RHS) Then
        Set This.Item(Key) = RHS
    Else
        This.Item(Key) = RHS
    End If
End Property

Public Property Set Item(Key As Variant, RHS As Variant)
    Set This.Item(Key) = RHS
End Property

Public Property Let Key(prmPrevious As Variant, RHS As Variant)
Attribute Key.VB_Description = "Change a key to a different key."
'Attribute key.VB_Description = "Change a key to a different key."
    This.Key(prmPrevious) = RHS
End Property

' =============================================
' Public Methods
' =============================================

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.[_NewEnum]
End Function

' ---------------------------------------------
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' ---------------------------------------------
Public Sub Add(Key As Variant, Item As Variant)
Attribute Add.VB_Description = "Add a new key and item to the dictionary."
'Attribute Add.VB_Description = "Add a new key and item to the dictionary."
    This.Add Key, Item
End Sub

' ---------------------------------------------
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' ---------------------------------------------
Public Function Exists(Key As Variant) As Boolean
Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
'Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
    Exists = This.Exists(Key)
End Function

' ---------------------------------------------
' Get an array of all items
'
' @return {Variant}
' ---------------------------------------------
Public Function Items() As Variant
Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
'Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
    Items = This.Items
End Function

' ---------------------------------------------
' Get an array of all keys
'
' @return {Variant}
' ---------------------------------------------
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
'Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
    Keys = This.Keys
End Function

' ---------------------------------------------
' TryGetValue
'
' @param {Variant} Key
' @return {Boolean}
' ---------------------------------------------
Public Function TryGetValue(Key As Variant, ByRef value As Variant) As Boolean
    If Exists(Key) Then
        If VBA.IsObject(This.Item(Key)) Then
            Set value = This.Item(Key)
        Else
            value = This.Item(Key)
        End If
        TryGetValue = True
    Else
        If VBA.VarType(value) = vbObject Then
            Set value = Nothing
        Else
            value = Null
        End If
        TryGetValue = False
    End If
End Function

' ---------------------------------------------
' Remove an item for the given key
'
' @param {Variant} Key
' ---------------------------------------------
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Remove a given key from the dictionary."
'Attribute Remove.VB_Description = "Remove a given key from the dictionary."
    This.Remove Key
End Sub

' ---------------------------------------------
' Remove all items
' ---------------------------------------------
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
'Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
    This.RemoveAll
End Sub

' =============================================
' Private Functions
' =============================================

Private Sub Class_Initialize()
    Set This = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set This = Nothing
End Sub
