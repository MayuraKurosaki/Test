VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "DataObject for clipboard operations"
Option Explicit

'DataObject for clipboard operations

'--------------Fields----------------
Private This As DataObjectFiles 'Collection of filenames for vbCFFiles format


'--------------Constructor and Destructor----------------
Private Sub Class_Initialize()
    Set This = New DataObjectFiles
End Sub

Private Sub Class_Terminate()
    Set This = Nothing
End Sub


'--------------Compatible Properties----------------
'A collection of filenames used by the vbCFFiles format.
Public Property Get Files() As DataObjectFiles
    Set Files = This
End Property


'--------------Compatible Methods----------------
'Clears all data and formats in a DataObject object.
Public Sub Clear()
    'No internal data to clear in this implementation.
End Sub

'Retrieves data of a specified format from a DataObject object.
Public Function GetData(ByVal Format As Integer) As Variant
    'This is a stub implementation. Actual data retrieval logic would go here.
    Dim Data As Variant
    'Return empty data for now.
    GetData = Data
End Function

'Determines if a specified clipboard format is supported by the DataObject object.
Public Function GetFormat(ByVal Format As Integer) As Boolean
    'This is a stub implementation. Actual format checking logic would go here.
    GetFormat = False
End Function

'Adds a supported format and possibly its data to a DataObject object.
Public Sub SetData(ByVal Data As Variant, ByVal Format As Integer)
    'This is a stub implementation. Actual data setting logic would go here.
End Sub
