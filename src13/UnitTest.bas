Attribute VB_Name = "UnitTest"
Option Explicit

Sub Test()
    Dim TestList As List
    Set TestList = New List
    
    With TestList
        Debug.Print .Count
        Debug.Print .Capacity
        .Add "Item1", "Key1"
        .Add "Item2", "Key2"
        Debug.Print .Count
        Debug.Print .Capacity
        Debug.Print "Item(0):" & .Item(0)
        Debug.Print "Item(""Key2""):" & .Item("Key2")
    End With
        
    Set TestList = Nothing
End Sub
