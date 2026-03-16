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

Sub GetCurrentFontName()
    Dim hdc As LongPtr
    Dim faceName As String * 255
    Dim result As Long
    
    ' アクティブなウィンドウのDCを取得する例
    ' ※実際には適切なhdcを取得する必要があります
    hdc = GetDC(0) ' 画面全体のDC（GetDCもAPI宣言が必要）
    
    ' フォント名を取得
    result = GetTextFace(hdc, 255, faceName)
    
    If result > 0 Then
        ' NULL文字を取り除いてフォント名を表示
        MsgBox "現在のフォント: " & Left(faceName, result)
    End If
    
    ' DCを解放
    ReleaseDC 0, hdc
End Sub

Sub 組み込みボックス()
    Application.Dialogs(xlDialogStandardFont).Show
End Sub
