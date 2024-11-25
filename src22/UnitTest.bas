Attribute VB_Name = "UnitTest"
Option Explicit
    
Public Sub Test()
    MainForm.Show
End Sub

Public Sub test2()
'    Dim delim() As Variant
'    delim = Array(vbCr, vbLf)
    Dim strTmp() As String
    
    strTmp = Util.SplitByChar("test" & vbCr & "123" & vbLf & "456" & vbCrLf & "789", vbCr, vbLf)
    Dim I As Long
    For I = LBound(strTmp) To UBound(strTmp)
        Debug.Print Len(strTmp(I)) & ":" & strTmp(I)
    Next I
    
    
End Sub

Sub AttributeTest(RHS As String)
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim ItemPair() As String
    ItemPair = Split(RHS, ";")
    Dim Pair As Variant
    Dim splitTmp() As String
    For Each Pair In ItemPair
        splitTmp = Split(Pair, ":")
        If UBound(splitTmp) > 0 Then
            dict.Add splitTmp(0), splitTmp(1)
        Else
            dict.Add splitTmp(0), ""
        End If
    Next Pair
    
    Dim Key As Variant
    For Each Key In dict.Keys
        Debug.Print Key & ":" & dict(Key)
    Next Key
    
    dict("Test") = "2"
    Dim Items As String
    For Each Key In dict.Keys
        Items = Items & Key & ":" & dict(Key) & ";"
    Next Key
    If Len(Items) > 0 Then
        Items = Left(Items, Len(Items) - 1)
    End If
    Debug.Print Items
    dict("Test") = "3"
    
    Debug.Print dict("Test")
    Set dict = Nothing
End Sub
