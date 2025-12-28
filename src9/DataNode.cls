VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attribute VB_PredeclaredId = True

Option Explicit

Implements Initializable

'public class DataNode: IEnumerable<DataNode>

Public Enum NodeKind
    NK_UNKNOWN
    NK_OBJECT
    NK_ARRAY
    NK_STRING
    NK_NUMERIC
    NK_BOOLEAN
    NK_DATETIME
    NK_ROOT
    NK_NULL
End Enum

Private ready_ As Boolean
Private children_ As Collection     ' List<DataNode>
'Private children_ As Dictionary     ' List<DataNode>
Private parent_ As DataNode
Private childCount_ As Long
Private name_ As String
Private value_ As String
Private kind_ As NodeKind
Private epochTicks As Long
Private depth_ As Long      'データ階層の深さ
'Private nodeData_ As Variant

'Private iterator As Collection

' private static readonly long epochTicks = new DateTime(1970, 1, 1).Ticks;
Private Sub Class_Initialize()
    Set children_ = New Collection
    Set parent_ = Nothing
'    Set children_ = New Dictionary
    epochTicks = VBA.Int(VBA.CDbl(VBA.DateTime.DateSerial(1970, 1, 1)))
End Sub

Private Sub Class_Terminate()
    Set children_ = Nothing
End Sub

Private Sub ReadyCheck()
    If Not ready_ Then VBA.Err.Raise ErrCode.NotReady, _
        VBA.TypeName(Me), "クラス " & VBA.TypeName(Me) & " は初期化されていません。"
End Sub

' p(0):Kind / p(1):Name / p(2):Value
Private Function Initializable_CreateInstance(ByRef p() As Variant) As Object
    If ready_ Then
        VBA.Err.Raise ErrCode.InitTwice, VBA.TypeName(Me), _
        "クラス " & VBA.TypeName(Me) & " はすでに初期化されています。"
    End If
    If VBA.VarType(p) < vbArray Then
        Err.Raise ErrCode.InitTwice, TypeName(Me), _
        "クラス " & TypeName(Me) & " パラメータ。"
    End If
    
    ' 第1パラメータ:NodeKind
    If UBound(p) = -1 Then
        '空配列の時
        Kind = NK_ROOT
        Set Parent = Nothing
    Else
        Kind = p(0)
    End If
    
    ' 第3パラメータ:Value:オプション
    If UBound(p) > 1 Then
        If VBA.VarType(p(2)) = VBA.vbString Then Value = p(2)
    Else
        Value = ""
    End If
    
    ' 第2パラメータ:Name:オプション
    If UBound(p) > 0 Then
        If VBA.VarType(p(1)) = VBA.vbString Then Name = p(1)
    Else
        Name = ""
    End If
       
'    ' 第2パラメータ:Parent:オプション
'    If UBound(p) > 0 Then
'        If VBA.TypeName(p(1)) = "DataNode" Then Set Parent = p(1)
'    Else
'        Set Parent = Nothing
'    End If
    
    ready_ = True
    
    Set Initializable_CreateInstance = Me
End Function

'子ノードのコレクション
Public Property Get Children() As Collection    ' List<DataNode>
'Public Property Get Children() As Dictionary    ' List<DataNode>
    Set Children = children_
End Property

Public Property Get Parent() As DataNode
    Set Parent = parent_
End Property

Public Property Set Parent(aNode As DataNode)
    Set parent_ = aNode
End Property

Public Property Get ChildCount() As Long
    ChildCount = children_.Count
End Property

Public Property Get HasChildren() As Boolean
    HasChildren = (children_.Count > 0)
End Property

Public Property Get Name() As String
    Name = name_
End Property

Public Property Let Name(aName As String)
    name_ = aName
End Property

Public Property Get Value() As String
    Value = value_
End Property

Public Property Let Value(aValue As String)
    value_ = aValue
End Property

Public Property Get Kind() As NodeKind
    Kind = kind_
End Property

Public Property Let Kind(aKind As NodeKind)
    kind_ = aKind
End Property

Public Property Get Depth() As Long
    Depth = 0
    If Not (Parent Is Nothing) Then Depth = Parent.Depth + 1
End Property

'Public Function GetEnumerator() As Collection   ' IEnumerator<DataNode>
'    Set GetEnumerator = children_
'End Function
'        public IEnumerator<DataNode> GetEnumerator()
'        {
'            return _children.GetEnumerator();
'        }
'
'        IEnumerator IEnumerable.GetEnumerator()
'        {
'            return _children.GetEnumerator();
'        }

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
'    Set NewEnum = iterator.[_NewEnum]
    Set NewEnum = children_.[_NewEnum]
End Function

' aNameOrIndex: String Name / Long Index
Public Property Get Item(aNameOrIndex As Variant) As DataNode
Attribute Item.VB_UserMemId = 0
'Attribute token.VB_UserMemId = 0
'    token = iterator.Item(pIndex)
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set Item = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set Item = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set Item = Nothing
    End Select
End Property

'        public DataNode this[string name]
'        {
'            get { return GetNode(name); }
'        }
'
'        public DataNode this[int index]
'        {
'            get { return GetNodeByIndex(index); }
'        }

Public Static Function CreateStringNode(Optional aName As String = "") As DataNode
    Set CreateStringNode = CreateInstance(New DataNode, NK_STRING, aName)
End Function

Public Static Function CreateNumberNode(Optional aName As String = "") As DataNode
    Set CreateNumberNode = CreateInstance(New DataNode, NK_NUMERIC, aName)
End Function

Public Static Function CreateObjectNode(Optional aName As String = "") As DataNode
    Set CreateObjectNode = CreateInstance(New DataNode, NK_OBJECT, aName)
End Function

Public Static Function CreateArrayNode(Optional aName As String = "") As DataNode
'    Set CreateArrayNode = New DataNode
'    CreateArrayNode.Kind = NodeKind.NK_ARRAY
'    CreateArrayNode.Name = aName
    
    Set CreateArrayNode = CreateInstance(New DataNode, NK_ARRAY, aName)
End Function

Public Static Function CreateValueNode(aValue As Object) As DataNode
    Dim nodeKind_ As NodeKind
    Dim val_ As String
    val_ = ConvertValue(aValue, nodeKind_)
    Set CreateValueNode = CreateInstance(New DataNode, nodeKind_, "", val_)
End Function

Public Function ToDictionary() As Dictionary
    Set ToDictionary = Nothing
    
'    Dim node_ As DataNode
'
'    Select Case VBA.VarType(aNameOrLine)
'        Case VBA.vbString
'            Set node_ = Me.GetNode(VBA.CStr(aNameOrLine))
'        Case VBA.vbInteger, VBA.vbLong
'            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrLine))
'        Case Else
'            Set node_ = Nothing
'    End Select
'
'    If node_ Is Nothing Then Exit Function
    
    Set ToDictionary = New Dictionary
    
    Dim child_ As Object    'DataNode
    For Each child_ In Me.Children
        Call ToDictionary.Add(child_.Name, child_.Value)
    Next child_
    
    Set child_ = Nothing
'    Set node_ = Nothing
End Function

Public Function ToString() As String
    If Me.ChildCount = 0 And Me.Value <> "" Then ToString = Me.Value: Exit Function
    If Me.Name <> "" Then ToString = Me.Name: Exit Function
    If Me.Parent Is Nothing Then ToString = "[Root]": Exit Function
    ToString = "[Null]"
End Function

Public Function AddNode(aNode As DataNode) As DataNode
    If aNode Is Nothing Then Set AddNode = Nothing: Exit Function
    
    Call Me.Children.Add(aNode)
    Set aNode.Parent = Me
    Set AddNode = aNode
End Function

Public Function AddValue(aValue As Variant) As DataNode
    Set AddValue = AddField("", aValue)
End Function

Public Function AddField(aName As String, aValue As Variant) As DataNode
    If Me.Kind <> NK_ARRAY And Me.Kind <> NK_OBJECT Then
        Err.Raise ErrCode.NodeIsNotObject, TypeName(Me), "The kind of this node is not 'object'!"
    End If
    If VBA.TypeName(aValue) = "DataNode" Then
'        Call Me.AddNode(aValue)
'        Set AddField = aValue
'        Exit Function
        Err.Raise ErrCode.CannotAddNodeAsField, TypeName(Me), "Cannot add a node as a field!"
    End If

    Dim nodeKind_ As NodeKind
    Dim val_ As String
    val_ = ConvertValue(aValue, nodeKind_)
    Dim child_ As DataNode
    Set child_ = CreateInstance(New DataNode, nodeKind_, aName, val_)
    Call Me.AddNode(child_)

    Set AddField = child_
End Function


'public DataNode AddField(string name, object value)
'{
'    if (this.Kind != NodeKind.Array && this.Kind != NodeKind.Object)
'    {
'        throw new Exception("The kind of this node is not 'object'!");
'    }
'
'    if (value is DataNode)
'    {
'        throw new Exception("Cannot add a node as a field!");
'    }
'
'    NodeKind kind;
'    string val = ConvertValue(value, out kind);
'
'    var child = new DataNode(kind, name, val);
'    this.AddNode(child);
'    return child;
'}

Private Static Function ConvertValue(aValue As Variant, ByRef aKind As NodeKind) As String
    If VBA.Information.VarType(aValue) = VBA.vbObject Then
        If aValue Is Nothing Then
            aKind = NK_NULL
            ConvertValue = ""
            Exit Function
        End If
    End If
    
    Dim val_ As String

    Select Case VBA.Information.VarType(aValue)
        Case vbEmpty, vbNull, vbError
            ConvertValue = ""
            aKind = NK_UNKNOWN
            Exit Function
        Case VBA.vbInteger To VBA.vbCurrency, VBA.vbDecimal, VBA.vbByte
            val_ = VBA.CStr(aValue)
            aKind = NK_NUMERIC
#If Win64 Then
        Case vbLongLong         ' 20  LongLong 整数 (64ビットでのみ有効)
            val_ = VBA.CStr(aValue)
            aKind = NK_NUMERIC
#End If
        Case vbBoolean          ' 11  Boolean
            val_ = VBA.IIf(aValue, "True", "False")
            aKind = NK_BOOLEAN
        Case vbString
            val_ = aValue
            aKind = NK_STRING
        Case vbObject, vbDataObject
            aKind = NK_OBJECT
'        Case vbUserDefinedType
        Case vbDate             ' 7   Date
            val_ = VBA.CStr(aValue)
            Kind = NK_DATETIME
        Case Else
            If VBA.IsArray(aValue) Then
                aKind = NK_ARRAY
            Else
                aKind = NK_UNKNOWN
            End If
    End Select
    ConvertValue = val_
End Function

'        private static string ConvertValue(object value, out NodeKind kind)
'        {
'            if (value == null)
'            {
'                kind = NodeKind.Null;
'                return "";
'            }
'
'            string val;
'
'            if (value is int)
'            {
'                val = ((int)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is uint)
'            {
'                val = ((uint)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is long)
'            {
'                val = ((long)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is ulong)
'            {
'                val = ((ulong)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is byte)
'            {
'                val = ((byte)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is sbyte)
'            {
'                val = ((sbyte)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is short)
'            {
'                val = ((short)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is ushort)
'            {
'                val = ((ushort)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is float)
'            {
'                val = ((float)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is double)
'            {
'                val = ((double)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            if (value is decimal)
'            {
'                val = ((decimal)value).ToString(CultureInfo.InvariantCulture);
'                kind = NodeKind.Numeric;
'            }
'            Else
'            {
'                val = value.ToString();
'                kind = NodeKind.String;
'            }
'
'            return val;
'        }
'
Public Function HasNode(aName As String, Optional aIndex As Long = 0) As Boolean
    HasNode = Not (GetNode(aName, aIndex) Is Nothing)
End Function

' internal auxiliary
Public Function FindNode_(aName As String, aDepth As Long, aMaxDepth As Long) As DataNode
    Set FindNode_ = Nothing
    
    If VBA.StrComp(VBA.LCase$(Me.Name), VBA.LCase$(aName), vbBinaryCompare) = 0 Then Set FindNode_ = Me: Exit Function
    If aDepth >= aMaxDepth Then Exit Function
    
    Dim child_ As DataNode
    For Each child_ In children_
        Set FindNode_ = child_.FindNode_(aName, aDepth + 1, aMaxDepth)
        If Not (FindNode_ Is Nothing) Then Exit For
    Next
    Set child_ = Nothing
End Function

Public Function FindNode(aName As String, Optional aDepth As Long = 0, Optional aMaxDepth As Long = 0) As DataNode
    If aMaxDepth > 0 Then
        Set FindNode = FindNode_(aName, 0, aMaxDepth)
    Else
        Set FindNode = FindNode_(aName, 0, Depth)
    End If
End Function

Public Function GetNode(aName As String, Optional aIndex As Long = 0) As DataNode
    Set GetNode = Nothing
    
    Dim n As Long: n = 0
    
    Dim child_ As DataNode
    For Each child_ In children_
        If (VBA.StrComp(VBA.UCase$(child_.Name), VBA.UCase$(aName), vbBinaryCompare) = 0) Then
            If (n >= aIndex) Then
                Set GetNode = child_
                Exit For
            Else
                n = n + 1
            End If
        End If
    Next
    
    Set child_ = Nothing
End Function

Public Function GetNodeByIndex(aIndex As Long) As DataNode
    Set GetNodeByIndex = Nothing
    If (aIndex < 0) Or (aIndex >= children_.Count) Then Exit Function
    Set GetNodeByIndex = children_(aIndex + 1)
End Function

'--- Excel 64bit Version Only---------------------------
#If Win64 Then
' INT64
Public Function AsInt64(Optional defaultValue As LongLong = 0) As LongLong
    Dim result As LongLong
    result = defaultValue
    On Error Resume Next
    result = VBA.CLngLng(Me.Value)
    If VBA.Err.Number = 0 Then
        AsInt64 = result
    Else
        AsInt64 = defaultValue
    End If
End Function

Public Function GetInt64(aNameOrIndex As Variant, Optional defaultValue As LongLong = 0) As LongLong
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetInt64 = node_.AsInt64(defaultValue)
    Else
        GetInt64 = defaultValue
    End If
End Function
#Else
'-------------------------------------------------------

Public Function AsInt64(Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim result As Variant
    result = VBA.CDec(defaultValue)
    On Error Resume Next
    result = VBA.CDec(Me.Value)
    If VBA.Err.Number = 0 Then
        AsInt64 = result
    Else
        AsInt64 = VBA.CDec(defaultValue)
    End If
End Function

Public Function GetInt64(aNameOrIndex As Variant, Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetInt64 = node_.AsInt64(defaultValue)
    Else
        GetInt64 = VBA.CDec(defaultValue)
    End If
End Function

#End If

' UINT64
Public Function AsUInt64(Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim result As Variant
    result = VBA.CDec(defaultValue)
    On Error Resume Next
    result = VBA.CDec(Me.Value)
    If VBA.Err.Number = 0 Then
        AsUInt64 = result
    Else
        AsUInt64 = VBA.CDec(defaultValue)
    End If
End Function

Public Function GetUInt64(aNameOrIndex As Variant, Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetUInt64 = node_.AsUInt64(defaultValue)
    Else
        GetUInt64 = VBA.CDec(defaultValue)
    End If
End Function

' INT32
Public Function AsInt32(Optional defaultValue As Long = 0) As Long
    Dim result As Long
    result = defaultValue
    On Error Resume Next
    result = VBA.CLng(Me.Value)
    If VBA.Err.Number = 0 Then
        AsInt32 = result
    Else
        AsInt32 = defaultValue
    End If
End Function

Public Function GetInt32(aNameOrIndex As Variant, Optional defaultValue As Long = 0) As Long
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetInt32 = node_.AsInt32(defaultValue)
    Else
        GetInt32 = defaultValue
    End If
End Function

' UINT32
Public Function AsUInt32(Optional defaultValue As Currency = 0) As Currency
    Dim result As Currency
    result = defaultValue
    On Error Resume Next
    result = VBA.CCur(Me.Value)
    If VBA.Err.Number = 0 Then
        AsUInt32 = result
    Else
        AsUInt32 = defaultValue
    End If
End Function

Public Function GetUInt32(aNameOrIndex As Variant, Optional defaultValue As Currency = 0) As Currency
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetUInt32 = node_.AsUInt32(defaultValue)
    Else
        GetUInt32 = defaultValue
    End If
End Function

' BYTE
Public Function AsByte(Optional defaultValue As Byte = 0) As Byte
    Dim result As Byte
    result = defaultValue
    On Error Resume Next
    result = VBA.CByte(Me.Value)
    If VBA.Err.Number = 0 Then
        AsByte = result
    Else
        AsByte = defaultValue
    End If
End Function

Public Function GetByte(aNameOrIndex As Variant, Optional defaultValue As Byte = 0) As Byte
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetByte = node_.AsByte(defaultValue)
    Else
        GetByte = defaultValue
    End If
End Function

' SBYTE
Public Function AsSByte(Optional defaultValue As Integer = 0) As Integer
    Dim result As Integer
    result = defaultValue
    On Error Resume Next
    result = VBA.CInt(Me.Value)
    If VBA.Err.Number = 0 Then
        AsSByte = result
    Else
        AsSByte = defaultValue
    End If
End Function

Public Function GetSByte(aNameOrIndex As Variant, Optional defaultValue As Integer = 0) As Integer
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetSByte = node_.AsSByte(defaultValue)
    Else
        GetSByte = defaultValue
    End If
End Function


' ENUM
'        public T AsEnum<T>(T defaultValue = default(T)) where T : IConvertible
'        {
'            try
'            {
'                return (T)Enum.Parse(typeof(T), this.Value, /* ignorecase */ true);
'            }
'            catch (Exception)
'            {
'                int result = 0;
'                if (int.TryParse(this.Value, out result))
'                {
'                    return (T)(object)result;
'                }
'            }
'
'            return defaultValue;
'        }
'
'        public T GetEnum<T>(string name, T defaultValue = default(T)) where T : IConvertible
'        {
'            DataNode node = this.GetNode(name);
'            if (node != null)
'            {
'                return node.AsEnum<T>(defaultValue);
'            }
'
'            return defaultValue;
'        }
'
'        public T GetEnum<T>(int index, T defaultValue = default(T)) where T : IConvertible
'        {
'            DataNode node = this.GetNodeByIndex(index);
'            if (node != null)
'            {
'                return node.AsEnum<T>(defaultValue);
'            }
'
'            return defaultValue;
'        }


' BOOL
Public Function AsBool(Optional defaultValue As Boolean = False) As Boolean
    If (Me.Value = "1" Or VBA.StrComp(VBA.LCase$(Me.Value), "true", vbBinaryCompare)) Then AsBool = True: Exit Function
    If (Me.Value = "0" Or VBA.StrComp(VBA.LCase$(Me.Value), "false", vbBinaryCompare)) Then AsBool = False: Exit Function
    AsBool = defaultValue
End Function

Public Function GetBool(aNameOrIndex As Variant, Optional defaultValue As Boolean = False) As Boolean
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetBool = node_.AsBool(defaultValue)
    Else
        GetBool = defaultValue
    End If
End Function

' FLOAT
Public Function AsFloat(Optional defaultValue As Single = 0) As Single
    Dim result As Single
    result = defaultValue
    On Error Resume Next
    result = VBA.CSng(Me.Value)
    If VBA.Err.Number = 0 Then
        AsFloat = result
    Else
        AsFloat = defaultValue
    End If
End Function

Public Function GetFloat(aNameOrIndex As Variant, Optional defaultValue As Single = 0) As Single
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetFloat = node_.AsFloat(defaultValue)
    Else
        GetFloat = defaultValue
    End If
End Function

' DOUBLE
Public Function AsDouble(Optional defaultValue As Double = 0) As Double
    Dim result As Double
    result = defaultValue
    On Error Resume Next
    result = VBA.CDbl(Me.Value)
    If VBA.Err.Number = 0 Then
        AsDouble = result
    Else
        AsDouble = defaultValue
    End If
End Function

Public Function GetDouble(aNameOrIndex As Variant, Optional defaultValue As Double = 0) As Double
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetDouble = node_.AsDouble(defaultValue)
    Else
        GetDouble = defaultValue
    End If
End Function

' DECIMAL
Public Function AsDecimal(Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim result As Variant
    result = VBA.CDec(defaultValue)
    On Error Resume Next
    result = VBA.CDec(Me.Value)
    If VBA.Err.Number = 0 Then
        AsDecimal = result
    Else
        AsDecimal = VBA.CDec(defaultValue)
    End If
End Function

Public Function GetDecimal(aNameOrIndex As Variant, Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = VBA.CDec(0)
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetDecimal = node_.AsDecimal(VBA.CDec(defaultValue))
    Else
        GetDecimal = VBA.CDec(defaultValue)
    End If
End Function

' STRING
Public Function AsString(Optional defaultValue As String = "") As String
    If Me.Value <> "" Then
        AsString = Me.Value
    Else
        AsString = defaultValue
    End If
End Function

Public Function GetString(aNameOrIndex As Variant, Optional defaultValue As String = "") As String
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetString = node_.Value
    Else
        GetString = defaultValue
    End If
End Function

' DATETIME
Public Function AsDateTime(Optional defaultValue As Date = "1899/12/30 00:00:00") As Date   ' 1899/12/30 00:00:00 Date型初期値
    Dim result As Date
    result = defaultValue
    On Error Resume Next
    result = VBA.CDate(Me.Value)
    If VBA.Err.Number = 0 Then
        AsDateTime = result
    Else
        AsDateTime = defaultValue
    End If
End Function

Public Function GetDateTime(aNameOrIndex As Variant, Optional defaultValue As Date = "1899/12/30 00:00:00") As Date
    Dim node_ As DataNode
    
    Select Case VBA.VarType(aNameOrIndex)
        Case VBA.vbString
            Set node_ = Me.GetNode(VBA.CStr(aNameOrIndex))
        Case VBA.vbInteger, VBA.vbLong
            Set node_ = Me.GetNodeByIndex(VBA.CLng(aNameOrIndex))
        Case Else
            Set node_ = Nothing
    End Select
    
    If Not (node_ Is Nothing) Then
        GetDateTime = node_.AsDateTime(defaultValue)
    Else
        GetDateTime = defaultValue
    End If
End Function

