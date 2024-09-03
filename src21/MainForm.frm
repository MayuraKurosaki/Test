VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7140
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'List View messages
Private Const LVM_FIRST = &H1000&
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

'Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr


Private searchCriteriaName As String
Private searchCriteriaAge As Long
Private searchCriteriaAddress As String
Private searchCriteriaSex As String
Private searchCriteriaBloodType As String
Private searchCriteriaDate As Date
Private searchCriteriaDateLevel As Long

Private originalTable As ListObject
Private workTable As ListObject

Private lvTop As Long
Private lvLeft As Long

'Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

'Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

Private Function CreateListView(hWndParent As LongPtr, iid As LongPtr, dwStyle As Long, dwExStyle As Long) As LongPtr
    Dim rc As RECT
    Dim hwndLV As LongPtr
    
    Call GetClientRect(hWndParent, rc)
'    hwndLV = CreateWindowEx(dwExStyle, WC_LISTVIEW, "", _
'                                                  dwStyle, 218, 2, 650, rc.Bottom - 30, _
'                                                  hWndParent, iid, App.hInstance, 0)
    hwndLV = CreateWindowEx(dwExStyle, WC_LISTVIEW, "", _
                                                  dwStyle, 218, 2, 650, rc.Bottom - 30, _
                                                  hWndParent, iid, 0, 0)
     ListView_SetItemCount hwndLV, UBound(VLItems) + 1
    CreateListView = hwndLV
End Function

Private Sub InitListView()
    Dim dwStyle As Long, dwStyle2 As Long
    Dim lvcol As LVCOLUMNW
    Dim i As Long
    Dim rc As RECT
    
    hLVVG = CreateListView(Me.hWnd, IDD_LISTVIEW, _
                      LVS_AUTOARRANGE Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_ALIGNTOP Or LVS_OWNERDATA Or _
                      WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, WS_EX_CLIENTEDGE)

    Call GetClientRect(Me.hWnd, rc)
    SetWindowPos hLVVG, 0, 200, 0, rc.Right - 200, rc.Bottom, 0
      
    Dim lvsex As LVStylesEx
    lvsex = LVS_EX_DOUBLEBUFFER Or LVS_EX_FULLROWSELECT
    
    Call ListView_SetExtendedStyle(hLVVG, lvsex)
    Dim swt1 As String
    Dim swt2 As String
    swt1 = "explorer"
    swt2 = ""
    Call SetWindowTheme(hLVVG, StrPtr(swt1), 0&)
    
    Dim iCurViewMode As Long
    iCurViewMode = LV_VIEW_DETAILS
    Call SendMessage(hLVVG, LVM_SETVIEW, iCurViewMode, ByVal 0&)
    
    ReDim sColText(1)
    sColText(0) = "Index"
    sColText(1) = "Name"
    
    lvcol.mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    lvcol.fmt = LVCFMT_CENTER
    lvcol.cchTextMax = Len(sColText(0))
    lvcol.pszText = StrPtr(sColText(0))
    lvcol.CX = 70
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 1, lvcol)

    lvcol.cchTextMax = Len(sColText(1))
    lvcol.pszText = StrPtr(sColText(1))
    lvcol.CX = 140
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 2, lvcol)
End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub Form_Activate()
    Dim i%, m%, x%
    Dim arrByte() As Byte
    Dim Guncode$
    Dim Sp1() As String, Sp2() As String
    
    arrByte = LoadResData(101, "CUSTOM")
    Guncode = ConvertedUTF8(arrByte)
    Guncode = Right$(Guncode, Len(Guncode) - 1)
    Sp1 = Split(Guncode, vbNewLine)
    m = UBound(Sp1)
    ReDim VLItems(m)
    
    For i = 0 To m
        ReDim VLItems(i).sSubItems(0)
        Sp2 = Split(Sp1(i), " ")
        VLItems(i).sText = Sp2(0)
        For x = 1 To UBound(Sp2): VLItems(i).sSubItems(0) = VLItems(i).sSubItems(0) & Sp2(x) & " ": Next x
        VLItems(i).sSubItems(0) = Left$(VLItems(i).sSubItems(0), Len(VLItems(i).sSubItems(0)) - 1)
    Next i
    
    Subclass2 Me.hWnd, AddressOf FGVWndProc
    InitListView
End Sub

Function ConvertedUTF8(ByRef Data() As Byte) As String
    Dim TotalBuffer() As Byte, Converted() As Byte, i As Long
    
    
    i = i + UBound(Data) + 1
    ReDim Preserve TotalBuffer(i - 1)
    RtlMoveMemory TotalBuffer(i - UBound(Data) - 1), Data(0), UBound(Data) + 1&
    
    Dim lSize As Long
    lSize = MultiByteToWideChar(CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, ByVal 0&, 0&)
    
    ReDim Converted(lSize * 2 - 1)
    MultiByteToWideChar CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, Converted(0), lSize
    ConvertedUTF8 = Converted
End Function







Private Sub CheckBoxFemale_AfterUpdate()
'    Call Filter
End Sub

'���F��:approver�@���F����:Approve�@���F:Approval
'����:signature
'����:constraint
'OperationProcedure
'Reason for operation
'Operation results
'TimeUnit
'�F��:authentication

Private Sub CheckBoxFemale_Change()
    If Me.CheckBoxFemale.Value Then searchCriteriaSex = "��"
    Call Filter
End Sub

Private Sub CheckBoxFemale_Click()

End Sub

Private Sub CheckBoxMale_AfterUpdate()
'    Call Filter
End Sub

Private Sub CheckBoxMale_Change()
    If Me.CheckBoxMale.Value Then searchCriteriaSex = "�j"
    Call Filter
End Sub

Private Sub CheckBoxMale_Click()

End Sub

Private Sub ComboBoxAddress_AfterUpdate()
'    Call Filter
End Sub

'Private Sub ComboBoxAddress_AfterUpdate()
'    Debug.Print Me.ComboBoxAddress.Text
'End Sub

Private Sub ComboBoxAddress_Change()
    searchCriteriaAddress = Me.ComboBoxAddress.Text
    Call Filter
End Sub

Private Sub ComboBoxAddress_DropButtonClick()
    Dim listRange As Range
    Set listRange = ThisWorkbook.Worksheets("List").ListObjects("T_�s���{��").ListColumns("�s���{����").DataBodyRange
    Dim i As Long
    With ComboBoxAddress
        For i = 1 To listRange.Rows.Count
            .AddItem listRange(i)
        Next
    End With
End Sub

Private Sub CommandButtonDatePicker_Click()
    Call DatePicker.Init
    Me.TextBoxDate.Text = DatePicker.SelectionDate 'Format(searchCriteriaDate, "YYYY/MM/DD")
End Sub

Private Sub ListBoxResultList_AfterUpdate()
    Debug.Print "ListBoxUpdated"
    Call FitListColumnWidthToText
End Sub

Private Sub ListBoxResultList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBoxResultList_Change()
    Debug.Print "ListBoxChanged"
    Call FitListColumnWidthToText
End Sub

Private Sub ListBoxResultList_Click()
    Debug.Print Me.ListBoxResultList.ListIndex
End Sub

Private Sub ListViewOrder_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub OptionButtonBloodTypeA_AfterUpdate()
'    Call Filter
End Sub

Private Sub OptionButtonBloodTypeB_Change()
    If Me.OptionButtonBloodTypeB.Value Then searchCriteriaBloodType = "B"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeB_Click()

End Sub

Private Sub OptionButtonBloodTypeA_Change()
    If Me.OptionButtonBloodTypeA.Value Then searchCriteriaBloodType = "A"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeA_Click()

End Sub

Private Sub OptionButtonBloodTypeAB_Change()
    If Me.OptionButtonBloodTypeAB.Value Then searchCriteriaBloodType = "AB"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeAB_Click()

End Sub

Private Sub OptionButtonBloodTypeO_Change()
    If Me.OptionButtonBloodTypeO.Value Then searchCriteriaBloodType = "O"
    Call Filter
End Sub

Private Sub OptionButtonBloodTypeO_Click()

End Sub

Private Sub OptionButtonFemale_Change()

End Sub

Private Sub OptionButtonFemale_Click()

End Sub

Private Sub OptionButtonMale_Change()

End Sub

Private Sub OptionButtonMale_Click()

End Sub

Private Sub TextBoxAge_AfterUpdate()
    If Me.TextBoxAge.Text = "" Then
        searchCriteriaAge = -1
    Else
        searchCriteriaAge = TextBoxAge.Text
    End If
    Call Filter
End Sub

Private Sub TextBoxAge_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxAge_Change()

End Sub

Private Sub TextBoxDate_AfterUpdate()
    Call Filter
End Sub

Private Sub TextBoxDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If VBA.IsDate(Me.TextBoxDate) Then
        searchCriteriaDate = Me.TextBoxDate.Value
        Me.TextBoxDate.Text = Format(searchCriteriaDate, "YYYY�NMM��DD��")
    Else
        If Me.TextBoxDate.Text <> "" Then
            Me.TextBoxDate.SelStart = 0
            Me.TextBoxDate.SelLength = VBA.Len(Me.TextBoxDate.Text)
            Cancel = True
        End If
    End If
End Sub

Private Sub TextBoxDate_Change()

End Sub

Private Sub TextBoxDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 187 And Shift = 2 Then TextBoxDate.Value = Format(Now, "YYYY�NMM��DD��") ' Ctrl + �u;�v
End Sub

Private Sub TextBoxName_AfterUpdate()
    searchCriteriaName = TextBoxName.Text
    Call Filter
End Sub

Private Sub TextBoxName_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBoxName_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    searchCriteriaAge = -1
    Call CopyTable
    With Me.ListBoxResultList
        .Clear
        .ColumnHeads = True
        .RowSource = workTable.DataBodyRange.Address
    End With
'    Call FitListColumnWidthToText
    Call ListViewInitialize
End Sub

Private Sub UserForm_Terminate()
    With ThisWorkbook.Worksheets("Temp")
            .Visible = True
    End With
    
    Set originalTable = Nothing
    Set workTable = Nothing
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Temp").Delete
    Application.DisplayAlerts = True
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub CopyTable()
    Application.ScreenUpdating = False
    With ThisWorkbook
        Dim sheetIndex As Long
        sheetIndex = .ActiveSheet.Index
        .Worksheets("Dummy").Copy After:=.Worksheets(.Worksheets.Count)
        .Worksheets(sheetIndex).Activate
        With .Worksheets(.Worksheets.Count)
            .Name = "Temp"
'            .Visible = False
        End With
    End With

    Set originalTable = ThisWorkbook.Worksheets("Dummy").ListObjects("T_Dummy")
    Set workTable = ThisWorkbook.Worksheets("Temp").ListObjects(1)
    Application.ScreenUpdating = True
End Sub

Private Sub ListViewInitialize()
'    With Me.ListViewOrder
'
'        .Clear
'        .ColumnHeads = True
'        .RowSource = workTable.DataBodyRange.Address
'    End With
'
'    With workTable
'        Dim maxColumn As Long
'        maxColumn = .ListColumns.Count
'    '    Dim widthArray() As Long
'        Dim widthArray() As String
'        ReDim widthArray(0 To maxColumn - 1)
'
'    Dim i As Long
'    Dim col As Long
'    Dim textString As String
'    Dim textWidth As Long
'    With Me.ListBoxResultList
'        For i = 0 To .ListCount - 1
'            For col = 0 To maxColumn - 1
'                textString = .List(i, col)
'                textWidth = MeasureTextWidth.MeasureTextWidth(textString, .Font.Name, .Font.size)
'                If textWidth > Val(widthArray(col)) Then widthArray(col) = CStr(textWidth)
'            Next col
'        Next i
'        .ColumnWidths = VBA.Join(widthArray, ";")
'    End With
    With ListViewOrder
        .Visible = False
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .OLEDropMode = ccOLEDropManual
        
        .ColumnHeaders.Add 1, "LineNo", "LN", 120, lvwColumnLeft
        .ColumnHeaders.Add 2, "Name", "����", 220, lvwColumnLeft
        .ColumnHeaders.Add 3, "Name2", "�����i�Ђ炪�ȁj", 0, lvwColumnLeft
        .ColumnHeaders.Add 4, "Age", "�N��", 0, lvwColumnCenter
        .ColumnHeaders.Add 5, "DateOfBirth", "���N����", 0, lvwColumnCenter
        .ColumnHeaders.Add 6, "Sex", "����", 0, lvwColumnCenter
        .ColumnHeaders.Add 7, "BloodType", "���t�^", 0, lvwColumnCenter
        .ColumnHeaders.Add 8, "EmailAddress", "���[���A�h���X", 0, lvwColumnLeft
        .ColumnHeaders.Add 9, "PhoneNumber", "�d�b�ԍ�", 0, lvwColumnCenter
        .ColumnHeaders.Add 10, "MobilePhoneNumber", "�g�ѓd�b�ԍ�", 0, lvwColumnCenter
        .ColumnHeaders.Add 11, "PostalCode", "�X�֔ԍ�", 0, lvwColumnCenter
        .ColumnHeaders.Add 12, "Address", "�Z��", 0, lvwColumnLeft
        .ColumnHeaders.Add 13, "CompanyName", "��Ж�", 0, lvwColumnLeft
        .ColumnHeaders.Add 14, "CreditCard", "�N���W�b�g�J�[�h", 0, lvwColumnCenter
        .ColumnHeaders.Add 15, "ExpirationDate", "�L������", 0, lvwColumnCenter
        .ColumnHeaders.Add 16, "IndividualNumber", "�}�C�i���o�[", 0, lvwColumnCenter
        
        Dim col As Long
        For col = 0 To .ColumnHeaders.Count - 1
            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, col, ByVal LVSCW_AUTOSIZE_USEHEADER
        Next col
'        .ColumnHeaders("SourceFilePath").Width = 0  ' �X�L�����t�@�C���̃t���p�X�͗�0�ŉB��
        lvTop = .Top
        lvLeft = .Left
        .Visible = True
    End With
End Sub

'Private Sub AutoFitListView()
'    With ListViewOrder
'        Dim col As Long
'        For col = 0 To .ColumnHeaders.Count - 1
'            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, col, ByVal LVSCW_AUTOSIZE_USEHEADER
''            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, col, ByVal LVSCW_AUTOSIZE
'        Next col
''        .ColumnHeaders("SourceFilePath").Width = 0  ' �X�L�����t�@�C���̃t���p�X�͗�0�ŉB��
'    End With
'End Sub

Private Sub Filter()
    Application.ScreenUpdating = False
    workTable.DataBodyRange.Delete
    With originalTable
        If Me.TextBoxName.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:="*" & searchCriteriaName & "*", VisibleDropDown:=False
        If Me.TextBoxAge.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�N��").Index, Criteria1:=">=" & searchCriteriaAge, VisibleDropDown:=False
        If Me.ComboBoxAddress.Value <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("�Z��").Index, Criteria1:=searchCriteriaAddress & "*", VisibleDropDown:=False
        If Me.CheckBoxFemale.Value Or Me.CheckBoxMale Then _
            .Range.AutoFilter Field:=.ListColumns("����").Index, Criteria1:=searchCriteriaSex, VisibleDropDown:=False
        If Me.OptionButtonBloodTypeA Or Me.OptionButtonBloodTypeB Or Me.OptionButtonBloodTypeAB Or Me.OptionButtonBloodTypeO Then _
            .Range.AutoFilter Field:=.ListColumns("���t�^").Index, Criteria1:=searchCriteriaBloodType, VisibleDropDown:=False
        If Me.TextBoxDate.Text <> "" Then _
            .Range.AutoFilter Field:=.ListColumns("���N����").Index, Criteria1:=Format(searchCriteriaDate, "YYYY�NMM��DD��"), VisibleDropDown:=False
    
'        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy ThisWorkbook.Worksheets("Temp").Range("A2")
        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy workTable.HeaderRowRange.Offset(1, 0)
        
    
    
'        Dim CellsCnt As Long    '���i�荞���ް��ٌ̾�
'        Dim ColCnt As Long      '��ð��ق̗�
'        Dim buf1 As Variant    '���e�[�u���S�̂̃f�[�^
'        buf1 = .Range
'        CellsCnt = .Range.SpecialCells(xlCellTypeVisible).Count
'        ColCnt = UBound(buf1, 2)
'
'        Dim buf2 As Variant    '���߂�l�Ƃ���ꎞ�I�Ȕz��
'        ReDim buf2(1 To (CellsCnt / ColCnt), 1 To ColCnt)
'
'        Dim i As Long            '�������ϐ��i�z��̍s�ʒu�j
'        Dim j As Long            '�������ϐ��i�z��̗�ʒu�j
'        Dim k As Long            '�e�[�u���̃f�[�^�s�{�^�C�g���s�̍s��
'        For k = 1 To UBound(buf1, 1)
'          If .Range.Rows(k).Hidden = False Then
'            i = i + 1
'            For j = 1 To ColCnt
'              buf2(i, j) = buf1(k, j)
'            Next j
'          End If
'        Next k
        
        '�I�[�g�t�B���^������
        .Range.AutoFilter
        .ShowAutoFilter = False
    End With
'    With Me.ListBoxResultList
'        .Clear
''        .ColumnHeads = True
'        .List = buf2
'    End With
'    Call FitListColumnWidthToText
    
'    Dim startTime As Single
'    Dim endTime As Single
'    startTime = VBA.Timer
'    With Me.ListViewOrder
'        .Visible = False
'        .ListItems.Clear
'        For i = 2 To UBound(buf2, 1)
''            If i > 20 Then Exit For
'            With .ListItems.Add '(Text:=buf2(i, 1))
'                If i > 20 Then GoTo CONTINUE
'                .Text = buf2(i, 1)
'                For j = 1 To UBound(buf2, 2) - 1
'                    .ListSubItems.Add index:=j, Text:=buf2(i, j + 1)
'                Next j
'            End With
'CONTINUE:
'        Next i
'        Dim col As Long
'        For col = 0 To .ColumnHeaders.Count - 1
'            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, col, ByVal LVSCW_AUTOSIZE_USEHEADER
'        Next col
'        .Visible = True
'        Me.Top = Me.Top + 1
'        Me.Top = Me.Top - 1
''        .Top = lvTop
''        .Left = lvLeft
'    End With

'    Call AutoFitListView
'    endTime = VBA.Timer
'    Debug.Print endTime - startTime
    Application.ScreenUpdating = True
End Sub

Private Sub FitListColumnWidthToText()
    Dim maxColumn As Long
    maxColumn = Me.ListBoxResultList.ColumnCount
'    Dim widthArray() As Long
    Dim widthArray() As String
    ReDim widthArray(0 To maxColumn - 1)
    
    Dim i As Long
    Dim col As Long
    Dim textString As String
    Dim textWidth As Long
    With Me.ListBoxResultList
        For i = 0 To .ListCount - 1
            For col = 0 To maxColumn - 1
                textString = .List(i, col)
                textWidth = MeasureTextWidth.MeasureTextWidth(textString, .Font.Name, .Font.size)
                If textWidth > Val(widthArray(col)) Then widthArray(col) = CStr(textWidth)
            Next col
        Next i
        .ColumnWidths = VBA.Join(widthArray, ";")
    End With
End Sub

