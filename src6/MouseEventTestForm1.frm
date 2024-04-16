VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MouseEventTestForm1 
   Caption         =   "MouseEventsForm"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "MouseEventTestForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MouseEventTestForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents Form As MouseEventForm 'MouseEventForm�N���X�̃I�u�W�F�N�g�ϐ��錾
Attribute Form.VB_VarHelpID = -1

Private Sub CommandButton1_Click()
    Unload Me
End Sub

'DropFiles�C�x���g�̋L�q��
'UserForm�Ƀh���b�O&�h���b�v���ꂽ�t�@�C�����擾����C�x���g
'���� DropFile�F�h���b�v���ꂽ�t���t�@�C����
Private Sub Form_DropFiles(ByVal DropFile As String)
    Debug.Print "Form_DropFiles"
    On Error Resume Next
    Static n As Long
    If InStr(TextBox1.Value, DropFile) Then Exit Sub
    n = n + 1
    TextBox1 = TextBox1 & n & ": " & DropFile & vbNewLine
End Sub

'MouseWheel�C�x���g�̋L�q��
'UserForm�ɂă}�E�X�z�C�[���̃X�N���[�����擾����C�x���g
'���� Control�FUserForm�̃A�N�e�B�u�R���g���[��
'�@�@ wParam�F����=Up�@����=Down
'�@�@ Shift�F1=Shift�L�[, 2=Ctrl�L�[, 4=Alt�L�[
Private Sub Form_MouseWheel(ByVal Control As MSForms.Control, ByVal wParam As LongPtr, ByVal Shift As Long)
    Debug.Print "Form_MouseWheel"
    On Error Resume Next
    Dim scroll As Long
    Const MINS = 3, MAXS = MINS * 4
    Select Case TypeName(Control)
    Case "ListBox", "ComboBox"
        scroll = IIf(Shift, MAXS, MINS)
        With Control
            If TypeOf Control Is MSForms.ComboBox Then .DropDown
            If 0 < wParam Then
                .TopIndex = IIf(.TopIndex < scroll, 0, .TopIndex - scroll)
            Else
                .TopIndex = .TopIndex + scroll
            End If
        End With
    End Select
End Sub

Private Sub UserForm_Activate()
    'MouseEventForm�N���X�̊J�n
    If Form Is Nothing Then
        Set Form = New MouseEventForm
        Form.Initialize Me
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 1 To 100
        ListBox1.AddItem i
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'MouseEventForm�N���X�̏I��
    Form.Terminate
    Set Form = Nothing
End Sub
