VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuForm 
   Caption         =   "MenuForm"
   ClientHeight    =   4770
   ClientLeft      =   13050
   ClientTop       =   8595
   ClientWidth     =   2160
   OleObjectBlob   =   "MenuForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ��
    �����ʒu = 5
End Enum
Private Enum �s
    Left = 2
    Top = 3
End Enum
Private Sub Button_��]�V�t�g_Click()
    StaffForm.Show
End Sub

Private Sub Button_�r��_Click()
    Call WorkSheetRuler.ruleLine
End Sub

Private Sub Button_���T�x_Click()
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetWriter.WriteNumOfPublicHoliday(�X�^�b�t)
    Next
End Sub

Private Sub Button_�o�Εs��_Click()
    '�S���̏o�Εs���ɐF��h��
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetPainter.paintImpossibleDay(�X�^�b�t)
    Next
End Sub

Private Sub Button_�F�Z�b�g_Click()
    '�y���ɐF��h��
    Call WorkSheetPainter.SetWeekendColor
End Sub

Private Sub Button_�\��_Click()
    Call WorkSheetWriter.WriteLegalHoliday(TableManager.�j��)
    Call WorkSheetWriter.WriteMeetingDay(TableManager.��c��)
End Sub

Private Sub Button_�J������_Click()
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetWriter.WriteWorkTime(�X�^�b�t)
    Next
End Sub

Private Sub Button_��{�V�t�g_Click()
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetWriter.WriteBasicShift(�X�^�b�t)
        Call WorkSheetWriter.CopyFromPreviousMonth(�X�^�b�t)
    Next
End Sub
Private Sub Button_���^�v�Z_Click()
    CostForm.Show
End Sub

Private Sub CommandButton1_Click()

End Sub

'���[�U�[�t�H�[���̏����ʒu���Z������ǂݏo��
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�ݒ�")
    Me.Left = sh.Cells(�s.Left, ��.�����ʒu)
    Me.Top = sh.Cells(�s.Top, ��.�����ʒu)
End Sub
'�t�H�[��������ꏊ�����[�N�V�[�g�ɕۑ�
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�ݒ�")
    sh.Cells(�s.Left, ��.�����ʒu) = Me.Left
    sh.Cells(�s.Top, ��.�����ʒu) = Me.Top
End Sub


