VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimeCardButtonForm 
   Caption         =   "�Αӓ���"
   ClientHeight    =   3165
   ClientLeft      =   13050
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "TimeCardButtonForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "TimeCardButtonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum ��
    �����ʒu = 5
End Enum
Private Enum �s
    Left = 5
    Top = 6
End Enum
Private Sub UserForm_Click()
Debug.Print "left="; Me.Left, "top="; Me.Top
End Sub
'���[�U�[�t�H�[���̏����ʒu���Z������ǂݏo��
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�ݒ�")
    Me.Left = sh.Cells(�s.Left, ��.�����ʒu)
    Me.Top = sh.Cells(�s.Top, ��.�����ʒu)
End Sub
'�t�H�[��������ꏊ�����[�N�V�[�g�ɕۑ�
'�f�X�g���N�^����Me�̈ʒu��񂪏����Ă��܂�����Ŕ�������̂ł����ɋL�q����B
'�f�X�g���N�^�̒��O�ɔ�������C�x���g
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�ݒ�")
    sh.Cells(�s.Left, ��.�����ʒu) = Me.Left
    sh.Cells(�s.Top, ��.�����ʒu) = Me.Top
End Sub

