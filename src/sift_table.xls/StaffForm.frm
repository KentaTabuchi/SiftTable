VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StaffForm 
   Caption         =   "��]�V�t�g"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11715
   OleObjectBlob   =   "StaffForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "StaffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�]�ƈ������ꗗ�\������t�H�[��
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Dim �X�^�b�t As Staff
    Dim record(5) As String
    Dim i As Integer
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Debug.Print i
        With Me.ListBox_StaffInfo
            .AddItem ("")
            .List(i, 0) = �X�^�b�t.���O
            .List(i, 1) = �X�^�b�t.��]�o�Ή�
            .List(i, 2) = �X�^�b�t.�o�Εs�j��
            .List(i, 3) = �X�^�b�t.���l
        End With
    i = i + 1
    Next
End Sub
