VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostForm 
   Caption         =   "���^�v�Z"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   OleObjectBlob   =   "CostForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�T�Z�̋����v�Z�̈ꗗ��\������t�H�[��
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Me.ListBox_StaffCost.ColumnWidths = "50;50;50;10;"
    TableManager.setTablePosition
    
    Dim �X�^�b�t As Staff
    Dim record(4) As String
    Dim i As Integer
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Debug.Print i
        With Me.ListBox_StaffCost
            .AddItem ("")
            .List(i, 0) = �X�^�b�t.���O
            .List(i, 1) = �X�^�b�t.����
            .List(i, 2) = �X�^�b�t.���ԘJ������
            .List(i, 3) = �X�^�b�t.����
        End With
    i = i + 1
    Next
End Sub
