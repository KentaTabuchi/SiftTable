Attribute VB_Name = "TableManager"
Option Explicit

Enum �s  '�V�[�g�̍s�ԍ��ɖ��O������
    �j���s = 6
    �\��s = 8
    �X�^�b�t1 = 10
    �X�^�b�t2 = 12
    �X�^�b�t3 = 14
    �X�^�b�t4 = 16
    �X�^�b�t5 = 18
    �X�^�b�t6 = 20
    �X�^�b�t7 = 22
    �X�^�b�t8 = 24
    �X�^�b�t9 = 26
    �X�^�b�t10 = 28
    �X�^�b�t11 = 30
    �X�^�b�t12 = 32
    �X�^�b�t13 = 34
    �X�^�b�t14 = 36
    �X�^�b�t15 = 38
    �X�^�b�t16 = 40
End Enum

Public �X�^�b�t���X�g As Collection
Public �j�� As Schedule
Public ��c�� As Schedule

Public Sub setTablePosition()
    Set �X�^�b�t���X�g = New Collection
    Dim i As Integer
    For i = 0 To 15
        Dim �X�^�b�t As Staff
        Set �X�^�b�t = New Staff
        �X�^�b�t���X�g.Add �X�^�b�t
    Next
    �X�^�b�t���X�g.Item(1).row = �s.�X�^�b�t1
    �X�^�b�t���X�g.Item(2).row = �s.�X�^�b�t2
    �X�^�b�t���X�g.Item(3).row = �s.�X�^�b�t3
    �X�^�b�t���X�g.Item(4).row = �s.�X�^�b�t4
    �X�^�b�t���X�g.Item(5).row = �s.�X�^�b�t5
    �X�^�b�t���X�g.Item(6).row = �s.�X�^�b�t6
    �X�^�b�t���X�g.Item(7).row = �s.�X�^�b�t7
    �X�^�b�t���X�g.Item(8).row = �s.�X�^�b�t8
    �X�^�b�t���X�g.Item(9).row = �s.�X�^�b�t9
    �X�^�b�t���X�g.Item(10).row = �s.�X�^�b�t10
    �X�^�b�t���X�g.Item(11).row = �s.�X�^�b�t11
    �X�^�b�t���X�g.Item(12).row = �s.�X�^�b�t12
    �X�^�b�t���X�g.Item(13).row = �s.�X�^�b�t13
    �X�^�b�t���X�g.Item(14).row = �s.�X�^�b�t14
    �X�^�b�t���X�g.Item(15).row = �s.�X�^�b�t15
    �X�^�b�t���X�g.Item(16).row = �s.�X�^�b�t16
    Set �j�� = New Schedule
    �j��.��ƍs = �s.�j���s
    Set ��c�� = New Schedule
    ��c��.��ƍs = �s.�\��s

End Sub




