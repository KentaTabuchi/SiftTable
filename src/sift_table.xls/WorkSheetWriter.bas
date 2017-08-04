Attribute VB_Name = "WorkSheetWriter"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�f�[�^�����[�N�V�[�g�ɓ]�L���郁�\�b�h��Z�߂����W���[��
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Enum ��
    �J�n�� = 3
    �O������ = 7
    ���� = 8   '�P�U���̎�
    �ŏI�� = 39
    �J�����ԗ� = 40
    ���x�� = 41
    �T�x�� = 42
End Enum
Private Enum �s
    ���t�s = 4
End Enum

'�����ɓn���ꂽ�o�ދΎ��Ԃ��A�N�e�B�u�Z���ɋL������
Public Sub WriteTimeCard(card As TimeCard)
    If card.�o�Ύ��� = "�N���A" Then
        Selection = ""
        ActiveCell.Offset(1, 0).Activate
        Selection = ""
        ActiveCell.Offset(-1, 1).Activate
    Else
        Selection = card.�o�Ύ���
        ActiveCell.Offset(1, 0).Activate
        Selection = card.�ދΎ���
        ActiveCell.Offset(-1, 1).Activate
    End If
End Sub
'�����ɓn���ꂽ�X�^�b�t�̊�{�V�t�g��]�L
Public Sub WriteBasicShift(�X�^�b�t As Staff)

    If �X�^�b�t.���O = "" Then  '���O�񂪋󗓂̏ꍇ�͂��̍s�̃V�t�g������ɂ���
        Dim j As Integer
        For j = ��.���� To ��.�ŏI��
            Cells(�X�^�b�t.row, j).Value = ""
            Cells(�X�^�b�t.row + 1, j).Value = ""
        Next j
    ElseIf �X�^�b�t.���O = "�s��" Then
        '���̏����ł͉������Ȃ��̂ł��̍s�͋󗓂ō����Ă���
    Else
    Dim currentDay As String '�����͈͂̓��t���������
    Dim column As Integer
    Dim ���j�� As TimeCard
    Dim �Ηj�� As TimeCard
    Dim ���j�� As TimeCard
    Dim �ؗj�� As TimeCard
    Dim ���j�� As TimeCard
    Dim �y�j�� As TimeCard
    Dim ���j�� As TimeCard
    On Error Resume Next
    Set ���j�� = �X�^�b�t.��{�V�t�g.Item("���j��")
    Set �Ηj�� = �X�^�b�t.��{�V�t�g.Item("�Ηj��")
    Set ���j�� = �X�^�b�t.��{�V�t�g.Item("���j��")
    Set �ؗj�� = �X�^�b�t.��{�V�t�g.Item("�ؗj��")
    Set ���j�� = �X�^�b�t.��{�V�t�g.Item("���j��")
    Set �y�j�� = �X�^�b�t.��{�V�t�g.Item("�y�j��")
    Set ���j�� = �X�^�b�t.��{�V�t�g.Item("���j��")
    
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = ��.���� To ��.�ŏI��
            currentDay = DateValue(Cells(�s.���t�s, i).Value)
                Select Case Weekday(currentDay)
                Case vbSunday
                    Cells(�X�^�b�t.row, i).Value = ���j��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = ���j��.�ދΎ���
                Case vbMonday
                    Cells(�X�^�b�t.row, i).Value = ���j��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = ���j��.�ދΎ���
                Case vbTuesday
                    Cells(�X�^�b�t.row, i).Value = �Ηj��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = �Ηj��.�ދΎ���
                Case vbWednesday
                    Cells(�X�^�b�t.row, i).Value = ���j��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = ���j��.�ދΎ���
                Case vbThursday
                    Cells(�X�^�b�t.row, i).Value = �ؗj��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = �ؗj��.�ދΎ���
                Case vbFriday
                    Cells(�X�^�b�t.row, i).Value = ���j��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = ���j��.�ދΎ���
                Case vbSaturday
                    Cells(�X�^�b�t.row, i).Value = �y�j��.�o�Ύ���
                    Cells(�X�^�b�t.row + 1, i).Value = �y�j��.�ދΎ���
        End Select
        Next i
    Application.ScreenUpdating = True
    End If
End Sub
'�����ɓn���ꂽ�X�^�b�t�̑O���Ƃ��Ԃ�P�P������P�T���܂ł�O���̃V�[�g����R�s�[����
Public Sub CopyFromPreviousMonth(�X�^�b�t As Staff)
    If �X�^�b�t.���O = "" Then
    Else
    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    For i = ��.�J�n�� To ��.�O������
        targetDay = Cells(�s.���t�s, i)
        Dim card As TimeCard
        For Each card In �X�^�b�t.�O���V�t�g
        If card.���t = targetDay Then
            Cells(�X�^�b�t.row, i).Value = card.�o�Ύ���
            Cells(�X�^�b�t.row + 1, i).Value = card.�ދΎ���
        End If
        Next
    Next i
    Application.ScreenUpdating = True
    End If
End Sub
'���[�N�V�[�g�ɏj������������
Public Sub WriteLegalHoliday(�j�� As Schedule)
    Dim targetDay As Date '�����͈͂̓��t���������
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = ��.���� To ��.�ŏI��
            targetDay = Cells(�s.���t�s, i).Value
            Dim �C�x���g As Events
            For Each �C�x���g In �j��.�j�����X�g
                If �C�x���g.���t = targetDay Then
                    Cells(�j��.��ƍs, i).Value = �C�x���g.���e
                End If
            Next
        Next i
    Application.ScreenUpdating = True
End Sub
'���[�N�V�[�g�ɉ�c���̗\�����������
Public Sub WriteMeetingDay(��c���X�g As Schedule)
    Dim targetDay As Date '�����͈͂̓��t���������
    Application.ScreenUpdating = False
    Range(Cells(��c���X�g.��ƍs, ��.�J�n��), Cells(��c���X�g.��ƍs + 1, ��.�ŏI��)).Select
    Selection.Interior.Color = vbWhite
    Selection.ClearContents
        Dim i As Integer
        For i = ��.���� To ��.�ŏI��
            targetDay = Cells(�s.���t�s, i).Value
            Dim �C�x���g As Events
            For Each �C�x���g In ��c���X�g.��c��
                If �C�x���g.���t = targetDay Then
                    Cells(��c���X�g.��ƍs, i).Value = �C�x���g.���e
                    Range(Cells(��c���X�g.��ƍs, i), (Cells(��c���X�g.��ƍs + 1, i + �C�x���g.����))).Interior.Color = vbYellow
                  
                End If
            Next
        Next i
    Application.ScreenUpdating = True
End Sub
'���[�N�V�[�g�̘J�����ԗ��ɘJ�����Ԃ��������ރ��\�b�h
Public Sub WriteWorkTime(�X�^�b�t As Staff)
    Cells(�X�^�b�t.row, ��.�J�����ԗ�).Value = �X�^�b�t.���ԘJ������
End Sub
'���[�N�V�[�g�Ɍ��x�A�T�x�񐔂��������ރ��\�b�h
Public Sub WriteNumOfPublicHoliday(�X�^�b�t As Staff)

    Dim rules As CampanyRules
    Set rules = New CampanyRules
    Dim ���t As Date: ���t = Cells(2, 2)
    Dim ������x As Byte
    Dim ����T�x As Byte
    Select Case �X�^�b�t.�E��
    Case False
        ������x = 0
        ����T�x = 0
    Case True
    ������x = rules.GetGivenPublicHolidays(Date)
    ����T�x = rules.GetGivenWeeklyHolidays(Date)
    End Select
    
    Cells(�X�^�b�t.row, ��.���x��).Value = �X�^�b�t.���x��
    Cells(�X�^�b�t.row, ��.�T�x��).Value = �X�^�b�t.�T�x��
    Cells(�X�^�b�t.row + 1, ��.���x��).Value = ������x
    Cells(�X�^�b�t.row + 1, ��.�T�x��).Value = ����T�x
End Sub


