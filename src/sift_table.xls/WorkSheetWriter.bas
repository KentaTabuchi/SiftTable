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
'�Ј������̗񂪃V�[�g�̉���ڂɂ��邩���������ĕԂ�
Private Property Get �Ј�������() As Integer

    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    For i = ��.�J�n�� To ��.�ŏI��
        targetDay = Cells(�s.���t�s, i)
        If Day(targetDay) = 10 Then
            �Ј������� = i
        End If
    Next i
    Application.ScreenUpdating = True
   
End Property
'�o�C�g�����̗񂪃V�[�g�̉���ڂɂ��邩���������ĕԂ�
Private Property Get �o�C�g������() As Integer

    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    For i = ��.�J�n�� To ��.�ŏI��
        targetDay = Cells(�s.���t�s, i)
        
        If Day(targetDay) = 15 Then
            �o�C�g������ = i
        End If
    
    Next i
    Application.ScreenUpdating = True

End Property
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
        
        Select Case �X�^�b�t.�E��
            Case True '�Ј��̏ꍇ
                For j = ��.�J�n�� To �Ј�������
                Cells(�X�^�b�t.row, j).Value = ""
                Cells(�X�^�b�t.row + 1, j).Value = ""
                Next j
            Case False '�o�C�g�̏ꍇ
                For j = ��.���� To �o�C�g������
                Cells(�X�^�b�t.row, j).Value = ""
                Cells(�X�^�b�t.row + 1, j).Value = ""
                Next j
        End Select
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
'�����ɓn���ꂽ�X�^�b�t�̊�{�V�t�g�����Ԃ��w�肵�ē]�L
Public Sub WriteBasicShiftByTurn(�X�^�b�t As Staff, �J�n�� As Date, �ŏI�� As Date)

    If �X�^�b�t.���O = "" Then  '���O�񂪋󗓂̏ꍇ�͂��̍s�̃V�t�g������ɂ���
        Dim j As Integer
        
        Select Case �X�^�b�t.�E��
            Case True '�Ј��̏ꍇ
                For j = ��.�J�n�� To �Ј�������
                Cells(�X�^�b�t.row, j).Value = ""
                Cells(�X�^�b�t.row + 1, j).Value = ""
                Next j
            Case False '�o�C�g�̏ꍇ
                For j = ��.���� To �o�C�g������
                Cells(�X�^�b�t.row, j).Value = ""
                Cells(�X�^�b�t.row + 1, j).Value = ""
                Next j
        End Select
    ElseIf �X�^�b�t.���O = "�s��" Then
        '���̏����ł͉������Ȃ��̂ł��̍s�͋󗓂ō����Ă���
    Else
    
    Dim currentDay As String '�����͈͂̓��t���������
    Dim �J�n���� As Integer
    Dim �ŏI���� As Integer
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
    
    '�J�n���ƍŏI�����\�̉���ڂɂ���̂���������
    Dim k
    For k = ��.���� To ��.�ŏI��
        currentDay = DateValue(Cells(�s.���t�s, k).Value)
        If �J�n�� = CDate(currentDay) Then
            �J�n���� = k
        End If
        If �ŏI�� = CDate(currentDay) Then
            �ŏI���� = k
        End If
    Next k
    
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = �J�n���� To �ŏI����
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
'�����ɓn���ꂽ�X�^�b�t�̑O���Ƃ��Ԃ镔����O���̃V�[�g����R�s�[����
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
'[�J����]�����ɓn���ꂽ�X�^�b�t�̎����Ƃ��Ԃ镔���������̃V�[�g����R�s�[����
'�����̗���v�Z�ŎZ�o����K�v������B�v���p�e�B�̐V�݂��K�v
Public Sub CopyFromNextMonth(�X�^�b�t As Staff)
    If �X�^�b�t.���O = "" Then
    Else
    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    Dim ���� As Integer
        If �X�^�b�t.�E�� = True Then
            ���� = �Ј�������
        ElseIf �X�^�b�t.�E�� = False Then
            ���� = �o�C�g������
        End If
    For i = ���� + 1 To ��.�ŏI��
        targetDay = Cells(�s.���t�s, i)
        Dim card As TimeCard
        For Each card In �X�^�b�t.�����V�t�g
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


