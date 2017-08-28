Attribute VB_Name = "WorkSheetRuler"
Option Explicit

'////////////////////////////////////////////////////////
'���[�N�V�[�g�Ɍr�����������\�b�h��Z�߂����W���[���i�J����)
'//////////////////////////////////////////////////////

Private Enum �s
    �\��[ = 3
    ���t = 4
    �j���s = 5
    �j���s = 6
    ��c�s = 8
    ��l�ڍs = 10
    �\���[ = 41
End Enum
Private Enum ��
    ���O�� = 2
    �J�n�� = 3
    �ŏI�� = 39
    �J�����ԗ� = 40
    �T�x�\��� = 42
End Enum
'�e�[�u���̌r���������������\�b�h
'�@�܂��]�ƈ��̃V�t�g���Ɍr���������@�O�g���׎����@���������j��
'�A���O���@�O�g����
'�B���ɏj�����Ɖ�c���ɐ��������@�O�g�����������@�O�g�c��=�׎����@���������Ȃ�
'�C�j���s�Ɖ�c�s�̃^�C�g����Ɍr���������@���E�������@�㉺�@�׎���
'���t���i���E���E�j���s�j�@�O�g�����������@�O�g�c�����׎����@���������Ȃ�
'���t���Ɍr��������'�D
'�J�����ԗ�Ɍr���������E
'�Ō�ɕ\�̊O�g����d�� '�F
Public Sub ruleLine()
    Call clearAllLine
    TableManager.initialize
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetRuler.ruleLineToStaffPane(�X�^�b�t) '�@
        Call WorkSheetRuler.ruleLineToNamePane(�X�^�b�t) '�A
        Call WorkSheetRuler.ruleLineToWorkTimePane(�X�^�b�t) '�E
    Next
    Call WorkSheetRuler.ruleLineToSchedulePane(TableManager.��c��) '�B
    Call WorkSheetRuler.ruleLineToSchedulePane(TableManager.�j��)
    Call WorkSheetRuler.ruleLineToScheduleNamePane(TableManager.��c��) '�C
    Call WorkSheetRuler.ruleLineToScheduleNamePane(TableManager.�j��)
    Call ruleLineToDatePane '�D
    Call WorkSheetRuler.ruleLineToAroundTable '�F
End Sub
'�r����S���������\�b�h
'�P�}�X�]���ɏ����Ȃ��Ǝ���VBA�Ōr�����������Ƃ���ƕ\�̒[�̂�����œ�̃G���[���N����̂ő��߂ɏ����Ă���B
'�S�~�f�[�^�ɂ��G�N�Z���̃o�O�̂悤��
'�ǂ����S�~�f�[�^�łȂ��Z�������ɖ�肪����悤�����ꉞ���̂܂܂ɂ��Ă���
Private Sub clearAllLine()
    Dim �ΏۃZ�� As Range
        Set �ΏۃZ�� = Range(Cells(�s.�\��[ - 1, ��.���O�� - 1), Cells(�s.�\���[ + 2, ��.�J�����ԗ� + 2))
        With �ΏۃZ��
            .Borders.LineStyle = xlLineStyleNone
        End With
    Application.ScreenUpdating = True
End Sub
'�J�����ԗ񂩂�T�x��܂Ōr��������
Private Sub ruleLineToWorkTimePane(�X�^�b�t As Staff)
    Application.ScreenUpdating = False
    Dim �ΏۃZ�� As Range
        Set �ΏۃZ�� = Range(Cells(�X�^�b�t.row, ��.�J�����ԗ�), Cells(�X�^�b�t.row + 1, ��.�T�x�\���))
        With �ΏۃZ��
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlDash
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
    Application.ScreenUpdating = True
End Sub
'���t���Ɍr���������@���������j���@�c�����׎���
Private Sub ruleLineToDatePane()
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim i As Integer
    Dim �ΏۃZ�� As Range
    For i = ��.���O�� To ��.�ŏI��
        Set �ΏۃZ�� = Range(Cells(�s.�\��[, i), Cells(�s.�j���s, i))
        With �ΏۃZ��
            .Borders(xlInsideHorizontal).LineStyle = xlDash '�^�񒆂̉����ɔj��������
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next i
    Set �ΏۃZ�� = Range(Cells(�s.�\��[, ��.�J�n��), Cells(�s.�\���[, ��.�J�n��))
    �ΏۃZ��.Borders(xlEdgeLeft).Weight = xlMedium
    Set �ΏۃZ�� = Range(Cells(�s.�\��[, ��.�J�����ԗ�), Cells(�s.�\���[, ��.�J�����ԗ�))
    �ΏۃZ��.Borders(xlEdgeLeft).Weight = xlMedium
    Application.ScreenUpdating = True
End Sub
'�\�̂����΂�O���̘g�ɓ�d��������
'�\�̉��[�͖��O�񂪌�������Ă��邽�߁A���̂܂܎w�肷���1004���s���G���[�iNull Pointer Exception)�ɂȂ��Ă��܂�
'���̂��ߏ����𕪂��Ĉ���̍s�����ӂɈ���
Private Sub ruleLineToAroundTable()
    Dim �ΏۃZ�� As Range
    
        Set �ΏۃZ�� = Range(Cells(�s.�\��[, ��.���O��), Cells(�s.�\���[, ��.�T�x�\���))
        With �ΏۃZ��
            .Borders(xlEdgeTop).LineStyle = xlDouble
            .Borders(xlEdgeLeft).LineStyle = xlDouble
        End With
        
        Set �ΏۃZ�� = Range(Cells(�s.�\���[ + 1, ��.���O��), Cells(�s.�\���[ + 1, ��.�T�x�\���))
        With �ΏۃZ��
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With
        
        Set �ΏۃZ�� = Range(Cells(�s.�\��[, ��.�T�x�\���), Cells(�s.�\���[, ��.�T�x�\���))
        With �ΏۃZ��
            .Borders(xlEdgeRight).LineStyle = xlDouble
        End With
    
    Application.ScreenUpdating = True
End Sub
'�����Ŏ󂯎�����X�^�b�t�̃V�t�g���Ɍr��������
Private Sub ruleLineToStaffPane(�X�^�b�t As Staff)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim �ΏۃZ�� As Range
    Dim �� As Integer
    For i = ��.�J�n�� To ��.�ŏI��
        �� = DAY(Cells(�s.���t, i))
        Set �ΏۃZ�� = Range(Cells(�X�^�b�t.row, i), Cells(�X�^�b�t.row + 1, i))
        �ΏۃZ��.Borders.LineStyle = xlContinuous '�㉺���E�ɍ׎���������
        �ΏۃZ��.Borders(xlInsideHorizontal).LineStyle = xlDash '�^�񒆂̉����ɔj��������
        
        If �� = 16 Then
            If �X�^�b�t.�E�� = False Then
                �ΏۃZ��.Borders(xlEdgeLeft).Weight = xlThick
                �ΏۃZ��.Borders(xlEdgeLeft).Color = RGB(255, 0, 0)
            End If
        ElseIf �� = 11 Then
            If �X�^�b�t.�E�� = True Then
                 �ΏۃZ��.Borders(xlEdgeLeft).Weight = xlThick
                 �ΏۃZ��.Borders(xlEdgeLeft).Color = RGB(255, 0, 0)
            End If
        End If
        
    Next i
    Application.ScreenUpdating = True
End Sub
Private Sub ruleLineToSchedulePane(�X�P�W���[�� As Schedule)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim �ΏۃZ�� As Range
    For i = ��.�J�n�� To ��.�ŏI��
        Set �ΏۃZ�� = Range(Cells(�X�P�W���[��.��ƍs, i), Cells(�X�P�W���[��.��ƍs + 1, i))
        With �ΏۃZ��
            .Borders.LineStyle = xlContinuous '�㉺���E�ɍ׎���������
            .Borders(xlInsideHorizontal).LineStyle = xlDash '�^�񒆂̉����ɔj��������
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next i
    Application.ScreenUpdating = True
End Sub
'�����Ŏ󂯎�����X�^�b�t�̖��O���Ɍr��������
Private Sub ruleLineToNamePane(�X�^�b�t As Staff)
    Application.ScreenUpdating = False
    Dim �ΏۃZ�� As Range
        Set �ΏۃZ�� = Range(Cells(�X�^�b�t.row, ��.���O��), Cells(�X�^�b�t.row + 1, ��.���O��))
        With �ΏۃZ��
            .Borders.LineStyle = xlContinuous '�㉺���E�ɍ׎���������
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
    Application.ScreenUpdating = True
End Sub
'�X�P�W���[���s�̃^�C�g����ɑ���������
Private Sub ruleLineToScheduleNamePane(�X�P�W���[�� As Schedule)
    Application.ScreenUpdating = False
    Dim �ΏۃZ�� As Range
        Set �ΏۃZ�� = Range(Cells(�X�P�W���[��.��ƍs, ��.���O��), Cells(�X�P�W���[��.��ƍs + 1, ��.���O��))
        With �ΏۃZ��
            .Borders.LineStyle = xlContinuous '�㉺���E�ɍ׎���������
            .Borders.Weight = xlMedium '�㉺���E�𑾐��ɂ���
        End With
    Application.ScreenUpdating = True
End Sub
