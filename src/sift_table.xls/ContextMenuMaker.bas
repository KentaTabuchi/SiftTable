Attribute VB_Name = "ContextMenuMaker"
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�E�N���b�N���j���[��ǉ����郁�\�b�h�Q
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'[�J����] �E�N���b�N���j���[��ǉ��B�������Č����炢�̂Őe���j���[�͎���ł��邱�Ƃ����������A
'���ۂ̃��j���[�͎q���j���[�Ɋi�[����
Private Menu_�e���j���[ As CommandBarPopup
Public Sub makeContextMenu()
    Application.CommandBars("cell").Reset
    Set Menu_�e���j���[ = Application.CommandBars("cell").Controls.Add(Type:=msoControlPopup)
    Menu_�e���j���[.Caption = "�V�t�g�\"
    Dim menuItems As Collection
    Call setSubMenu("��{�V�t�g", "��{�V�t�g_�̈�w��", "��]�V�t�g", "���^�v�Z", "�J������", "���x�T�x", "�F�Z�b�g", "�o�Εs��", "�r���C��", "�\��Z�b�g", "���݂ɔw�i�h��")
End Sub
'�T�u���j���[�̍쐬
Private Sub setSubMenu(ParamArray menuNames())
    Dim menuName As Variant
    For Each menuName In menuNames
        Dim subMenu As CommandBarButton
        Set subMenu = Menu_�e���j���[.Controls.Add(Type:=msoControlButton, temporary:=True)
        With subMenu
            .Caption = menuName
            .TooltipText = menuName
            .onAction = menuName
        End With
    Next
End Sub
'�f�[�^�x�[�X����e�X�^�b�t�̊�{�V�t�g���z���グ�G�N�Z���V�[�g�֓]�L����
Private Sub ��{�V�t�g()
    Dim �X�^�b�t As Staff
    Dim �i���� As String
    Dim �X�^�b�t���� As Integer
    Dim count As Integer: count = 0
    
    TableManager.initialize
    �X�^�b�t���� = TableManager.�X�^�b�t���X�g.count
    ProgressForm.Show vbModeless
    �i���� = count & "/" & �X�^�b�t���� & "�l�@����"
    ProgressForm.ProgressLabel.Caption = �i����
    
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        
        If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "�����𒆒f���܂����B"
            End
        End If
        
        count = count + 1
            If �X�^�b�t.���O = "" Then
                �i���� = "�X�^�b�t�s�݁A�X�L�b�v���܂��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
            ElseIf �X�^�b�t.���O = "�s��" Then
                �i���� = "�s�����͕ύX���܂���A�X�L�b�v���܂��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
            Else
                �i���� = �X�^�b�t.���O & "�̃V�t�g���쐬���E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
            End If
        ProgressForm.ProgressLabel.Caption = �i����
        DoEvents 'Wait�����H����������Ȃ��ƊG�揈�����ǂ��t�������������Ȃ��܂܃v���O���X�\�����I����Ă��܂�
        Call WorkSheetWriter.WriteBasicShift(�X�^�b�t)
        Call WorkSheetWriter.CopyFromPreviousMonth(�X�^�b�t)
        Call WorkSheetWriter.CopyFromNextMonth(�X�^�b�t)
    Next
    Unload ProgressForm
End Sub
Public Sub ��{�V�t�g_�̈�w��()
    SelectForm.Show vbModeless
End Sub
Private Sub ��]�V�t�g()
    StaffForm.Show
End Sub
Private Sub ���^�v�Z()
    CostForm.Show
End Sub
'�J�����Ԃ��V�[�g�ɋL�����郁�\�b�h�B
'�����ł̓v���O���X�o�[�̏��������Ă��邾���B
'���W�b�N�̖{�̂́@staff�N���X�̌��ԘJ�����ԃv���p�e�B�Ōv�Z���Ă���B
Private Sub �J������()

    Dim �X�^�b�t As Staff
    Dim �i���� As String
    Dim �X�^�b�t���� As Integer
    Dim count As Integer: count = 0
    TableManager.initialize
    �X�^�b�t���� = TableManager.�X�^�b�t���X�g.count
    ProgressForm.Show vbModeless
    �i���� = count & "/" & �X�^�b�t���� & "�l�@����"
    ProgressForm.ProgressLabel.Caption = �i����
    
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g

        If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "�����𒆒f���܂����B"
            End
        End If
        
        count = count + 1
            If �X�^�b�t.���O = "" Then
                �i���� = "�X�^�b�t�s�݁A�X�L�b�v���܂��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
            Else
                �i���� = �X�^�b�t.���O & "�̌��ԘJ�����Ԃ��v�Z���E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
            End If
        ProgressForm.ProgressLabel.Caption = �i����
        DoEvents 'Wait�����B����������Ȃ��ƊG�揈�����ǂ��t�������������Ȃ��܂܃v���O���X�\�����I����Ă��܂�
        Call WorkSheetWriter.WriteWorkTime(�X�^�b�t)
    Next
    Unload ProgressForm
    
End Sub
Private Sub ���x�T�x()
    Dim �X�^�b�t As Staff
    TableManager.initialize
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetWriter.WriteNumOfPublicHoliday(�X�^�b�t)
    Next
End Sub
'�y���ɐF��h��
Private Sub �F�Z�b�g()

    Dim �X�^�b�t As Staff
    TableManager.initialize
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
    Call WorkSheetPainter.SetWeekendColorUnit(�X�^�b�t)
    Call WorkSheetPainter.ToDarkenOutOfTheCurrentMonth(�X�^�b�t)
    Next
    
End Sub
'�S���̏o�Εs���ɐF��h��
Private Sub �o�Εs��()
    Dim �i���� As String
    Dim �X�^�b�t���� As Integer
    Dim count As Integer: count = 0
    Dim �X�^�b�t As Staff
    
    TableManager.initialize
    �X�^�b�t���� = TableManager.�X�^�b�t���X�g.count
    ProgressForm.Show vbModeless
    �i���� = count & "/" & �X�^�b�t���� & "�l�@����"
    ProgressForm.ProgressLabel.Caption = �i����
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        
            If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "�����𒆒f���܂����B"
            End
        End If
        
        count = count + 1
        If �X�^�b�t.���O = "" Then
            �i���� = "�X�^�b�t�s�݁A�X�L�b�v���܂��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
        ElseIf �X�^�b�t.���O = "�s��" Then
            �i���� = "�s���l�����͐G��܂���B�X�L�b�v���܂��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
        Else
            �i���� = �X�^�b�t.���O & "�̏o�Εs����ǂݍ��ݒ��E�E�E" & vbNewLine & count & "/" & �X�^�b�t���� & "�l�@����"
        End If
        ProgressForm.ProgressLabel.Caption = �i����
        DoEvents
        
        Call WorkSheetPainter.paintImpossibleDay(�X�^�b�t)
    Next
    Unload ProgressForm
End Sub
Private Sub �r���C��()
    Call WorkSheetRuler.ruleLine
End Sub
Private Sub �\��Z�b�g()
    Call WorkSheetWriter.WriteLegalHoliday(TableManager.�j��)
    Call WorkSheetWriter.WriteMeetingDay(TableManager.��c��)
End Sub
Private Sub ���݂ɔw�i�h��()
    TableManager.initialize
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        Call WorkSheetPainter.paintBackColorInTurn(�X�^�b�t)
    Next
End Sub
