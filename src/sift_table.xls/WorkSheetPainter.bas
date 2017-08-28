Attribute VB_Name = "WorkSheetPainter"
Option Explicit
'////////////////////////////////////////////////////////
'���[�N�V�[�g�ɐF��h�郁�\�b�h��Z�߂郂�W���[��
'//////////////////////////////////////////////////////

Private Enum �s
    ���t = 4
    �\��[ = 3
    �\���[ = 41
End Enum
Private Enum ��
    ���O�� = 2
    �J�n�� = 3
    �ŏI�� = 39
End Enum


'�����Ŏ󂯎�����X�^�b�t�̏o�Εs���ɐF��h��
Public Sub paintImpossibleDay(�X�^�b�t As Staff)
    Application.ScreenUpdating = False
    Dim currentDay As String '�����͈͂̓��t���������
    Dim i As Integer
    
    If �X�^�b�t.���O = "" Then
        '���̏ꍇ�������Ȃ�
    Else
        For i = ��.�J�n�� To ��.�ŏI��
            currentDay = DateValue(Cells(�s.���t, i).Value)
            Dim �o�Εs�� As Variant
                For Each �o�Εs�� In �X�^�b�t.�o�Εs�����X�g
                    If currentDay = �o�Εs�� Then
                        Range(Cells(�X�^�b�t.row, i), Cells(�X�^�b�t.row + 1, i)).Select
                        Selection.Interior.Color = vbBlack
                    End If
                Next
        Next i
    End If
    Application.ScreenUpdating = True
End Sub
'�\�S�̂̓y���̗�ɐF��h��
Public Sub SetWeekendColor()
    Dim column As Integer
    For column = ��.�J�n�� To ��.�ŏI��
        Range(Cells(�s.�\��[, column), Cells(�s.�\���[, column)) _
        .Interior.ColorIndex = JudgeColor(column)
    Next column

End Sub
'�����Ŏw�肳�ꂽ�X�^�b�t�̍s�̓y���ɐF��h��
'�y���łȂ������ꍇ�͍X�ɂP�񂨂��Ƀx�[�W���ɓh��悤��IF���򂳂���
Public Sub SetWeekendColorUnit(�X�^�b�t As Staff)
    Dim column As Integer
    Dim weekdayIndex As Integer
    Dim rgbIndex As Long
    
    For column = ��.�J�n�� To ��.�ŏI��
        weekdayIndex = Weekday(ActiveSheet.Cells(�s.���t, column).Value, 1)
        
            Select Case weekdayIndex
            Case vbSunday
                rgbIndex = RGB(255, 153, 204) '�����ԐF
            Case vbSaturday
                rgbIndex = RGB(204, 255, 255) '�����F
            Case Else
                If (�X�^�b�t.row - 10) Mod 4 Then
                    rgbIndex = RGB(255, 255, 255) '��
                Else
                    rgbIndex = RGB(255, 255, 153) '�x�[�W��
                End If
           End Select
                Range(Cells(�X�^�b�t.row, column), Cells(�X�^�b�t.row + 1, column)).Interior.Color = rgbIndex
    Next column
    Call paintNameColumnAlternate(�X�^�b�t)
End Sub
'���O������s�����݂ɐF��������
Private Sub paintNameColumnAlternate(�X�^�b�t As Staff)
    Dim rgbIndex As Long
    
    If (�X�^�b�t.row - 10) Mod 4 Then
        rgbIndex = RGB(255, 255, 255) '��
    Else
        rgbIndex = RGB(255, 255, 153) '�x�[�W��
    End If
    
    Range(Cells(�X�^�b�t.row, ��.���O��), Cells(�X�^�b�t.row + 1, ��.���O��)).Interior.Color = rgbIndex
End Sub
'@unused
'�����Ŏw�肳�ꂽ��񂾂��y���F��h��
Public Sub SetWeekendColorVertical(column As Integer)
         Range(Cells(�s.�\��[, column), Cells(�s.�\���[, column)) _
        .Interior.ColorIndex = JudgeColor(column)
End Sub
'@unused
'�����̗�̗j���𔻒肵�ēh��F��Ԃ�
Private Function JudgeColor(column As Integer) As Integer
    Dim weekdayIndex As Integer
    weekdayIndex = Weekday(ActiveSheet.Cells(�s.���t, column).Value, 1)
        Select Case weekdayIndex
        Case vbSunday
            JudgeColor = 38
        Case vbSaturday
            JudgeColor = 20
        Case Else
            JudgeColor = 0
        End Select
End Function

'�����Ŏ󂯎�����X�^�b�t�̃V�t�g���̓������^�ɊY�����Ȃ������Â� RGB(200,200,200)=�O���[�@�ɓh��Ԃ�
'�Ј��̂Ƃ��A�Q��ڂ̂P0����肠�ƂȂ�h��
'�o�C�g�̂Ƃ��A1��ڂ�15�����O�Ȃ�h��@�Q��ڂ�15����肠�ƂȂ�h��
Public Sub ToDarkenOutOfTheCurrentMonth(�X�^�b�t As Staff)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim �ΏۃZ�� As Range
    Dim �� As Integer
    Dim �J�E���g�t���O As Integer

    If �X�^�b�t.�E�� = True Then
        For i = ��.�J�n�� To ��.�ŏI��
        �� = DAY(Cells(�s.���t, i))
        If �� = 11 Then
            �J�E���g�t���O = �J�E���g�t���O + 1
        End If
        If �J�E���g�t���O = 2 Then
            Set �ΏۃZ�� = Range(Cells(�X�^�b�t.row, i), Cells(�X�^�b�t.row + 1, i))
            �ΏۃZ��.Interior.Color = RGB(150, 150, 150)
       End If
       Next i
    End If

    If �X�^�b�t.�E�� = False Then
        For i = ��.�J�n�� To ��.�ŏI��
        �� = DAY(Cells(�s.���t, i))
        If �� = 16 Then
            �J�E���g�t���O = �J�E���g�t���O + 1
        End If
        If �J�E���g�t���O = 0 Or �J�E���g�t���O = 2 Then
            Set �ΏۃZ�� = Range(Cells(�X�^�b�t.row, i), Cells(�X�^�b�t.row + 1, i))
            �ΏۃZ��.Interior.Color = RGB(150, 150, 150)
       End If
       Next i
    End If
     
    Application.ScreenUpdating = True
End Sub
