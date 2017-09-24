Attribute VB_Name = "NewSheetBuilder"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�V�����V�[�g���������郂�W���[��
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'���C�����\�b�h
'�V�K�V�[�g����������
Private Const LEFT_SPACE = 1
Private Const DAY_ROW = 2
Private Const NAME_COLUMN = 2
Private Const LEFT_EDGE_OF_SHIFT = 3
Private Const RGHIT_EDGE_OF_SHIFT = 39
Private Const WORK_TIME_COLUMN = 40
Private Const HOLIDAY_COUNT_COLUMN = 41

Public Sub setNewSheet()
Dim year_ As Integer
Dim month_ As Integer
   year_ = Application.InputBox(prompt:="���N�̃V�t�g�����܂����H���p�����œ��͂��Ă��������B�@�� 2017", Title:="�V�K�V�t�g�쐬", Type:=1)
   month_ = Application.InputBox(prompt:="�����̃V�t�g�����܂����H���p�����œ��͂��Ă��������B ��@6", Title:="�V�K�V�t�g�쐬", Type:=1)
    If (month_ > 12 Or month_ < 1) Then
        GoTo NumericValueOutOfRangeException
    End If
   Call main(year_, month_)
   Exit Sub
NumericValueOutOfRangeException:
    MsgBox "1����12�Ŏw�肵�Ă�������"
    Exit Sub
End Sub
'���̃N���X�̃��C�����\�b�h
Private Sub main(year_, month_ As Integer)
    
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = addNewSheet(month_)
    Call setStartDay(newWorkSheet, year_)
    Call setEndDay(newWorkSheet, year_)
    Call WorkSheetRuler.ruleLine
    Call margeNameColumnCells
    Call adjustColumnWidth
    Call setColumnTitleText
    Call setTextFont
    Call setDateText
    Call setTextFormat
    Call setAlignment
End Sub

'�����ɓn���ꂽ���̖��O�Ɂu���v�����ăV�[�g���ɂ��쐬����
Private Function addNewSheet(monthNo As Integer) As Worksheet
    On Error GoTo SheetNameException
        Dim sheetName As String
    sheetName = Trim(str(monthNo) & "��")
    
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = Worksheets.Add()
    newWorkSheet.name = sheetName
    Set addNewSheet = newWorkSheet
    Exit Function
SheetNameException:
    MsgBox "���ɂ���V�[�g�Ɠ��������͍��܂���B�P�N�O�̃V�[�g�͍폜���Ă��������B"
    End
End Function

Private Sub setStartDay(newWorkSheet As Worksheet, year_)
    Dim text As String
    Dim month_ As Integer
    month_ = Val(newWorkSheet.name)
    text = year_ & "/" & (month_ - 1) & "/" & 16
    newWorkSheet.Cells(DAY_ROW, NAME_COLUMN).Value = text
End Sub
Private Sub setEndDay(newWorkSheet As Worksheet, year_)
    Dim text As String
    Dim month_ As Integer
    month_ = Val(newWorkSheet.name)
    text = year_ & "/" & (month_) & "/" & 15
    newWorkSheet.Cells(DAY_ROW, NAME_COLUMN + 2).Value = text
End Sub
'���O��̃Z�����c�ɂQ�R�}�ÂŌ�������
Private Sub margeNameColumnCells()
    Application.DisplayAlerts = False
    
    Dim i As Integer
    Dim rg As Range
    For i = 6 To 40 Step 2
        Set rg = Range(Cells(i, NAME_COLUMN), Cells(i + 1, NAME_COLUMN))
        rg.MergeCells = True
        rg.HorizontalAlignment = xlCenter
    Next i
        Set rg = Range(Cells(2, 4), Cells(2, 6))
        rg.MergeCells = True
    
    Application.DisplayAlerts = True
    
End Sub
'�񕝂̒���
Private Sub adjustColumnWidth()
    Dim rg As Range
    Set rg = Range(Cells(1, LEFT_SPACE), Cells(1, LEFT_SPACE))
    rg.ColumnWidth = 1
    Set rg = Range(Cells(1, LEFT_EDGE_OF_SHIFT), Cells(1, RGHIT_EDGE_OF_SHIFT))
    rg.ColumnWidth = 3.38
    Set rg = Range(Cells(1, NAME_COLUMN), Cells(1, NAME_COLUMN))
    rg.ColumnWidth = 8
    Set rg = Range(Cells(1, WORK_TIME_COLUMN), Cells(1, WORK_TIME_COLUMN))
    rg.ColumnWidth = 6.88
    Set rg = Range(Cells(1, HOLIDAY_COUNT_COLUMN), Cells(1, HOLIDAY_COUNT_COLUMN + 1))
    rg.ColumnWidth = 3.13
End Sub
'�e�L�X�g���V�[�g�ɒ��ڑ��
Private Sub setColumnTitleText()
    Cells(4, NAME_COLUMN) = "���t"
    Cells(5, NAME_COLUMN) = "�j��"
    Cells(8, WORK_TIME_COLUMN) = "�J������"
    Cells(7, HOLIDAY_COUNT_COLUMN) = "�T�x"
    Cells(8, HOLIDAY_COUNT_COLUMN) = "�擾"
    Cells(9, HOLIDAY_COUNT_COLUMN) = "����"
    Cells(7, HOLIDAY_COUNT_COLUMN + 1) = "���x"
    Cells(8, HOLIDAY_COUNT_COLUMN + 1) = "�擾"
    Cells(9, HOLIDAY_COUNT_COLUMN + 1) = "����"
    Cells(1, 3) = "���x"
End Sub
   '�����A�j���̍s�ɓ��t���������ޏ����B
Private Sub setDateText()
    '���߂̗񂾂������Ⴄ�̂ŕʏ���
    Range(Cells(1, 2), Cells(1, 2)).FormulaR1C1 = "=MONTH(R[1]C+30)"
    Range(Cells(3, LEFT_EDGE_OF_SHIFT), Cells(3, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=MONTH(R[1]C)"
    Range(Cells(4, LEFT_EDGE_OF_SHIFT), Cells(4, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=R[-2]C[-1]"
    Range(Cells(5, LEFT_EDGE_OF_SHIFT), Cells(5, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=TEXT(R[-1]C,""aaa"")"
    
    '�Q��ڂ���Ō�܂ŌJ��Ԃ�
    Dim i As Integer
    Dim rg As Range
    For i = LEFT_EDGE_OF_SHIFT + 1 To RGHIT_EDGE_OF_SHIFT
        Set rg = Range(Cells(3, i), Cells(3, i))
        rg.FormulaR1C1 = "=MONTH(R[1]C)"
        Set rg = Range(Cells(4, i), Cells(4, i))
        rg.FormulaR1C1 = "=RC[-1]+1"
        Set rg = Range(Cells(5, i), Cells(5, i))
        rg.FormulaR1C1 = "=TEXT(R[-1]C,""aaa"")"

    Next i
End Sub
Private Sub setTextFont()
     Dim cellsList As New Collection
     With cellsList
        .Add Cells(DAY_ROW, NAME_COLUMN)
        .Add Cells(DAY_ROW, NAME_COLUMN + 2)
        .Add Cells(4, NAME_COLUMN)
        .Add Cells(5, NAME_COLUMN)
        .Add Cells(8, WORK_TIME_COLUMN)
        .Add Cells(7, HOLIDAY_COUNT_COLUMN)
        .Add Cells(8, HOLIDAY_COUNT_COLUMN)
        .Add Cells(9, HOLIDAY_COUNT_COLUMN)
        .Add Cells(7, HOLIDAY_COUNT_COLUMN + 1)
        .Add Cells(8, HOLIDAY_COUNT_COLUMN + 1)
        .Add Cells(9, HOLIDAY_COUNT_COLUMN + 1)
    End With
    
    Dim i As Integer
    Dim rg As Range
    For i = 1 To cellsList.count
        Set rg = cellsList(i)
        rg.Font.Size = 9
    Next i
    
    Range(Cells(1, 2), Cells(1, 3)).Font.Size = 15
    
    Range(Cells(3, LEFT_EDGE_OF_SHIFT), Cells(41, RGHIT_EDGE_OF_SHIFT)).Font.Size = 9
 
End Sub

Private Sub setTextFormat()
'���t�s�̕\���`���̐ݒ�
    Dim i As Integer
    Dim rg As Range
    For i = LEFT_EDGE_OF_SHIFT To RGHIT_EDGE_OF_SHIFT
        Set rg = Range(Cells(3, i), Cells(3, i))
        rg.NumberFormatLocal = "###.#""��"""
        Set rg = Range(Cells(4, i), Cells(4, i))
        rg.NumberFormat = "d""��"""
    Next i
'�J�����ԗ�̕\���`���̐ݒ�
    Range(Cells(10, WORK_TIME_COLUMN), Cells(41, WORK_TIME_COLUMN)).NumberFormatLocal = "##""����"""
End Sub

Private Sub setAlignment()
    Dim i As Integer
    Dim rg As Range
    For i = 10 To 40 Step 2
        Set rg = Range(Cells(i, LEFT_EDGE_OF_SHIFT), Cells(i, RGHIT_EDGE_OF_SHIFT))
        rg.HorizontalAlignment = xlLeft
    Next i
End Sub
   
    

