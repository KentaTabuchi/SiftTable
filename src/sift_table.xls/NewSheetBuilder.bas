Attribute VB_Name = "NewSheetBuilder"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�V�����V�[�g���������郂�W���[��
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'���C�����\�b�h
'�V�K�V�[�g����������

Private Const DAY_ROW = 2
Private Const DAY_COLUMN = 2
Public Sub setNewSheet()
    Call test(11)
    
End Sub
'�J���e�X�g�p�̃��b�N���\�b�h
Private Sub test(monthNo As Integer)
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = addNewSheet(monthNo)
    Call setStartDay(newWorkSheet)
End Sub

'�����ɓn���ꂽ���̖��O�Ɂu���v�����ăV�[�g���ɂ��쐬����
Private Function addNewSheet(monthNo As Integer) As Worksheet
    
    Dim sheetName As String
    sheetName = Trim(str(monthNo) & "��")
    
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = Worksheets.Add()
    newWorkSheet.name = sheetName
    Set addNewSheet = newWorkSheet
    
End Function

Private Sub setStartDay(newWorkSheet As Worksheet)
    Dim text As String
    Dim month_ As Integer
    month_ = Val(newWorkSheet.name)
    text = Year(Date) & "/" & (month_ - 1) & "/" & 11
    newWorkSheet.Cells(DAY_ROW, DAY_COLUMN).Value = text
End Sub
