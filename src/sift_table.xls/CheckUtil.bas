Attribute VB_Name = "CheckUtil"
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'������̃`�F�b�N�Ȃǂ����郁�\�b�h
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'�V�[�g���������Ȃ�true��Ԃ�
Public Function ChkWsName() As Boolean
    Dim ws As Worksheet
    Dim wsname As String '�A�N�e�B�u�V�[�g�̖��O���i�[
    Dim chkname As String
    Dim i As Integer
    Set ws = ActiveSheet
    wsname = ws.name
    ChkWsName = False
    For i = 1 To 12
        chkname = (i & "��")
        If wsname Like chkname Then
        ChkWsName = True
        End If
    Next i
End Function
'�����̃V�[�g������O���̃V�[�g�̖��O�𐮌`���ĕԂ�
Public Function PreviousSheetName(currentSheetName As String) As String
    currentSheetName = Val(currentSheetName)
    If currentSheetName = 1 Then
        currentSheetName = 12
    Else
        currentSheetName = Int(currentSheetName) - 1
    End If
    currentSheetName = str(currentSheetName) & "��"
    Debug.Print currentSheetName
    PreviousSheetName = Trim(currentSheetName)
End Function
'�����̃V�[�g�����玟���̃V�[�g�̖��O�𐮌`���ĕԂ�
Public Function NextSheetName(currentSheetName As String) As String
    currentSheetName = Val(currentSheetName)
    If currentSheetName = 12 Then
        currentSheetName = 1
    Else
        currentSheetName = Int(currentSheetName) + 1
    End If
    currentSheetName = str(currentSheetName) & "��"
    Debug.Print currentSheetName
    NextSheetName = Trim(currentSheetName)
End Function
