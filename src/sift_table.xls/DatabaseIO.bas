Attribute VB_Name = "DatabaseIO"
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�f�[�^�A�N�Z�X�I�u�W�F�N�g
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private dbpath As String
Public adoCn As Object
Public adoRs As Object '���R�[�h�Z�b�g

'�A�N�Z�X�f�[�^�ɐڑ�
Public Sub DBConnect()
dbpath = ThisWorkbook.Path & "\�V�t�g�\.accdb"
Set adoCn = CreateObject("adodb.connection")
adoCn.Open "Provider=microsoft.ace.oledb.12.0;data source=" & dbpath & ";"
End Sub
'SQL�̔��s
Sub OpenAdo(SQL As String)
Set adoRs = CreateObject("adodb.recordset")
adoRs.cursorLocation = 3
    If Not adoCn Is Nothing Then
        DBConnect
    End If
adoRs.Open SQL, adoCn
End Sub
'���U���g�Z�b�g�̔j��
Sub CloseAdo()
    If Not adoRs Is Nothing Then
        adoRs.Close
        Set adoRs = Nothing
    End If
End Sub
'�ؒf
Sub DBClose()
adoCn.Close
Set adoCn = Nothing
End Sub
