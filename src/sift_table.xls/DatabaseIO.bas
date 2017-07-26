Attribute VB_Name = "DatabaseIO"
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'データアクセスオブジェクト
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private dbpath As String
Public adoCn As Object
Public adoRs As Object 'レコードセット

'アクセスデータに接続
Public Sub DBConnect()
dbpath = ThisWorkbook.Path & "\シフト表.accdb"
Set adoCn = CreateObject("adodb.connection")
adoCn.Open "Provider=microsoft.ace.oledb.12.0;data source=" & dbpath & ";"
End Sub
'SQLの発行
Sub OpenAdo(SQL As String)
Set adoRs = CreateObject("adodb.recordset")
adoRs.cursorLocation = 3
    If Not adoCn Is Nothing Then
        DBConnect
    End If
adoRs.Open SQL, adoCn
End Sub
'リザルトセットの破棄
Sub CloseAdo()
    If Not adoRs Is Nothing Then
        adoRs.Close
        Set adoRs = Nothing
    End If
End Sub
'切断
Sub DBClose()
adoCn.Close
Set adoCn = Nothing
End Sub
