Attribute VB_Name = "CheckUtil"
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'文字列のチェックなどをするメソッド
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'シート名が月名ならtrueを返す
Public Function ChkWsName() As Boolean
    Dim ws As Worksheet
    Dim wsname As String 'アクティブシートの名前を格納
    Dim chkname As String
    Dim i As Integer
    Set ws = ActiveSheet
    wsname = ws.name
    ChkWsName = False
    For i = 1 To 12
        chkname = (i & "月")
        If wsname Like chkname Then
        ChkWsName = True
        End If
    Next i
End Function
'当月のシート名から前月のシートの名前を整形して返す
Public Function PreviousSheetName(currentSheetName As String) As String
    currentSheetName = Val(currentSheetName)
    If currentSheetName = 1 Then
        currentSheetName = 12
    Else
        currentSheetName = Int(currentSheetName) - 1
    End If
    currentSheetName = str(currentSheetName) & "月"
    Debug.Print currentSheetName
    PreviousSheetName = Trim(currentSheetName)
End Function
'当月のシート名から次月のシートの名前を整形して返す
Public Function NextSheetName(currentSheetName As String) As String
    currentSheetName = Val(currentSheetName)
    If currentSheetName = 12 Then
        currentSheetName = 1
    Else
        currentSheetName = Int(currentSheetName) + 1
    End If
    currentSheetName = str(currentSheetName) & "月"
    Debug.Print currentSheetName
    NextSheetName = Trim(currentSheetName)
End Function
