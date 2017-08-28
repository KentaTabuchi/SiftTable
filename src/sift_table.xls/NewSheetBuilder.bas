Attribute VB_Name = "NewSheetBuilder"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'新しいシートを準備するモジュール
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'メインメソッド
'新規シートを準備する

Private Const DAY_ROW = 2
Private Const DAY_COLUMN = 2
Public Sub setNewSheet()
    Call test(11)
    
End Sub
'開発テスト用のモックメソッド
Private Sub test(monthNo As Integer)
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = addNewSheet(monthNo)
    Call setStartDay(newWorkSheet)
End Sub

'引数に渡された月の名前に「月」をつけてシート名にし作成する
Private Function addNewSheet(monthNo As Integer) As Worksheet
    
    Dim sheetName As String
    sheetName = Trim(str(monthNo) & "月")
    
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
