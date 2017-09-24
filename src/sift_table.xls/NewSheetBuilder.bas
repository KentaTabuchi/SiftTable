Attribute VB_Name = "NewSheetBuilder"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'新しいシートを準備するモジュール
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'メインメソッド
'新規シートを準備する
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
   year_ = Application.InputBox(prompt:="何年のシフトを作りますか？半角数字で入力してください。　例 2017", Title:="新規シフト作成", Type:=1)
   month_ = Application.InputBox(prompt:="何月のシフトを作りますか？半角数字で入力してください。 例　6", Title:="新規シフト作成", Type:=1)
    If (month_ > 12 Or month_ < 1) Then
        GoTo NumericValueOutOfRangeException
    End If
   Call main(year_, month_)
   Exit Sub
NumericValueOutOfRangeException:
    MsgBox "1から12で指定してください"
    Exit Sub
End Sub
'このクラスのメインメソッド
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

'引数に渡された月の名前に「月」をつけてシート名にし作成する
Private Function addNewSheet(monthNo As Integer) As Worksheet
    On Error GoTo SheetNameException
        Dim sheetName As String
    sheetName = Trim(str(monthNo) & "月")
    
    Dim newWorkSheet As Worksheet
    Set newWorkSheet = Worksheets.Add()
    newWorkSheet.name = sheetName
    Set addNewSheet = newWorkSheet
    Exit Function
SheetNameException:
    MsgBox "既にあるシートと同じ月名は作れません。１年前のシートは削除してください。"
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
'名前列のセルを縦に２コマづつで結合する
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
'列幅の調整
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
'テキストをシートに直接代入
Private Sub setColumnTitleText()
    Cells(4, NAME_COLUMN) = "日付"
    Cells(5, NAME_COLUMN) = "曜日"
    Cells(8, WORK_TIME_COLUMN) = "労働時間"
    Cells(7, HOLIDAY_COUNT_COLUMN) = "週休"
    Cells(8, HOLIDAY_COUNT_COLUMN) = "取得"
    Cells(9, HOLIDAY_COUNT_COLUMN) = "所定"
    Cells(7, HOLIDAY_COUNT_COLUMN + 1) = "公休"
    Cells(8, HOLIDAY_COUNT_COLUMN + 1) = "取得"
    Cells(9, HOLIDAY_COUNT_COLUMN + 1) = "所定"
    Cells(1, 3) = "月度"
End Sub
   '月日、曜日の行に日付を書き込む処理。
Private Sub setDateText()
    '初めの列だけ式が違うので別処理
    Range(Cells(1, 2), Cells(1, 2)).FormulaR1C1 = "=MONTH(R[1]C+30)"
    Range(Cells(3, LEFT_EDGE_OF_SHIFT), Cells(3, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=MONTH(R[1]C)"
    Range(Cells(4, LEFT_EDGE_OF_SHIFT), Cells(4, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=R[-2]C[-1]"
    Range(Cells(5, LEFT_EDGE_OF_SHIFT), Cells(5, LEFT_EDGE_OF_SHIFT)).FormulaR1C1 = "=TEXT(R[-1]C,""aaa"")"
    
    '２列目から最後まで繰り返し
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
'日付行の表示形式の設定
    Dim i As Integer
    Dim rg As Range
    For i = LEFT_EDGE_OF_SHIFT To RGHIT_EDGE_OF_SHIFT
        Set rg = Range(Cells(3, i), Cells(3, i))
        rg.NumberFormatLocal = "###.#""月"""
        Set rg = Range(Cells(4, i), Cells(4, i))
        rg.NumberFormat = "d""日"""
    Next i
'労働時間列の表示形式の設定
    Range(Cells(10, WORK_TIME_COLUMN), Cells(41, WORK_TIME_COLUMN)).NumberFormatLocal = "##""時間"""
End Sub

Private Sub setAlignment()
    Dim i As Integer
    Dim rg As Range
    For i = 10 To 40 Step 2
        Set rg = Range(Cells(i, LEFT_EDGE_OF_SHIFT), Cells(i, RGHIT_EDGE_OF_SHIFT))
        rg.HorizontalAlignment = xlLeft
    Next i
End Sub
   
    

