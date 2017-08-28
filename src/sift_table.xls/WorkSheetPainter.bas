Attribute VB_Name = "WorkSheetPainter"
Option Explicit
'////////////////////////////////////////////////////////
'ワークシートに色を塗るメソッドを纏めるモジュール
'//////////////////////////////////////////////////////

Private Enum 行
    日付 = 4
    表上端 = 3
    表下端 = 41
End Enum
Private Enum 列
    名前列 = 2
    開始日 = 3
    最終日 = 39
End Enum


'引数で受け取ったスタッフの出勤不可日に色を塗る
Public Sub paintImpossibleDay(スタッフ As Staff)
    Application.ScreenUpdating = False
    Dim currentDay As String '検索範囲の日付を順次代入
    Dim i As Integer
    
    If スタッフ.名前 = "" Then
        'この場合何もしない
    Else
        For i = 列.開始日 To 列.最終日
            currentDay = DateValue(Cells(行.日付, i).Value)
            Dim 出勤不可日 As Variant
                For Each 出勤不可日 In スタッフ.出勤不可日リスト
                    If currentDay = 出勤不可日 Then
                        Range(Cells(スタッフ.row, i), Cells(スタッフ.row + 1, i)).Select
                        Selection.Interior.Color = vbBlack
                    End If
                Next
        Next i
    End If
    Application.ScreenUpdating = True
End Sub
'表全体の土日の列に色を塗る
Public Sub SetWeekendColor()
    Dim column As Integer
    For column = 列.開始日 To 列.最終日
        Range(Cells(行.表上端, column), Cells(行.表下端, column)) _
        .Interior.ColorIndex = JudgeColor(column)
    Next column

End Sub
'引数で指定されたスタッフの行の土日に色を塗る
'土日でなかった場合は更に１列おきにベージュに塗るようにIF分岐させる
Public Sub SetWeekendColorUnit(スタッフ As Staff)
    Dim column As Integer
    Dim weekdayIndex As Integer
    Dim rgbIndex As Long
    
    For column = 列.開始日 To 列.最終日
        weekdayIndex = Weekday(ActiveSheet.Cells(行.日付, column).Value, 1)
        
            Select Case weekdayIndex
            Case vbSunday
                rgbIndex = RGB(255, 153, 204) '薄い赤色
            Case vbSaturday
                rgbIndex = RGB(204, 255, 255) '薄い青色
            Case Else
                If (スタッフ.row - 10) Mod 4 Then
                    rgbIndex = RGB(255, 255, 255) '白
                Else
                    rgbIndex = RGB(255, 255, 153) 'ベージュ
                End If
           End Select
                Range(Cells(スタッフ.row, column), Cells(スタッフ.row + 1, column)).Interior.Color = rgbIndex
    Next column
    Call paintNameColumnAlternate(スタッフ)
End Sub
'名前欄を一行ずつ交互に色分けする
Private Sub paintNameColumnAlternate(スタッフ As Staff)
    Dim rgbIndex As Long
    
    If (スタッフ.row - 10) Mod 4 Then
        rgbIndex = RGB(255, 255, 255) '白
    Else
        rgbIndex = RGB(255, 255, 153) 'ベージュ
    End If
    
    Range(Cells(スタッフ.row, 列.名前列), Cells(スタッフ.row + 1, 列.名前列)).Interior.Color = rgbIndex
End Sub
'@unused
'引数で指定された一列だけ土日色を塗る
Public Sub SetWeekendColorVertical(column As Integer)
         Range(Cells(行.表上端, column), Cells(行.表下端, column)) _
        .Interior.ColorIndex = JudgeColor(column)
End Sub
'@unused
'引数の列の曜日を判定して塗る色を返す
Private Function JudgeColor(column As Integer) As Integer
    Dim weekdayIndex As Integer
    weekdayIndex = Weekday(ActiveSheet.Cells(行.日付, column).Value, 1)
        Select Case weekdayIndex
        Case vbSunday
            JudgeColor = 38
        Case vbSaturday
            JudgeColor = 20
        Case Else
            JudgeColor = 0
        End Select
End Function

'引数で受け取ったスタッフのシフト欄の当月給与に該当しない欄を暗く RGB(200,200,200)=グレー　に塗りつぶす
'社員のとき、２回目の１0日よりあとなら塗る
'バイトのとき、1回目の15日より前なる塗る　２回目は15日よりあとなら塗る
Public Sub ToDarkenOutOfTheCurrentMonth(スタッフ As Staff)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim 対象セル As Range
    Dim 日 As Integer
    Dim カウントフラグ As Integer

    If スタッフ.職位 = True Then
        For i = 列.開始日 To 列.最終日
        日 = DAY(Cells(行.日付, i))
        If 日 = 11 Then
            カウントフラグ = カウントフラグ + 1
        End If
        If カウントフラグ = 2 Then
            Set 対象セル = Range(Cells(スタッフ.row, i), Cells(スタッフ.row + 1, i))
            対象セル.Interior.Color = RGB(150, 150, 150)
       End If
       Next i
    End If

    If スタッフ.職位 = False Then
        For i = 列.開始日 To 列.最終日
        日 = DAY(Cells(行.日付, i))
        If 日 = 16 Then
            カウントフラグ = カウントフラグ + 1
        End If
        If カウントフラグ = 0 Or カウントフラグ = 2 Then
            Set 対象セル = Range(Cells(スタッフ.row, i), Cells(スタッフ.row + 1, i))
            対象セル.Interior.Color = RGB(150, 150, 150)
       End If
       Next i
    End If
     
    Application.ScreenUpdating = True
End Sub
