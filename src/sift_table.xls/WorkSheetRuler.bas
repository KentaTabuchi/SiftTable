Attribute VB_Name = "WorkSheetRuler"
Option Explicit

'////////////////////////////////////////////////////////
'ワークシートに罫線を引くメソッドを纏めたモジュール（開発中)
'//////////////////////////////////////////////////////

Private Enum 行
    表上端 = 3
    日付 = 4
    曜日行 = 5
    祝日行 = 6
    会議行 = 8
    一人目行 = 10
    表下端 = 41
End Enum
Private Enum 列
    名前列 = 2
    開始日 = 3
    最終日 = 39
    労働時間列 = 40
    週休予定列 = 42
End Enum
'テーブルの罫線を引き直すメソッド
'①まず従業員のシフト欄に罫線を引く　外枠＝細実線　中横線＝破線
'②名前欄　外枠太線
'③次に祝日欄と会議欄に線を引く　外枠横線＝太線　外枠縦線=細実線　中横線＝なし
'④祝日行と会議行のタイトル列に罫線を引く　左右＝太線　上下　細実線
'日付欄（月・日・曜日行）　外枠横線＝太線　外枠縦線＝細実線　中横線＝なし
'日付欄に罫線を引く'⑤
'労働時間列に罫線を引く⑥
'最後に表の外枠＝二重線 '⑦
Public Sub ruleLine()
    Call clearAllLine
    TableManager.initialize
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetRuler.ruleLineToStaffPane(スタッフ) '①
        Call WorkSheetRuler.ruleLineToNamePane(スタッフ) '②
        Call WorkSheetRuler.ruleLineToWorkTimePane(スタッフ) '⑥
    Next
    Call WorkSheetRuler.ruleLineToSchedulePane(TableManager.会議等) '③
    Call WorkSheetRuler.ruleLineToSchedulePane(TableManager.祝日)
    Call WorkSheetRuler.ruleLineToScheduleNamePane(TableManager.会議等) '④
    Call WorkSheetRuler.ruleLineToScheduleNamePane(TableManager.祝日)
    Call ruleLineToDatePane '⑤
    Call WorkSheetRuler.ruleLineToAroundTable '⑦
End Sub
'罫線を全部消すメソッド
'１マス余分に消さないと次にVBAで罫線を引こうとすると表の端のあたりで謎のエラーが起きるので多めに消している。
'ゴミデータによるエクセルのバグのようだ
'どうやらゴミデータでなくセル結合に問題があるようだが一応このままにしておく
Private Sub clearAllLine()
    Dim 対象セル As Range
        Set 対象セル = Range(Cells(行.表上端 - 1, 列.名前列 - 1), Cells(行.表下端 + 2, 列.労働時間列 + 2))
        With 対象セル
            .Borders.LineStyle = xlLineStyleNone
        End With
    Application.ScreenUpdating = True
End Sub
'労働時間列から週休列まで罫線を引く
Private Sub ruleLineToWorkTimePane(スタッフ As Staff)
    Application.ScreenUpdating = False
    Dim 対象セル As Range
        Set 対象セル = Range(Cells(スタッフ.row, 列.労働時間列), Cells(スタッフ.row + 1, 列.週休予定列))
        With 対象セル
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlDash
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With
    Application.ScreenUpdating = True
End Sub
'日付欄に罫線を引く　中横線＝破線　縦線＝細実線
Private Sub ruleLineToDatePane()
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim i As Integer
    Dim 対象セル As Range
    For i = 列.名前列 To 列.最終日
        Set 対象セル = Range(Cells(行.表上端, i), Cells(行.曜日行, i))
        With 対象セル
            .Borders(xlInsideHorizontal).LineStyle = xlDash '真ん中の横線に破線を引く
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next i
    Set 対象セル = Range(Cells(行.表上端, 列.開始日), Cells(行.表下端, 列.開始日))
    対象セル.Borders(xlEdgeLeft).Weight = xlMedium
    Set 対象セル = Range(Cells(行.表上端, 列.労働時間列), Cells(行.表下端, 列.労働時間列))
    対象セル.Borders(xlEdgeLeft).Weight = xlMedium
    Application.ScreenUpdating = True
End Sub
'表のいちばん外側の枠に二重線を引く
'表の下端は名前列が結合されているため、そのまま指定すると1004実行時エラー（Null Pointer Exception)になってしまう
'そのため処理を分けて一つ下の行から上辺に引く
Private Sub ruleLineToAroundTable()
    Dim 対象セル As Range
    
        Set 対象セル = Range(Cells(行.表上端, 列.名前列), Cells(行.表下端, 列.週休予定列))
        With 対象セル
            .Borders(xlEdgeTop).LineStyle = xlDouble
            .Borders(xlEdgeLeft).LineStyle = xlDouble
        End With
        
        Set 対象セル = Range(Cells(行.表下端 + 1, 列.名前列), Cells(行.表下端 + 1, 列.週休予定列))
        With 対象セル
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With
        
        Set 対象セル = Range(Cells(行.表上端, 列.週休予定列), Cells(行.表下端, 列.週休予定列))
        With 対象セル
            .Borders(xlEdgeRight).LineStyle = xlDouble
        End With
    
    Application.ScreenUpdating = True
End Sub
'引数で受け取ったスタッフのシフト欄に罫線を引く
Private Sub ruleLineToStaffPane(スタッフ As Staff)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim 対象セル As Range
    Dim 日 As Integer
    For i = 列.開始日 To 列.最終日
        日 = DAY(Cells(行.日付, i))
        Set 対象セル = Range(Cells(スタッフ.row, i), Cells(スタッフ.row + 1, i))
        対象セル.Borders.LineStyle = xlContinuous '上下左右に細実線を引く
        対象セル.Borders(xlInsideHorizontal).LineStyle = xlDash '真ん中の横線に破線を引く
        
        If 日 = 16 Then
            If スタッフ.職位 = False Then
                対象セル.Borders(xlEdgeLeft).Weight = xlThick
                対象セル.Borders(xlEdgeLeft).Color = RGB(255, 0, 0)
            End If
        ElseIf 日 = 11 Then
            If スタッフ.職位 = True Then
                 対象セル.Borders(xlEdgeLeft).Weight = xlThick
                 対象セル.Borders(xlEdgeLeft).Color = RGB(255, 0, 0)
            End If
        End If
        
    Next i
    Application.ScreenUpdating = True
End Sub
Private Sub ruleLineToSchedulePane(スケジュール As Schedule)
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim 対象セル As Range
    For i = 列.開始日 To 列.最終日
        Set 対象セル = Range(Cells(スケジュール.作業行, i), Cells(スケジュール.作業行 + 1, i))
        With 対象セル
            .Borders.LineStyle = xlContinuous '上下左右に細実線を引く
            .Borders(xlInsideHorizontal).LineStyle = xlDash '真ん中の横線に破線を引く
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next i
    Application.ScreenUpdating = True
End Sub
'引数で受け取ったスタッフの名前欄に罫線を引く
Private Sub ruleLineToNamePane(スタッフ As Staff)
    Application.ScreenUpdating = False
    Dim 対象セル As Range
        Set 対象セル = Range(Cells(スタッフ.row, 列.名前列), Cells(スタッフ.row + 1, 列.名前列))
        With 対象セル
            .Borders.LineStyle = xlContinuous '上下左右に細実線を引く
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
    Application.ScreenUpdating = True
End Sub
'スケジュール行のタイトル列に太線を引く
Private Sub ruleLineToScheduleNamePane(スケジュール As Schedule)
    Application.ScreenUpdating = False
    Dim 対象セル As Range
        Set 対象セル = Range(Cells(スケジュール.作業行, 列.名前列), Cells(スケジュール.作業行 + 1, 列.名前列))
        With 対象セル
            .Borders.LineStyle = xlContinuous '上下左右に細実線を引く
            .Borders.Weight = xlMedium '上下左右を太線にする
        End With
    Application.ScreenUpdating = True
End Sub
