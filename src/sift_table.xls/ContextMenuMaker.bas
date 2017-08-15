Attribute VB_Name = "ContextMenuMaker"
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'右クリックメニューを追加するメソッド群
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'[開発中] 右クリックメニューを追加。多すぎて見ずらいので親メニューは自作であることだけを示し、
'実際のメニューは子メニューに格納する
Private Menu_親メニュー As CommandBarPopup
Public Sub makeContextMenu()
    Application.CommandBars("cell").Reset
    Set Menu_親メニュー = Application.CommandBars("cell").Controls.Add(Type:=msoControlPopup)
    Menu_親メニュー.Caption = "シフト表"
    Dim menuItems As Collection
    Call setSubMenu("基本シフト", "基本シフト_領域指定", "希望シフト", "給与計算", "労働時間", "公休週休", "色セット", "出勤不可日", "罫線修正", "予定セット", "交互に背景塗る")
End Sub
'サブメニューの作成
Private Sub setSubMenu(ParamArray menuNames())
    Dim menuName As Variant
    For Each menuName In menuNames
        Dim subMenu As CommandBarButton
        Set subMenu = Menu_親メニュー.Controls.Add(Type:=msoControlButton, temporary:=True)
        With subMenu
            .Caption = menuName
            .TooltipText = menuName
            .onAction = menuName
        End With
    Next
End Sub
'データベースから各スタッフの基本シフトを吸い上げエクセルシートへ転記する
Private Sub 基本シフト()
    Dim スタッフ As Staff
    Dim 進捗状況 As String
    Dim スタッフ総数 As Integer
    Dim count As Integer: count = 0
    
    TableManager.initialize
    スタッフ総数 = TableManager.スタッフリスト.count
    ProgressForm.Show vbModeless
    進捗状況 = count & "/" & スタッフ総数 & "人　完了"
    ProgressForm.ProgressLabel.Caption = 進捗状況
    
    For Each スタッフ In TableManager.スタッフリスト
        
        If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "処理を中断しました。"
            End
        End If
        
        count = count + 1
            If スタッフ.名前 = "" Then
                進捗状況 = "スタッフ不在、スキップします・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
            ElseIf スタッフ.名前 = "不足" Then
                進捗状況 = "不足欄は変更しません、スキップします・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
            Else
                進捗状況 = スタッフ.名前 & "のシフトを作成中・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
            End If
        ProgressForm.ProgressLabel.Caption = 進捗状況
        DoEvents 'Wait処理？これを書かないと絵画処理が追い付かず何も見えないままプログレス表示が終わってしまう
        Call WorkSheetWriter.WriteBasicShift(スタッフ)
        Call WorkSheetWriter.CopyFromPreviousMonth(スタッフ)
        Call WorkSheetWriter.CopyFromNextMonth(スタッフ)
    Next
    Unload ProgressForm
End Sub
Public Sub 基本シフト_領域指定()
    SelectForm.Show vbModeless
End Sub
Private Sub 希望シフト()
    StaffForm.Show
End Sub
Private Sub 給与計算()
    CostForm.Show
End Sub
'労働時間をシートに記入するメソッド。
'ここではプログレスバーの処理をしているだけ。
'ロジックの本体は　staffクラスの月間労働時間プロパティで計算している。
Private Sub 労働時間()

    Dim スタッフ As Staff
    Dim 進捗状況 As String
    Dim スタッフ総数 As Integer
    Dim count As Integer: count = 0
    TableManager.initialize
    スタッフ総数 = TableManager.スタッフリスト.count
    ProgressForm.Show vbModeless
    進捗状況 = count & "/" & スタッフ総数 & "人　完了"
    ProgressForm.ProgressLabel.Caption = 進捗状況
    
    For Each スタッフ In TableManager.スタッフリスト

        If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "処理を中断しました。"
            End
        End If
        
        count = count + 1
            If スタッフ.名前 = "" Then
                進捗状況 = "スタッフ不在、スキップします・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
            Else
                進捗状況 = スタッフ.名前 & "の月間労働時間を計算中・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
            End If
        ProgressForm.ProgressLabel.Caption = 進捗状況
        DoEvents 'Wait処理。これを書かないと絵画処理が追い付かず何も見えないままプログレス表示が終わってしまう
        Call WorkSheetWriter.WriteWorkTime(スタッフ)
    Next
    Unload ProgressForm
    
End Sub
Private Sub 公休週休()
    Dim スタッフ As Staff
    TableManager.initialize
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetWriter.WriteNumOfPublicHoliday(スタッフ)
    Next
End Sub
'土日に色を塗る
Private Sub 色セット()

    Dim スタッフ As Staff
    TableManager.initialize
    For Each スタッフ In TableManager.スタッフリスト
    Call WorkSheetPainter.SetWeekendColorUnit(スタッフ)
    Call WorkSheetPainter.ToDarkenOutOfTheCurrentMonth(スタッフ)
    Next
    
End Sub
'全員の出勤不可日に色を塗る
Private Sub 出勤不可日()
    Dim 進捗状況 As String
    Dim スタッフ総数 As Integer
    Dim count As Integer: count = 0
    Dim スタッフ As Staff
    
    TableManager.initialize
    スタッフ総数 = TableManager.スタッフリスト.count
    ProgressForm.Show vbModeless
    進捗状況 = count & "/" & スタッフ総数 & "人　完了"
    ProgressForm.ProgressLabel.Caption = 進捗状況
    For Each スタッフ In TableManager.スタッフリスト
        
            If ProgressForm.IsCancel = True Then
            Unload ProgressForm
            MsgBox "処理を中断しました。"
            End
        End If
        
        count = count + 1
        If スタッフ.名前 = "" Then
            進捗状況 = "スタッフ不在、スキップします・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
        ElseIf スタッフ.名前 = "不足" Then
            進捗状況 = "不足人員欄は触りません。スキップします・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
        Else
            進捗状況 = スタッフ.名前 & "の出勤不可日を読み込み中・・・" & vbNewLine & count & "/" & スタッフ総数 & "人　完了"
        End If
        ProgressForm.ProgressLabel.Caption = 進捗状況
        DoEvents
        
        Call WorkSheetPainter.paintImpossibleDay(スタッフ)
    Next
    Unload ProgressForm
End Sub
Private Sub 罫線修正()
    Call WorkSheetRuler.ruleLine
End Sub
Private Sub 予定セット()
    Call WorkSheetWriter.WriteLegalHoliday(TableManager.祝日)
    Call WorkSheetWriter.WriteMeetingDay(TableManager.会議等)
End Sub
Private Sub 交互に背景塗る()
    TableManager.initialize
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetPainter.paintBackColorInTurn(スタッフ)
    Next
End Sub
