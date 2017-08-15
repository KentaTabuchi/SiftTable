Attribute VB_Name = "WorkSheetWriter"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'データをワークシートに転記するメソッドを纏めたモジュール
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Enum 列
    開始列 = 3
    前月締日 = 7
    月初 = 8   '１６日の事
    最終列 = 39
    労働時間列 = 40
    公休列 = 41
    週休列 = 42
End Enum
Private Enum 行
    日付行 = 4
End Enum
'社員締日の列がシートの何列目にあるかを検索して返す
Private Property Get 社員締日列() As Integer

    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    For i = 列.開始列 To 列.最終列
        targetDay = Cells(行.日付行, i)
        If Day(targetDay) = 10 Then
            社員締日列 = i
        End If
    Next i
    Application.ScreenUpdating = True
   
End Property
'バイト締日の列がシートの何列目にあるかを検索して返す
Private Property Get バイト締日列() As Integer

    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    For i = 列.開始列 To 列.最終列
        targetDay = Cells(行.日付行, i)
        
        If Day(targetDay) = 15 Then
            バイト締日列 = i
        End If
    
    Next i
    Application.ScreenUpdating = True

End Property
'引数に渡された出退勤時間をアクティブセルに記入する
Public Sub WriteTimeCard(card As TimeCard)
    If card.出勤時間 = "クリア" Then
        Selection = ""
        ActiveCell.Offset(1, 0).Activate
        Selection = ""
        ActiveCell.Offset(-1, 1).Activate
    Else
        Selection = card.出勤時間
        ActiveCell.Offset(1, 0).Activate
        Selection = card.退勤時間
        ActiveCell.Offset(-1, 1).Activate
    End If
End Sub
'引数に渡されたスタッフの基本シフトを転記
Public Sub WriteBasicShift(スタッフ As Staff)

    If スタッフ.名前 = "" Then  '名前列が空欄の場合はその行のシフト欄を空にする
        Dim j As Integer
        
        Select Case スタッフ.職位
            Case True '社員の場合
                For j = 列.開始列 To 社員締日列
                Cells(スタッフ.row, j).Value = ""
                Cells(スタッフ.row + 1, j).Value = ""
                Next j
            Case False 'バイトの場合
                For j = 列.月初 To バイト締日列
                Cells(スタッフ.row, j).Value = ""
                Cells(スタッフ.row + 1, j).Value = ""
                Next j
        End Select
    ElseIf スタッフ.名前 = "不足" Then
        'この条件では何もしないのでこの行は空欄で合っている
    Else
    Dim currentDay As String '検索範囲の日付を順次代入
    Dim column As Integer
    Dim 月曜日 As TimeCard
    Dim 火曜日 As TimeCard
    Dim 水曜日 As TimeCard
    Dim 木曜日 As TimeCard
    Dim 金曜日 As TimeCard
    Dim 土曜日 As TimeCard
    Dim 日曜日 As TimeCard
    On Error Resume Next
    Set 月曜日 = スタッフ.基本シフト.Item("月曜日")
    Set 火曜日 = スタッフ.基本シフト.Item("火曜日")
    Set 水曜日 = スタッフ.基本シフト.Item("水曜日")
    Set 木曜日 = スタッフ.基本シフト.Item("木曜日")
    Set 金曜日 = スタッフ.基本シフト.Item("金曜日")
    Set 土曜日 = スタッフ.基本シフト.Item("土曜日")
    Set 日曜日 = スタッフ.基本シフト.Item("日曜日")
    
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = 列.月初 To 列.最終列
            currentDay = DateValue(Cells(行.日付行, i).Value)
                Select Case Weekday(currentDay)
                Case vbSunday
                    Cells(スタッフ.row, i).Value = 日曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 日曜日.退勤時間
                Case vbMonday
                    Cells(スタッフ.row, i).Value = 月曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 月曜日.退勤時間
                Case vbTuesday
                    Cells(スタッフ.row, i).Value = 火曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 火曜日.退勤時間
                Case vbWednesday
                    Cells(スタッフ.row, i).Value = 水曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 水曜日.退勤時間
                Case vbThursday
                    Cells(スタッフ.row, i).Value = 木曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 木曜日.退勤時間
                Case vbFriday
                    Cells(スタッフ.row, i).Value = 金曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 金曜日.退勤時間
                Case vbSaturday
                    Cells(スタッフ.row, i).Value = 土曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 土曜日.退勤時間
        End Select
        Next i
    Application.ScreenUpdating = True
    End If
End Sub
'引数に渡されたスタッフの基本シフトを期間を指定して転記
Public Sub WriteBasicShiftByTurn(スタッフ As Staff, 開始日 As Date, 最終日 As Date)

    If スタッフ.名前 = "" Then  '名前列が空欄の場合はその行のシフト欄を空にする
        Dim j As Integer
        
        Select Case スタッフ.職位
            Case True '社員の場合
                For j = 列.開始列 To 社員締日列
                Cells(スタッフ.row, j).Value = ""
                Cells(スタッフ.row + 1, j).Value = ""
                Next j
            Case False 'バイトの場合
                For j = 列.月初 To バイト締日列
                Cells(スタッフ.row, j).Value = ""
                Cells(スタッフ.row + 1, j).Value = ""
                Next j
        End Select
    ElseIf スタッフ.名前 = "不足" Then
        'この条件では何もしないのでこの行は空欄で合っている
    Else
    
    Dim currentDay As String '検索範囲の日付を順次代入
    Dim 開始日列 As Integer
    Dim 最終日列 As Integer
    Dim column As Integer
    
    Dim 月曜日 As TimeCard
    Dim 火曜日 As TimeCard
    Dim 水曜日 As TimeCard
    Dim 木曜日 As TimeCard
    Dim 金曜日 As TimeCard
    Dim 土曜日 As TimeCard
    Dim 日曜日 As TimeCard
    On Error Resume Next
    Set 月曜日 = スタッフ.基本シフト.Item("月曜日")
    Set 火曜日 = スタッフ.基本シフト.Item("火曜日")
    Set 水曜日 = スタッフ.基本シフト.Item("水曜日")
    Set 木曜日 = スタッフ.基本シフト.Item("木曜日")
    Set 金曜日 = スタッフ.基本シフト.Item("金曜日")
    Set 土曜日 = スタッフ.基本シフト.Item("土曜日")
    Set 日曜日 = スタッフ.基本シフト.Item("日曜日")
    
    '開始日と最終日が表の何列目にあるのか検索する
    Dim k
    For k = 列.月初 To 列.最終列
        currentDay = DateValue(Cells(行.日付行, k).Value)
        If 開始日 = CDate(currentDay) Then
            開始日列 = k
        End If
        If 最終日 = CDate(currentDay) Then
            最終日列 = k
        End If
    Next k
    
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = 開始日列 To 最終日列
            currentDay = DateValue(Cells(行.日付行, i).Value)
                Select Case Weekday(currentDay)
                Case vbSunday
                    Cells(スタッフ.row, i).Value = 日曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 日曜日.退勤時間
                Case vbMonday
                    Cells(スタッフ.row, i).Value = 月曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 月曜日.退勤時間
                Case vbTuesday
                    Cells(スタッフ.row, i).Value = 火曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 火曜日.退勤時間
                Case vbWednesday
                    Cells(スタッフ.row, i).Value = 水曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 水曜日.退勤時間
                Case vbThursday
                    Cells(スタッフ.row, i).Value = 木曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 木曜日.退勤時間
                Case vbFriday
                    Cells(スタッフ.row, i).Value = 金曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 金曜日.退勤時間
                Case vbSaturday
                    Cells(スタッフ.row, i).Value = 土曜日.出勤時間
                    Cells(スタッフ.row + 1, i).Value = 土曜日.退勤時間
        End Select
        Next i
    Application.ScreenUpdating = True
    End If
End Sub
'引数に渡されたスタッフの前月とかぶる部分を前月のシートからコピーする
Public Sub CopyFromPreviousMonth(スタッフ As Staff)
    If スタッフ.名前 = "" Then
    Else
    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    For i = 列.開始列 To 列.前月締日
        targetDay = Cells(行.日付行, i)
        Dim card As TimeCard
        For Each card In スタッフ.前月シフト
        If card.日付 = targetDay Then
            Cells(スタッフ.row, i).Value = card.出勤時間
            Cells(スタッフ.row + 1, i).Value = card.退勤時間
        End If
        Next
    Next i
    Application.ScreenUpdating = True
    End If
End Sub
'[開発中]引数に渡されたスタッフの次月とかぶる部分を次月のシートからコピーする
'締日の列を計算で算出する必要がある。プロパティの新設が必要
Public Sub CopyFromNextMonth(スタッフ As Staff)
    If スタッフ.名前 = "" Then
    Else
    Dim rg As Range
    Dim targetDay As Date
    Dim i As Integer
    Application.ScreenUpdating = False
    
    Dim 締日 As Integer
        If スタッフ.職位 = True Then
            締日 = 社員締日列
        ElseIf スタッフ.職位 = False Then
            締日 = バイト締日列
        End If
    For i = 締日 + 1 To 列.最終列
        targetDay = Cells(行.日付行, i)
        Dim card As TimeCard
        For Each card In スタッフ.次月シフト
        If card.日付 = targetDay Then
            Cells(スタッフ.row, i).Value = card.出勤時間
            Cells(スタッフ.row + 1, i).Value = card.退勤時間
        End If
        Next
    Next i
    Application.ScreenUpdating = True
    End If
End Sub
'ワークシートに祝日を書き込む
Public Sub WriteLegalHoliday(祝日 As Schedule)
    Dim targetDay As Date '検索範囲の日付を順次代入
    Application.ScreenUpdating = False
        Dim i As Integer
        For i = 列.月初 To 列.最終列
            targetDay = Cells(行.日付行, i).Value
            Dim イベント As Events
            For Each イベント In 祝日.祝日リスト
                If イベント.日付 = targetDay Then
                    Cells(祝日.作業行, i).Value = イベント.内容
                End If
            Next
        Next i
    Application.ScreenUpdating = True
End Sub
'ワークシートに会議等の予定を書き込む
Public Sub WriteMeetingDay(会議リスト As Schedule)
    Dim targetDay As Date '検索範囲の日付を順次代入
    Application.ScreenUpdating = False
    Range(Cells(会議リスト.作業行, 列.開始列), Cells(会議リスト.作業行 + 1, 列.最終列)).Select
    Selection.Interior.Color = vbWhite
    Selection.ClearContents
        Dim i As Integer
        For i = 列.月初 To 列.最終列
            targetDay = Cells(行.日付行, i).Value
            Dim イベント As Events
            For Each イベント In 会議リスト.会議等
                If イベント.日付 = targetDay Then
                    Cells(会議リスト.作業行, i).Value = イベント.内容
                    Range(Cells(会議リスト.作業行, i), (Cells(会議リスト.作業行 + 1, i + イベント.期間))).Interior.Color = vbYellow
                  
                End If
            Next
        Next i
    Application.ScreenUpdating = True
End Sub
'ワークシートの労働時間欄に労働時間を書き込むメソッド
Public Sub WriteWorkTime(スタッフ As Staff)
    Cells(スタッフ.row, 列.労働時間列).Value = スタッフ.月間労働時間
End Sub
'ワークシートに公休、週休回数を書き込むメソッド
Public Sub WriteNumOfPublicHoliday(スタッフ As Staff)

    Dim rules As CampanyRules
    Set rules = New CampanyRules
    Dim 日付 As Date: 日付 = Cells(2, 2)
    Dim 所定公休 As Byte
    Dim 所定週休 As Byte
    Select Case スタッフ.職位
    Case False
        所定公休 = 0
        所定週休 = 0
    Case True
    所定公休 = rules.GetGivenPublicHolidays(Date)
    所定週休 = rules.GetGivenWeeklyHolidays(Date)
    End Select
    
    Cells(スタッフ.row, 列.公休列).Value = スタッフ.公休回数
    Cells(スタッフ.row, 列.週休列).Value = スタッフ.週休回数
    Cells(スタッフ.row + 1, 列.公休列).Value = 所定公休
    Cells(スタッフ.row + 1, 列.週休列).Value = 所定週休
End Sub


