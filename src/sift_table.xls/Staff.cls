VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'一人分の従業員情報を表すクラス
'名前、シフト、社員情報
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Enum 列
    名前列 = 2
    開始日 = 3
    最終日 = 39
End Enum
Private Enum 行
    日付行 = 4
End Enum
Private Enum 職位区分
    バイト = False
    正社員 = True
End Enum
Public row As Integer
'Public 月間シフト As Collection 'TimeCard型の配列が格納されている

Public Property Get 名前()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim rg As Range
    Set rg = ws.Cells(row, 列.名前列)
    名前 = rg.Value
End Property
'メンバ変数の名前をキーにSQLを発行して職位を判別する
Public Property Get 職位() As Boolean
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
    On Error Resume Next
    職位 = DatabaseIO.adoRs!正臨
    DatabaseIO.DBClose
End Property
Public Property Get 出勤不可曜日() As String
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
    On Error Resume Next
    出勤不可曜日 = DatabaseIO.adoRs!不可曜日
    DatabaseIO.DBClose
End Property
Public Property Get 希望出勤回数() As String
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
    On Error Resume Next
    希望出勤回数 = DatabaseIO.adoRs!週○回
    DatabaseIO.DBClose
End Property
Public Property Get 備考() As String
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
    On Error Resume Next
    備考 = DatabaseIO.adoRs!備考
    DatabaseIO.DBClose
End Property
Public Property Get 月間労働時間() As Integer
Application.ScreenUpdating = False
    Dim card As TimeCard
    Dim cnt As Integer
    
    If Me.名前 = "" Then
        '名前が空の行なので何もせずスキップする
    Else
        For Each card In Me.月間シフト
        ProgressForm.SubInfoLabel.Caption = "Scanning....   " & card.日付
        DoEvents
            Select Case 職位
            Case 職位区分.正社員
                If DAY(card.日付) = 10 Then
                    cnt = cnt + 1
                End If
                If cnt = 0 Then
                    月間労働時間 = 月間労働時間 + card.労働時間
                End If
            Case 職位区分.バイト
                If DAY(card.日付) = 15 Then
                    cnt = cnt + 1
                End If
                If cnt = 1 Then
                    月間労働時間 = 月間労働時間 + card.労働時間
                End If
    
            End Select
        Next
    End If
Application.ScreenUpdating = True
End Property
Public Property Get 給料() As Long
    On Error Resume Next
    Select Case 職位
    Case 職位区分.正社員
        DatabaseIO.DBConnect
        DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
        給料 = DatabaseIO.adoRs!月給
        DatabaseIO.DBClose
    Case 職位区分.バイト
        DatabaseIO.DBConnect
        DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
        給料 = DatabaseIO.adoRs!時給 * 月間労働時間
        DatabaseIO.DBClose
    End Select
End Property
Public Property Get 時給() As Integer
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前 = '" & 名前 & "'")
    On Error Resume Next
    時給 = DatabaseIO.adoRs!時給
    DatabaseIO.DBClose
End Property

Public Property Get 公休回数() As Integer
    Dim isClosingDay As Boolean
    Dim card As TimeCard
    公休回数 = 0
    For Each card In Me.月間シフト
        If DAY(card.日付) = 10 Then
            isClosingDay = True
        End If
        If card.出勤時間 = "公休" And isClosingDay = False Then
            公休回数 = 公休回数 + 1
        End If
    Next
End Property
Public Property Get 週休回数() As Integer
    Dim isClosingDay As Boolean
    Dim card As TimeCard
    週休回数 = 0
    For Each card In Me.月間シフト
        If DAY(card.日付) = 10 Then
            isClosingDay = True
        End If
        If card.出勤時間 = "週休" And isClosingDay = False Then
            週休回数 = 週休回数 + 1
        End If
    Next
End Property
'Date型の日付リストを格納したコレクション
Public Property Get 出勤不可日リスト() As Collection
    Debug.Print Me.名前; ":出勤不可日="
    Set 出勤不可日リスト = New Collection
    Dim recordCount As Integer
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT 日付 FROM 出勤不可 WHERE 名前='" & Me.名前 & "'")
    recordCount = DatabaseIO.adoRs.recordCount
    Dim i As Integer
    Dim 出勤不可日 As Date
    If recordCount = 0 Then
        Debug.Print "出勤不可日はありません"
    Else
    For i = 0 To recordCount - 1
        出勤不可日 = adoRs!日付
        出勤不可日リスト.Add 出勤不可日
        adoRs.MoveNext
    Next i
    End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property
'TimeCard型のリストを格納したコレクション
Public Property Get 基本シフト() As Collection
    Set 基本シフト = New Collection
    Dim recordCount As Integer
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 社員マスタ WHERE 名前='" & Me.名前 & "'")
    recordCount = DatabaseIO.adoRs.recordCount
    Dim i As Integer
    Dim 月曜日 As TimeCard
    Dim 火曜日 As TimeCard
    Dim 水曜日 As TimeCard
    Dim 木曜日 As TimeCard
    Dim 金曜日 As TimeCard
    Dim 土曜日 As TimeCard
    Dim 日曜日 As TimeCard
    Set 月曜日 = New TimeCard
    Set 火曜日 = New TimeCard
    Set 水曜日 = New TimeCard
    Set 木曜日 = New TimeCard
    Set 金曜日 = New TimeCard
    Set 土曜日 = New TimeCard
    Set 日曜日 = New TimeCard
    
    If recordCount = 0 Then
        Debug.Print "DBに基本シフトが入っていません"
    Else
    For i = 0 To recordCount - 1
    On Error Resume Next 'データベースにNULLが混ざっているとエラーが発生するが初期値０が入れば問題ないので無視する
        月曜日.出勤時間 = adoRs!月出
        月曜日.退勤時間 = adoRs!月退
        火曜日.出勤時間 = adoRs!火出
        火曜日.退勤時間 = adoRs!火退
        水曜日.出勤時間 = adoRs!水出
        水曜日.退勤時間 = adoRs!水退
        木曜日.出勤時間 = adoRs!木出
        木曜日.退勤時間 = adoRs!木退
        金曜日.出勤時間 = adoRs!金出
        金曜日.退勤時間 = adoRs!金退
        土曜日.出勤時間 = adoRs!土出
        土曜日.退勤時間 = adoRs!土退
        日曜日.出勤時間 = adoRs!日出
        日曜日.退勤時間 = adoRs!日退
        With 基本シフト
            .Add Item:=月曜日, key:="月曜日"
            .Add Item:=火曜日, key:="火曜日"
            .Add Item:=水曜日, key:="水曜日"
            .Add Item:=木曜日, key:="木曜日"
            .Add Item:=金曜日, key:="金曜日"
            .Add Item:=土曜日, key:="土曜日"
            .Add Item:=日曜日, key:="日曜日"
        End With
        adoRs.MoveNext
    Next i
    End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property
'ワークシートに記入されたシフトを[月間シフト]に格納する
Public Property Get 月間シフト() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim rg As Range
    Set rg = ws.Cells(row, 列.開始日)
    Set 月間シフト = New Collection
    Dim i As Integer
    For i = 列.開始日 To 列.最終日
        Dim 単日シフト As TimeCard
        Set 単日シフト = New TimeCard
        単日シフト.日付 = ws.Cells(行.日付行, i).Value
        単日シフト.出勤時間 = ws.Cells(row, i).Value
        単日シフト.退勤時間 = ws.Cells(row + 1, i).Value
        月間シフト.Add 単日シフト
    Next i
End Property
'TimeCard型のリストを格納したコレクション
Public Property Get 前月シフト() As Collection
    On Error GoTo catch:
    Set 前月シフト = New Collection
    Dim 当月シート As Worksheet
    Dim 前月シート As Worksheet
    Set 当月シート = ThisWorkbook.ActiveSheet
    Set 前月シート = Worksheets(CheckUtil.PreviousSheetName(当月シート.name))
    Dim rg As Range
    Dim 前月行 As Integer
    前月行 = 前月シート.Range("B:B").Find(Me.名前).row
    Set rg = 前月シート.Cells(前月行, 列.開始日)
    Dim i As Integer
    For i = 列.開始日 To 列.最終日
        Dim 単日シフト As TimeCard
        Set 単日シフト = New TimeCard
        単日シフト.日付 = 前月シート.Cells(行.日付行, i).Value
        単日シフト.出勤時間 = 前月シート.Cells(前月行, i).Value
        単日シフト.退勤時間 = 前月シート.Cells(前月行 + 1, i).Value
        前月シフト.Add 単日シフト
    Next i
    Exit Property
catch:
   Debug.Print "前月シートが存在しません"
End Property
'TimeCard型のリストを格納したコレクション
Public Property Get 次月シフト() As Collection
    On Error GoTo catch:
    Set 次月シフト = New Collection
    Dim 当月シート As Worksheet
    Dim 次月シート As Worksheet
    Set 当月シート = ThisWorkbook.ActiveSheet
    Set 次月シート = Worksheets(CheckUtil.NextSheetName(当月シート.name))
    Dim rg As Range
    Dim 次月行 As Integer
    
    次月行 = 次月シート.Range("B:B").Find(Me.名前).row
    Set rg = 次月シート.Cells(次月行, 列.開始日)
    Dim i As Integer
    For i = 列.開始日 To 列.最終日
        Dim 単日シフト As TimeCard
        Set 単日シフト = New TimeCard
        単日シフト.日付 = 次月シート.Cells(行.日付行, i).Value
        単日シフト.出勤時間 = 次月シート.Cells(次月行, i).Value
        単日シフト.退勤時間 = 次月シート.Cells(次月行 + 1, i).Value
        次月シフト.Add 単日シフト
    Next i
    Exit Property
catch:
   Debug.Print "次月シートが存在しないか該当する人がいません"
End Property
'引数に渡された日付の出退勤時間を返す
Public Function getShift(日付_ As Date) As TimeCard
    Dim card As TimeCard
    For Each card In Me.月間シフト
        If card.日付 = 日付_ Then
            Debug.Print card.日付; "出勤:"; card.出勤時間; "退勤:"; card.退勤時間
            Set getShift = card
        End If
    Next
End Function


