VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'予定を表すクラス
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Enum 列
    開始日 = 3
    最終日 = 39
End Enum
Private Enum 行
    日付行 = 4
End Enum
Public 作業行 As Integer

'データベースから予定のレコードセットを取り出しコレクションに格納する
Public Property Get 会議等() As Collection
    Set 会議等 = New Collection
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 予定")
    Dim recordCount As Integer
    recordCount = DatabaseIO.adoRs.recordCount
    If recordCount = 0 Then
        Debug.Print "予定がありません"
    Else
        Dim i As Integer
        For i = 0 To recordCount - 1
        On Error Resume Next 'NULL文字が入ると止まるが続けても空文字が代入されるので無視
            Dim イベント As Events
            Set イベント = New Events
            イベント.日付 = DatabaseIO.adoRs!日付
            イベント.内容 = DatabaseIO.adoRs!会議等
            イベント.期間 = DatabaseIO.adoRs!期間 - 1
            会議等.Add イベント
            adoRs.MoveNext
        Next i
     End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property
'データベースから予定のレコードセットを取り出しコレクションに格納する
Public Property Get 祝日リスト() As Collection
    Set 祝日リスト = New Collection
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 祝日")
    Dim recordCount As Integer
    recordCount = DatabaseIO.adoRs.recordCount
    If recordCount = 0 Then
        Debug.Print "祝日がありません"
    Else
        Dim i As Integer
        For i = 0 To recordCount - 1
            Dim イベント As Events
            Set イベント = New Events
            イベント.日付 = DatabaseIO.adoRs!日付
            イベント.内容 = DatabaseIO.adoRs!祝日
            イベント.期間 = 0
            祝日リスト.Add イベント
            adoRs.MoveNext
        Next i
     End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property

