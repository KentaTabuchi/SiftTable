Attribute VB_Name = "TableManager"
Option Explicit

Enum 行  'シートの行番号に名前をつける
    祝日行 = 6
    予定行 = 8
    スタッフ1 = 10
    スタッフ2 = 12
    スタッフ3 = 14
    スタッフ4 = 16
    スタッフ5 = 18
    スタッフ6 = 20
    スタッフ7 = 22
    スタッフ8 = 24
    スタッフ9 = 26
    スタッフ10 = 28
    スタッフ11 = 30
    スタッフ12 = 32
    スタッフ13 = 34
    スタッフ14 = 36
    スタッフ15 = 38
    スタッフ16 = 40
End Enum

Public スタッフリスト As Collection
Public 祝日 As Schedule
Public 会議等 As Schedule

Public Sub setTablePosition()
    Set スタッフリスト = New Collection
    Dim i As Integer
    For i = 0 To 15
        Dim スタッフ As Staff
        Set スタッフ = New Staff
        スタッフリスト.Add スタッフ
    Next
    スタッフリスト.Item(1).row = 行.スタッフ1
    スタッフリスト.Item(2).row = 行.スタッフ2
    スタッフリスト.Item(3).row = 行.スタッフ3
    スタッフリスト.Item(4).row = 行.スタッフ4
    スタッフリスト.Item(5).row = 行.スタッフ5
    スタッフリスト.Item(6).row = 行.スタッフ6
    スタッフリスト.Item(7).row = 行.スタッフ7
    スタッフリスト.Item(8).row = 行.スタッフ8
    スタッフリスト.Item(9).row = 行.スタッフ9
    スタッフリスト.Item(10).row = 行.スタッフ10
    スタッフリスト.Item(11).row = 行.スタッフ11
    スタッフリスト.Item(12).row = 行.スタッフ12
    スタッフリスト.Item(13).row = 行.スタッフ13
    スタッフリスト.Item(14).row = 行.スタッフ14
    スタッフリスト.Item(15).row = 行.スタッフ15
    スタッフリスト.Item(16).row = 行.スタッフ16
    Set 祝日 = New Schedule
    祝日.作業行 = 行.祝日行
    Set 会議等 = New Schedule
    会議等.作業行 = 行.予定行

End Sub




