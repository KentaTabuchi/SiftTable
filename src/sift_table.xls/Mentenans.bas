Attribute VB_Name = "Mentenans"
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'ワークシートの余計な書式を消すユーティリティー
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'書式を全てクリアする
Sub DeletFormat()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Cells.FormatConditions.Delete
    Next ws
End Sub
'入力規則の削除
Sub DeleteInputRule()
Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Cells.Validation.Delete
    Next ws
End Sub
