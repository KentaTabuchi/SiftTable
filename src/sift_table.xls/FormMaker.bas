Attribute VB_Name = "FormMaker"
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'シートに出退勤時間を書き込むボタンがあるフォームを作る
'動的にボタンを生成してフォームに張り付ける
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Enum 列
        出勤時間 = 1
        退勤時間 = 2
    End Enum
Private newButton() As TimeCardButton
Sub make_form()
    Dim i As Integer
    Dim btn As Control
    Dim buttons As New Collection
    ReDim newButton(1 To timeCards.count)
        i = 1
        Dim key As Variant
        For Each key In timeCards
        Dim card As TimeCard
        Set card = key
        Dim name As String
        name = card.出勤時間 & "-" & card.退勤時間
            Set btn = TimeCardButtonForm.Controls.Add("Forms.CommandButton.1", name) '第一引数はシステム定数なのでこの通りかかないとダメ
            With btn
                .Top = 5 + (i - 1) * 20
                .Left = 5
                .Height = 20
                .Width = 60
                .Caption = name
            End With

            Set newButton(i) = New TimeCardButton
            Set newButton(i).button = btn
            Set newButton(i).出退勤 = card
            buttons.Add newButton(i)
            i = i + 1
            Next
        With TimeCardButtonForm
            .Height = i * 20 + 10
            .Width = 70
            .Show vbModeless
        End With
End Sub
'ワークシートから作りたいボタンのリストを吸い上げる
Private Property Get timeCards() As Collection
    Dim cards As Collection
    Set cards = New Collection
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("設定")
    Dim maxRow As Integer
    maxRow = sh.Range("A1").End(xlDown).row
    Dim card() As TimeCard
    ReDim card(maxRow) As TimeCard
    Dim i As Integer
    For i = 2 To maxRow
        Set card(i) = New TimeCard
        card(i).出勤時間 = sh.Cells(i, 列.出勤時間).Value
        card(i).退勤時間 = sh.Cells(i, 列.退勤時間).Value
        cards.Add card(i)
    Next i
    Set timeCards = cards
End Property
Sub test6()
    make_form timeCards
End Sub
