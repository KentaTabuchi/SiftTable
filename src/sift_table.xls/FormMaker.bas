Attribute VB_Name = "FormMaker"
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�V�[�g�ɏo�ދΎ��Ԃ��������ރ{�^��������t�H�[�������
'���I�Ƀ{�^���𐶐����ăt�H�[���ɒ���t����
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Enum ��
        �o�Ύ��� = 1
        �ދΎ��� = 2
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
        name = card.�o�Ύ��� & "-" & card.�ދΎ���
            Set btn = TimeCardButtonForm.Controls.Add("Forms.CommandButton.1", name) '�������̓V�X�e���萔�Ȃ̂ł��̒ʂ肩���Ȃ��ƃ_��
            With btn
                .Top = 5 + (i - 1) * 20
                .Left = 5
                .Height = 20
                .Width = 60
                .Caption = name
            End With

            Set newButton(i) = New TimeCardButton
            Set newButton(i).button = btn
            Set newButton(i).�o�ދ� = card
            buttons.Add newButton(i)
            i = i + 1
            Next
        With TimeCardButtonForm
            .Height = i * 20 + 10
            .Width = 70
            .Show vbModeless
        End With
End Sub
'���[�N�V�[�g�����肽���{�^���̃��X�g���z���グ��
Private Property Get timeCards() As Collection
    Dim cards As Collection
    Set cards = New Collection
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�ݒ�")
    Dim maxRow As Integer
    maxRow = sh.Range("A1").End(xlDown).row
    Dim card() As TimeCard
    ReDim card(maxRow) As TimeCard
    Dim i As Integer
    For i = 2 To maxRow
        Set card(i) = New TimeCard
        card(i).�o�Ύ��� = sh.Cells(i, ��.�o�Ύ���).Value
        card(i).�ދΎ��� = sh.Cells(i, ��.�ދΎ���).Value
        cards.Add card(i)
    Next i
    Set timeCards = cards
End Property
Sub test6()
    make_form timeCards
End Sub
