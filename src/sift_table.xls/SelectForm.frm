VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "SelectForm"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DATE_FORMAT = "YYYY/MM/DD"

Private Sub CancelButton_Click()
    Unload Me
End Sub

'�e�L�X�g�{�b�N�X�ɓ��͂��ꂽ�e�L�X�g����t�ɐ��`����
Private Sub tidyText(ByRef textBox As MSForms.textBox)

    Dim tempDate As Date
    Dim tempText As String

    tempText = Trim(textBox.text)
    If IsNumeric(Left$(tempText, 4)) <> True Then  '���S���������łȂ���ΔN��������
        tempText = Year(Date) & "/" & tempText
    End If
    If IsDate(tempText) = True Then
        tempDate = CDate(tempText)
        textBox.Tag = CLng(tempDate) '���t��Long�^��Tag�ɕۑ�
        textBox.text = Format$(tempDate, "YYYY/MM/DD")
    Else
        MsgBox "���t����͂��Ă��������B" & vbCrLf & DATE_FORMAT
    End If
    
End Sub

Private Sub EndDate_Text_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call tidyText(Me.EndDate_Text)
End Sub

Private Sub OK_Button_Click()
    
    TableManager.initialize
    Dim �X�^�b�t As Staff
    Dim �Ώێ� As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        If �X�^�b�t.���O = Me.NameListCombo.text Then
            Set �Ώێ� = �X�^�b�t
        End If
    Next
    Call WorkSheetWriter.WriteBasicShiftByTurn(�Ώێ�, Me.StartDate_Text.text, Me.EndDate_Text.text)

End Sub

Private Sub StartDate_Text_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call tidyText(Me.StartDate_Text)
End Sub

'�t�H�[���̃��[�h���ɃX�^�b�t�̖��O�����X�g�ɂ��ăR���{�{�b�N�X�֊i�[����
Private Sub UserForm_Initialize()
    TableManager.initialize
    Dim �X�^�b�t As Staff
    For Each �X�^�b�t In TableManager.�X�^�b�t���X�g
        If �X�^�b�t.���O = "" Or �X�^�b�t.���O = "�s��" Then
            '�󕶎��ƕs�����͗v��Ȃ��̂Ŕ�΂�
        Else
        Me.NameListCombo.AddItem (�X�^�b�t.���O)
        End If
    Next
    
    Me.StartDate_Text.Tag = CLng(Date)
    Me.StartDate_Text.text = Format$(Me.StartDate_Text.Tag, DATE_FORMAT)
End Sub
