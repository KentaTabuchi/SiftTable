VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "SelectForm"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

'テキストボックスに入力されたテキストを日付に整形する
Private Sub tidyText(ByRef textBox As MSForms.textBox)

    Dim tempDate As Date
    Dim tempText As String

    tempText = Trim(textBox.text)
    If IsNumeric(Left$(tempText, 4)) <> True Then  '左４桁が数字でなければ年を加える
        tempText = Year(Date) & "/" & tempText
    End If
    If IsDate(tempText) = True Then
        tempDate = CDate(tempText)
        textBox.Tag = CLng(tempDate) '日付をLong型でTagに保存
        textBox.text = Format$(tempDate, "YYYY/MM/DD")
    Else
        MsgBox "日付を入力してください。" & vbCrLf & DATE_FORMAT
    End If
    
End Sub

Private Sub EndDate_Text_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call tidyText(Me.EndDate_Text)
End Sub

Private Sub OK_Button_Click()
    
    TableManager.initialize
    Dim スタッフ As Staff
    Dim 対象者 As Staff
    For Each スタッフ In TableManager.スタッフリスト
        If スタッフ.名前 = Me.NameListCombo.text Then
            Set 対象者 = スタッフ
        End If
    Next
    Call WorkSheetWriter.WriteBasicShiftByTurn(対象者, Me.StartDate_Text.text, Me.EndDate_Text.text)

End Sub

Private Sub StartDate_Text_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call tidyText(Me.StartDate_Text)
End Sub

'フォームのロード時にスタッフの名前をリストにしてコンボボックスへ格納する
Private Sub UserForm_Initialize()
    TableManager.initialize
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        If スタッフ.名前 = "" Or スタッフ.名前 = "不足" Then
            '空文字と不足欄は要らないので飛ばす
        Else
        Me.NameListCombo.AddItem (スタッフ.名前)
        End If
    Next
    
    Me.StartDate_Text.Tag = CLng(Date)
    Me.StartDate_Text.text = Format$(Me.StartDate_Text.Tag, DATE_FORMAT)
End Sub
