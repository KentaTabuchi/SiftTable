VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "ProgressForm"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370
   OleObjectBlob   =   "ProgressForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'キャンセル処理用のフラグ
Public IsCancel As Boolean

Private Sub Label1_Click()

End Sub

Private Sub SubInfoLabel_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.IsCancel = False
End Sub

Private Sub CancelButton_Click()
    Me.IsCancel = True
End Sub
'プログレスバーのロジック。引数に受け取った進捗を描画する
Public Sub updateProgressBar(progress As Single)
    Dim max As Single
    Dim barSizeCoefficient As Single
    barSizeCoefficient = Me.BackProgressBar.Width / 100
    Me.FrontProgressBar.Width = progress * 100 * barSizeCoefficient
    
    Me.progressNum.Caption = str(Int(progress * 100)) & "%"
    DoEvents
End Sub
