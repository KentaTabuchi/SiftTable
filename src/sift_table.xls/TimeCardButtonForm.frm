VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimeCardButtonForm 
   Caption         =   "勤怠入力"
   ClientHeight    =   3165
   ClientLeft      =   13050
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "TimeCardButtonForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "TimeCardButtonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum 列
    初期位置 = 5
End Enum
Private Enum 行
    Left = 5
    Top = 6
End Enum
Private Sub UserForm_Click()
Debug.Print "left="; Me.Left, "top="; Me.Top
End Sub
'ユーザーフォームの初期位置をセルから読み出す
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("設定")
    Me.Left = sh.Cells(行.Left, 列.初期位置)
    Me.Top = sh.Cells(行.Top, 列.初期位置)
End Sub
'フォームを閉じた場所をワークシートに保存
'デストラクタだとMeの位置情報が消えてしまった後で発動するのでここに記述する。
'デストラクタの直前に発動するイベント
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("設定")
    sh.Cells(行.Left, 列.初期位置) = Me.Left
    sh.Cells(行.Top, 列.初期位置) = Me.Top
End Sub

