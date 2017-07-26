VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuForm 
   Caption         =   "MenuForm"
   ClientHeight    =   4770
   ClientLeft      =   13050
   ClientTop       =   8595
   ClientWidth     =   2160
   OleObjectBlob   =   "MenuForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum 列
    初期位置 = 5
End Enum
Private Enum 行
    Left = 2
    Top = 3
End Enum
Private Sub Button_希望シフト_Click()
    StaffForm.Show
End Sub

Private Sub Button_罫線_Click()
    Call WorkSheetRuler.ruleLine
End Sub

Private Sub Button_公週休_Click()
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetWriter.WriteNumOfPublicHoliday(スタッフ)
    Next
End Sub

Private Sub Button_出勤不可日_Click()
    '全員の出勤不可日に色を塗る
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetPainter.paintImpossibleDay(スタッフ)
    Next
End Sub

Private Sub Button_色セット_Click()
    '土日に色を塗る
    Call WorkSheetPainter.SetWeekendColor
End Sub

Private Sub Button_予定_Click()
    Call WorkSheetWriter.WriteLegalHoliday(TableManager.祝日)
    Call WorkSheetWriter.WriteMeetingDay(TableManager.会議等)
End Sub

Private Sub Button_労働時間_Click()
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetWriter.WriteWorkTime(スタッフ)
    Next
End Sub

Private Sub Button_基本シフト_Click()
    Dim スタッフ As Staff
    For Each スタッフ In TableManager.スタッフリスト
        Call WorkSheetWriter.WriteBasicShift(スタッフ)
        Call WorkSheetWriter.CopyFromPreviousMonth(スタッフ)
    Next
End Sub
Private Sub Button_給与計算_Click()
    CostForm.Show
End Sub

Private Sub CommandButton1_Click()

End Sub

'ユーザーフォームの初期位置をセルから読み出す
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("設定")
    Me.Left = sh.Cells(行.Left, 列.初期位置)
    Me.Top = sh.Cells(行.Top, 列.初期位置)
End Sub
'フォームを閉じた場所をワークシートに保存
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("設定")
    sh.Cells(行.Left, 列.初期位置) = Me.Left
    sh.Cells(行.Top, 列.初期位置) = Me.Top
End Sub


