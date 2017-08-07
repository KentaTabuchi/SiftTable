VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkSheetEventer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'シートイベントをフックしてイベントを発生させるクラス
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public WithEvents workSheetEventer As Application
Attribute workSheetEventer.VB_VarHelpID = -1
'シートアクティブ時のイベント
Private Sub workSheetEventer_SheetActivate(ByVal sh As Object)
    preparedApplication
End Sub
'エクセルブックを開いた時のイベント。
Private Sub workSheetEventer_WorkbookOpen(ByVal Wb As Workbook)
    preparedApplication
End Sub
'シートに変更が起きると発生するイベント。
'Intersect関数でチェンジイベントを発生させる変更箇所を限定している
Private Sub workSheetEventer_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error Resume Next '原因不明のエラーが出るがとりあえず動くので無視しているが、改善するべき。
    If Intersect(Target, sh.Range(Cells(10, 2), Cells(41, 2))) Is Nothing Then
        Exit Sub
    Else
        preparedApplication
    End If
End Sub
'スタッフクラスなどの初期化を行っている｡はじめにこれを動かさないとほとんどのクラスがNULLpointerExceptionになる
Private Sub preparedApplication()
    On Error Resume Next
    TableManager.initialize
End Sub
