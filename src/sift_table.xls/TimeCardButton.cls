VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeCardButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'動的にボタンを作るためのクラス。
'TimeCardButtonFormに張り付けるボタンを定義する
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public WithEvents button As CommandButton
Attribute button.VB_VarHelpID = -1
Public 出退勤 As TimeCard
'TimeCardButtonFormのボタンが押されたときの処理
Private Sub button_Click()
    WorkSheetWriter.WriteTimeCard Me.出退勤
End Sub

