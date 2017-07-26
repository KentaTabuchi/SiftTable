VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'一日分の出勤日、出勤時間、退勤時間を表すクラス
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public 日付 As Date
Public 出勤時間 As String
Public 退勤時間 As String

'この日の労働時間をメンバ変数から計算する。
'IsNumericで「公休」等の数字以外の文字が入っていないかチェック
'深夜労働の計算:深夜に出勤するとマイナスになるので24足している
'休憩時間の計算:８時間以上働いたら１時間休んでいるはずという計算
Public Property Get 労働時間() As Integer
    If IsNumeric(出勤時間) And IsNumeric(退勤時間) Then
        労働時間 = 退勤時間 - 出勤時間
        If 労働時間 < -1 Then
            労働時間 = 労働時間 + 24
        End If
        If 労働時間 >= 8 Then
            労働時間 = 労働時間 - 1
        End If
    Else
    労働時間 = 0
    End If
End Property
