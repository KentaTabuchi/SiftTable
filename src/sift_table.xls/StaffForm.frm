VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StaffForm 
   Caption         =   "希望シフト"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11715
   OleObjectBlob   =   "StaffForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "StaffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'従業員情報を一覧表示するフォーム
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Dim スタッフ As Staff
    Dim record(5) As String
    Dim i As Integer
    For Each スタッフ In TableManager.スタッフリスト
        Debug.Print i
        With Me.ListBox_StaffInfo
            .AddItem ("")
            .List(i, 0) = スタッフ.名前
            .List(i, 1) = スタッフ.希望出勤回数
            .List(i, 2) = スタッフ.出勤不可曜日
            .List(i, 3) = スタッフ.備考
        End With
    i = i + 1
    Next
End Sub
