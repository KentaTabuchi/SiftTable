VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CostForm 
   Caption         =   "給与計算"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   OleObjectBlob   =   "CostForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'概算の給料計算の一覧を表示するフォーム
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Me.ListBox_StaffCost.ColumnWidths = "50;50;50;10;"
    TableManager.setTablePosition
    
    Dim スタッフ As Staff
    Dim record(4) As String
    Dim i As Integer
    For Each スタッフ In TableManager.スタッフリスト
        Debug.Print i
        With Me.ListBox_StaffCost
            .AddItem ("")
            .List(i, 0) = スタッフ.名前
            .List(i, 1) = スタッフ.時給
            .List(i, 2) = スタッフ.月間労働時間
            .List(i, 3) = スタッフ.給料
        End With
    i = i + 1
    Next
End Sub
