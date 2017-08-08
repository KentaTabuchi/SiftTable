VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "ProgressForm"
   ClientHeight    =   3555
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

Private Sub UserForm_Initialize()
    Me.IsCancel = False
End Sub

Private Sub CancelButton_Click()
    Me.IsCancel = True
End Sub

