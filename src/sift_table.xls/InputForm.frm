VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "入力パレット"
   ClientHeight    =   1470
   ClientLeft      =   1050
   ClientTop       =   8595
   ClientWidth     =   8595
   OleObjectBlob   =   "InputForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 8
    card.退勤時間 = 13
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton11_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 22
    card.退勤時間 = 8
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton12_Click()
     Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = ""
    card.退勤時間 = ""
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton13_Click()
    ActiveCell.Offset(-4, 0).Activate
End Sub

Private Sub CommandButton14_Click()
    ActiveCell.Offset(0, -1).Activate
End Sub

Private Sub CommandButton15_Click()
    ActiveCell.Offset(0, 1).Activate
End Sub

Private Sub CommandButton16_Click()
    ActiveCell.Offset(4, 0).Activate
End Sub

Private Sub CommandButton17_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 8
    card.退勤時間 = 18
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton18_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 8
    card.退勤時間 = 22
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton19_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = "公休"
    card.退勤時間 = ""
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton2_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 8
    card.退勤時間 = 17
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton20_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = "週休"
    card.退勤時間 = ""
    Call WorkSheetWriter.WriteTimeCard(card)

End Sub



Private Sub CommandButton24_Click()
    
    
End Sub

Private Sub CommandButton25_Click()


End Sub

Private Sub CommandButton26_Click()


End Sub

Private Sub CommandButton27_Click()

End Sub

Private Sub CommandButton3_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 8
    card.退勤時間 = 12
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton5_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 13
    card.退勤時間 = 17
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton6_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 12
    card.退勤時間 = 17
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton8_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 17
    card.退勤時間 = 22
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

Private Sub CommandButton9_Click()
    Dim card As TimeCard
    Set card = New TimeCard
    card.出勤時間 = 18
    card.退勤時間 = 22
    Call WorkSheetWriter.WriteTimeCard(card)
End Sub

