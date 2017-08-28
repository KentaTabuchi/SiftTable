VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CampanyRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'社内規定（所定公休回数や労働時間上限など）を格納するクラス

'所定公休を返すメソッド
Public Function GetGivenPublicHolidays(日付 As Date) As Byte
    Dim 年 As Integer
    Dim 月 As Integer
    年 = Year(日付)
    月 = month(日付)
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 公週休回数 WHERE 年度 = " & 年 & "AND  月度 = " & 月)
    On Error Resume Next
    GetGivenPublicHolidays = DatabaseIO.adoRs!公休
    DatabaseIO.DBClose
    Debug.Print "日付="; 日付; "年="; 年; "月="; 月; "公休="; GetGivenPublicHolidays
End Function
'所定週休を返すメソッド
Public Function GetGivenWeeklyHolidays(日付 As Date) As Byte
    Dim 年 As Integer
    Dim 月 As Integer
    年 = Year(日付)
    月 = month(日付)
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM 公週休回数 WHERE 年度 = " & 年 & "AND  月度 = " & 月)
    On Error Resume Next
    GetGivenWeeklyHolidays = DatabaseIO.adoRs!週休
    DatabaseIO.DBClose
    Debug.Print "日付="; 日付; "年="; 年; "月="; 月; "週休="; GetGivenWeeklyHolidays
End Function
