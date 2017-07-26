Attribute VB_Name = "Test"
Option Explicit

Public Sub test()
    Dim rules As CampanyRules
    Set rules = New CampanyRules
    Dim “ú•t As Date: “ú•t = (#5/1/2017#)
    Dim Œö‹x As Byte
    Dim T‹x As Byte
    Œö‹x = rules.GetGivenPublicHolidays(Date)
    T‹x = rules.GetGivenWeeklyHolidays(Date)
    
End Sub
