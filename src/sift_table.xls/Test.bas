Attribute VB_Name = "Test"
Option Explicit

Public Sub test()
    Dim rules As CampanyRules
    Set rules = New CampanyRules
    Dim ���t As Date: ���t = (#5/1/2017#)
    Dim ���x As Byte
    Dim �T�x As Byte
    ���x = rules.GetGivenPublicHolidays(Date)
    �T�x = rules.GetGivenWeeklyHolidays(Date)
    
End Sub
