Sub CalculateRetirement()
    Dim currentDate  '当前日期
    Dim birthDate  '生日
    Dim retirementAge '退休年限
    Dim gundanDay '退休日期
    Dim daysDiff '退休日期差

    currentDate = DateValue(Year(Now) & "-" & Month(Now) & "-" & Day(Now)) 
    retirementAge = 60
    ' 手动更改出生年月
    birthDate = DateValue("2000-05-17")
    gundanDay = DateValue(Year(DateAdd("yyyy", 60, birthDate)) & "-" & Month(Now) & "-" & Day(Now)) 
    daysDiff = DateDiff("d", currentDate, gundanDay)
    MsgBox "退休还有" & daysDiff & "天"
End Sub

CalculateRetirement