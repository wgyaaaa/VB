Sub CalculateRetirement()
    Dim currentDate  '��ǰ����
    Dim birthDate  '����
    Dim retirementAge '��������
    Dim gundanDay '��������
    Dim daysDiff '�������ڲ�

    currentDate = DateValue(Year(Now) & "-" & Month(Now) & "-" & Day(Now)) 
    retirementAge = 60
    ' �ֶ����ĳ�������
    birthDate = DateValue("2000-05-17")
    gundanDay = DateValue(Year(DateAdd("yyyy", 60, birthDate)) & "-" & Month(Now) & "-" & Day(Now)) 
    daysDiff = DateDiff("d", currentDate, gundanDay)
    MsgBox "���ݻ���" & daysDiff & "��"
End Sub

CalculateRetirement