' �Թ���Ա����
Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if


Set objShell = CreateObject("WScript.Shell")

' ���� yes�ػ� noȡ���ػ�
Dim result
result = objShell.Popup("�ǣ��ػ���\��ȡ���ػ���", 0, "MoYuWang", vbQuestion + vbYesNo)

' �жϷ�֧
If result = vbYes Then
    Shutdown
ElseIf result = vbNo Then
    CancelShutdown
Else
    MsgBox "��ʱ"
End If


' ȡ���ػ�
Sub CancelShutdown()
    Dim result
    result = objShell.Run("shutdown -a", 0, True)
    If result = 0 Then
        MsgBox "ȡ���ɹ�"
    Else
        MsgBox "��ǰδ���ö�ʱ�ػ�"
    End If
End Sub


' �ػ�
Sub Shutdown()
    Dim timeInput
    timeInput = InputBox("������Ԥ���ػ�ʱ�䣨���ӣ�", "MoYuWang")
    ' �жϰ�ťѡ��
    If timeInput <> "" Then
        Dim timeMinutes
        timeMinutes = CInt(timeInput)
        
        ' ?�ػ�����
        Dim result
        result = objShell.Run ("shutdown -s -t " & timeMinutes * 60, 0, True)

        If result = 0 Then
            ' MsgBox "����" & timeMinutes & "���Ӻ�ػ�"
        Else
            MsgBox "�ػ�����ִ��ʧ��"
        End If

    Else
        ' ȡ������
        MsgBox "ȡ����ʱ�ػ�"
    End If
End SUb





