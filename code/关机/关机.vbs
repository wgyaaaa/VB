' 以管理员运行
Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if


Set objShell = CreateObject("WScript.Shell")

' 弹窗 yes关机 no取消关机
Dim result
result = objShell.Popup("是（关机）\否（取消关机）", 0, "MoYuWang", vbQuestion + vbYesNo)

' 判断分支
If result = vbYes Then
    Shutdown
ElseIf result = vbNo Then
    CancelShutdown
Else
    MsgBox "超时"
End If


' 取消关机
Sub CancelShutdown()
    Dim result
    result = objShell.Run("shutdown -a", 0, True)
    If result = 0 Then
        MsgBox "取消成功"
    Else
        MsgBox "当前未设置定时关机"
    End If
End Sub


' 关机
Sub Shutdown()
    Dim timeInput
    timeInput = InputBox("请输入预定关机时间（分钟）", "MoYuWang")
    ' 判断按钮选择
    If timeInput <> "" Then
        Dim timeMinutes
        timeMinutes = CInt(timeInput)
        
        ' ?关机命令
        Dim result
        result = objShell.Run ("shutdown -s -t " & timeMinutes * 60, 0, True)

        If result = 0 Then
            ' MsgBox "将在" & timeMinutes & "分钟后关机"
        Else
            MsgBox "关机命令执行失败"
        End If

    Else
        ' 取消输入
        MsgBox "取消定时关机"
    End If
End SUb





