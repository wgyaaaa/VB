Dim info
info= InputBox("请输入 我是大傻逼 否则立刻关机")

If info="我是大傻逼" Then
  MsgBox "算你识相，但还是要关机"
  Set ws = CreateObject("Wscript.Shell")
  ws.run "shutdown /s /f /t 00"
Else
  MsgBox "直接关机"
  Set ws = CreateObject("Wscript.Shell")
  ws.run "shutdown /s /f /t 00"
End If 