Dim info
info= InputBox("������ ���Ǵ�ɵ�� �������̹ػ�")

If info="���Ǵ�ɵ��" Then
  MsgBox "����ʶ�࣬������Ҫ�ػ�"
  Set ws = CreateObject("Wscript.Shell")
  ws.run "shutdown /s /f /t 00"
Else
  MsgBox "ֱ�ӹػ�"
  Set ws = CreateObject("Wscript.Shell")
  ws.run "shutdown /s /f /t 00"
End If 