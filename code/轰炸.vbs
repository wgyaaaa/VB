Set WshShell=WScript.CreateObject("WScript.Shell")

for i=1 to 15

WScript.Sleep 100

WshShell.SendKeys "^v"

WshShell.SendKeys i

WshShell.SendKeys "%s"

Next