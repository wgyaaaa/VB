' >>=========================================================<<
' ||                                                         ||
' ||     __  __    __   __   __        __                    ||
' ||    |  \/  | __\ \ / /   \ \      / /_ _ _ __   __ _     ||
' ||    | |\/| |/ _ \ V / | | \ \ /\ / / _` | '_ \ / _` |    ||
' ||    | |  | | (_) | || |_| |\ V  V / (_| | | | | (_| |    ||
' ||    |_|  |_|\___/|_| \__,_| \_/\_/ \__,_|_| |_|\__, |    ||
' ||                                               |___/     ||
' ||                                                         ||
' >>=========================================================<<
' 
' ##########  logic  ###########
' 
'              msgBox
'           /           \
'        Yes/shutdown No/cancel
'         /               \
'     Enter min      shutdown exist?
'      /   \             /           \
'  shuntdown exit  Cancel shutdown no exist
'
'###################################

' Set WshShell = WScript.CreateObject("WScript.Shell")
' If WScript.Arguments.Length = 0 Then
'   Set ObjShell = CreateObject("Shell.Application")
'   ObjShell.ShellExecute "wscript.exe" _
'     , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
'   WScript.Quit
' End If

' Admin

Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")

' choose box
Dim result
result = objShell.Popup("shutdown/no shutdown", 0, "please choose", vbQuestion + vbYesNo)

If result = vbYes Then
    Shutdown
ElseIf result = vbNo Then
    CancelShutdown
Else
    MsgBox "timeout"
End If


Sub CancelShutdown()
    Dim result
    result = objShell.Run("shutdown -a", 0, True)
    If result = 0 Then
        MsgBox "Shutdown cancelled successfully."
    Else
        MsgBox "Not created or fail"
    End If
End Sub

Sub Shutdown()
    Dim timeInput
    timeInput = InputBox("Please enter the time(mintues)", "MoYuWang")
    If timeInput <> "" Then
        Dim timeMinutes
        timeMinutes = CInt(timeInput)
        Dim result
        result = objShell.Run ("shutdown -s -t -f" & timeMinutes * 60, 0, True)
        If result = 0 Then
            ' MsgBox "???" & timeMinutes & "??Ӻ???"
        Else
            MsgBox "fail"
        End If

    Else
        MsgBox "exit"
    End If
End SUb





