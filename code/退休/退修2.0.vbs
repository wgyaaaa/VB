Option Explicit

' ------------.·:'''''''''''''''''''''''''''''''''''''''''''''''''''''''''':·.------------
' ------------: :  __  __      __   __     __        __                    : :------------
' ------------: : |  \/  |  ___\ \ / /_   _\ \      / /__ _  _ __    __ _  : :------------
' ------------: : | |\/| | / _ \\ V /| | | |\ \ /\ / // _` || '_ \  / _` | : :------------
' ------------: : | |  | || (_) || | | |_| | \ V  V /| (_| || | | || (_| | : :------------
' ------------: : |_|  |_| \___/ |_|  \__,_|  \_/\_/  \__,_||_| |_| \__, | : :------------
' ------------: :                                                   |___/  : :------------
' ------------'·:..........................................................:·'------------
'
' ##########  执行逻辑  ###########
' 
'            读取本地文件 
'           /           \
'         存在         不存在
'         /               \
'    读取必要信息       键入必要信息
'         \               /
'              执行逻辑
'                 |
'            存储必要信息
'
'###################################
'
' 文件存储路径为 D:\retire.txt
' 计算逻辑 diff(当前时间, 出生年月+退休年限) 取时间差(包含闰年366天)

Dim filePath
filePath = "D:\file.txt"
isFileExist  filePath ' 调用函数读取文件内容


Sub isFileExist(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        ReadFile  filePath
    Else
        MsgBox "noexist"
    End If
End Sub

' 读取文件内容并提取生日和退休年龄
Sub ReadFile(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile(filePath, 1) ' 1 表示只读模式
    Dim fileContent
    fileContent = file.ReadAll
    file.Close
    Dim lines
    lines = Split(fileContent, vbCrLf) ' 将文件内容按行拆分为数组
    If UBound(lines) >= 1 Then
        ' 提取生日和退休年龄
        Dim birthday, tuixiu
        Dim data
        data = Split(lines(0), "=") ' 按等号拆分数据
        If UBound(data) = 1 Then
            If Trim(data(0)) = "birthday" Then
                birthday = Trim(data(1))
            End If
        End If

        data = Split(lines(1), "=") ' 按等号拆分数据

        If UBound(data) = 1 Then
            If Trim(data(0)) = "tuixiu" Then
                tuixiu = Trim(data(1))
            End If
        End If
        ' MsgBox "birthday: " & birthday & vbCrLf & "tuixiu: " & tuixiu
    Else
        MsgBox "error"
    End If
    Set fso = Nothing
End Sub

' Dim input1, input2
' input1 = InputBox("1")
' input2 = InputBox("2")

' Dim btn1, btn2
' btn1 = MsgBox("1", vbInformation, "1")
' btn2 = MsgBox("2", vbInformation, "2")

Set objShell = CreateObject("WScript.Shell")
Set objForm = objShell.Popup("1", 0, "1", 64 + 1)

If objForm = 1 Then
    MsgBox "1"
ElseIf objForm = 2 Then
    MsgBox "2"
End If