msgbox"编写者 LanMengCha，版本V2.20（结束更新）"
msgbox"接下来请输入命令，指南请输入：0"
Dim link,num1,num0
do
Set loopmsg=WScript.CreateObject("Wscript.Shell")
link=inputbox("输入命令","设置本次[link]变量值")
If link = 1 Then
msgbox"请复制要刷的内容,然后输入次数,打开需要刷屏的窗口"
Wscript.Sleep 50
num1=inputbox("输入要刷屏的次数","设置本次[num1]变量值")
For adder = 1 to num1
loopmsg.SendKeys "^v"
Wscript.Sleep 50
loopmsg.SendKeys "%s"
Next
ElseIf link=2 Then
msgbox"请复制要刷的内容,然后输入数量，打开需要刷屏的窗口。"
num0=inputbox("输入要刷屏的次数","设置本次[num0]变量值")
For adder = 1 to num0
loopmsg.SendKeys "^v"
Wscript.Sleep 50
loopmsg.SendKeys "{ENTER}"
Next
loopmsg.SendKeys "%s"
Elseif link=0 Then
msgbox"命令集显示"
msgbox"命令'1'分条刷屏，命令'2'单条刷屏，命令'0'显示命令集"& vbCrLf &"命令'3'打开任务管理器查询ip，关闭窗口则输入命令之外的数字，后续命令欢迎投稿"
Elseif link=3 Then
msgbox"查询ip，请打开要查询ip的QQ好友（私聊也行）窗口"
loopmsg.SendKeys "^+{ESC}"
MsgBox"请保持窗口焦点顺序为：本程序-QQ聊天-任务管理器,并关闭除了需要查询ip的窗口之外的所有聊天窗口"
For adder = 1 to 66
loopmsg.SendKeys "^v"
Wscript.Sleep 50
loopmsg.SendKeys "{ENTER}"
Next
loopmsg.SendKeys "%s"
Wscript.Sleep 50
MsgBox"请在任务管理器中选择'性能'然后选择'资源监视器'找到'网络'里的QQ.exe"& vbCrLf &"在'网络活动'一栏可以看到消息量最大的第一行的ip就是需要查询的"
Else
msgbox"运行结束"
exit do
End If
loop

