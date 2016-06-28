set ie=wscript.createobject("internetexplorer.application","event_") '创建ie对象'
ie.menubar=0 '取消菜单栏'
ie.addressbar=0 '取消地址栏'
ie.toolbar=0 '取消工具栏'
ie.statusbar=0 '取消状态栏'
ie.width=500 '宽400'
ie.height=600 '高400'
ie.resizable=0 '不允许用户改变窗口大小'
ie.navigate "about:blank" '打开空白页面'
ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) '水平居中'
ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2) '垂直居中'
ie.visible=1 '窗口可见'
attachfile = ""

with ie.document 
.write "<html><body bgcolor=#dddddd scroll=no>" 
.write "<h2 align=center>密送群发邮件</h2><br>"
.write "<p>主题  ：<input id=theme type=text size=30>" 
.write "<p>正文  ：<input type=file name=fileField class=file id=text accept='.txt'/>" 
'.write "<input id=addattach type=button value=添加附件 onclick=alert(document.getElementByIdx_x('attach1').value);/>"
.write "<p>附件1 ：<input type=file name=fileField class=file id=attach1 >" 
'.write "<p>附件2 ：<input type=file name=fileField class=file id=attach2/>" 
'.write "<p>附件2 ：<form name=thisform method=post action='<%=request.getContextPath()%>/movieManage.do' id=thisform enctype=multipart/form-data> <input type=file name=theFile onchange=document.getElementById('theFilePath').value=this.value/> <input type=hidden id=theFilePath name=theFilePath></form> "
.write "<p>附件2 ：<input type=file name=fileField class=file id=attach2 onchange=document.getElementByIdx_x('attach2').value=this.value/>" 
.write "<p>附件3 ：<input type=file name=fileField class=file id=attach3/>" 
.write "<p>账号  ：<input id=user type=text size=15 value=liuyu> @mail.hust.edu.cn" 
.write "<p>发件人：<input id=usrshown type=text size=12 value=刘玉>" 
.write "<p>密码  ：<input id=pass type=password size=30>"
.write "<p align=center><br>" 
.write "<input id=confirm type=button value=确定>"
.write "<input id=cancel type=button value=取消>"
.write "</body></html>"
end with

dim wmi '显式定义一个全局变量'
set wnd=ie.document.parentwindow '设置wnd为窗口对象'
set id=ie.document.all '设置id为document中全部对象的集合'
id.confirm.onclick=getref("confirm") '设置点击"确定"按钮时的处理函数'
id.cancel.onclick=getref("cancel") '设置点击"取消"按钮时的处理函数'

do while true '由于ie对象支持事件，所以相应的，'
wscript.sleep 200 '脚本以无限循环来等待各种事件。'
loop

sub event_onquit 'ie退出事件处理过程'
wscript.quit '当ie退出时，脚本也退出'
end sub

sub cancel '"取消"事件处理过程'
ie.quit '调用ie的quit方法，关闭IE窗口'
end sub '随后会触发event_onquit，于是脚本也退出了'

sub confirm '"确定"事件处理过程，这是关键'
with id
wscript.echo .theme.value
wscript.echo .user.value
wscript.echo .user.value
wscript.echo .usrshown.value
wscript.echo .pass.value
wscript.echo .attach1.value
file = ie.document.getElementById("attach1").value
fakepath = left(file,12)
name = replace(file,fakepath,"")
wscript.echo name
end with
end sub

sub clearlog(name)
wql="select * from Win32_NTEventLogFile where logfilename='"&name&"'"
set logs=wmi.execquery(wql) '注意，logs的成员不是每条日志，'
for each l in logs '而是指定日志的文件对象。'
if l.cleareventlog() then
wnd.alert("清除日志"&name&"时出错！")
ie.quit
wscript.quit
end if
next
end sub