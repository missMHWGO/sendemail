set ie=wscript.createobject("internetexplorer.application","event_") '����ie����'
ie.menubar=0 'ȡ���˵���'
ie.addressbar=0 'ȡ����ַ��'
ie.toolbar=0 'ȡ��������'
ie.statusbar=0 'ȡ��״̬��'
ie.width=500 '��400'
ie.height=600 '��400'
ie.resizable=0 '�������û��ı䴰�ڴ�С'
ie.navigate "about:blank" '�򿪿հ�ҳ��'
ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) 'ˮƽ����'
ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2) '��ֱ����'
ie.visible=1 '���ڿɼ�'
attachfile = ""

with ie.document 
.write "<html><body bgcolor=#dddddd scroll=no>" 
.write "<h2 align=center>����Ⱥ���ʼ�</h2><br>"
.write "<p>����  ��<input id=theme type=text size=30>" 
.write "<p>����  ��<input type=file name=fileField class=file id=text accept='.txt'/>" 
'.write "<input id=addattach type=button value=��Ӹ��� onclick=alert(document.getElementByIdx_x('attach1').value);/>"
.write "<p>����1 ��<input type=file name=fileField class=file id=attach1 >" 
'.write "<p>����2 ��<input type=file name=fileField class=file id=attach2/>" 
'.write "<p>����2 ��<form name=thisform method=post action='<%=request.getContextPath()%>/movieManage.do' id=thisform enctype=multipart/form-data> <input type=file name=theFile onchange=document.getElementById('theFilePath').value=this.value/> <input type=hidden id=theFilePath name=theFilePath></form> "
.write "<p>����2 ��<input type=file name=fileField class=file id=attach2 onchange=document.getElementByIdx_x('attach2').value=this.value/>" 
.write "<p>����3 ��<input type=file name=fileField class=file id=attach3/>" 
.write "<p>�˺�  ��<input id=user type=text size=15 value=liuyu> @mail.hust.edu.cn" 
.write "<p>�����ˣ�<input id=usrshown type=text size=12 value=����>" 
.write "<p>����  ��<input id=pass type=password size=30>"
.write "<p align=center><br>" 
.write "<input id=confirm type=button value=ȷ��>"
.write "<input id=cancel type=button value=ȡ��>"
.write "</body></html>"
end with

dim wmi '��ʽ����һ��ȫ�ֱ���'
set wnd=ie.document.parentwindow '����wndΪ���ڶ���'
set id=ie.document.all '����idΪdocument��ȫ������ļ���'
id.confirm.onclick=getref("confirm") '���õ��"ȷ��"��ťʱ�Ĵ�����'
id.cancel.onclick=getref("cancel") '���õ��"ȡ��"��ťʱ�Ĵ�����'

do while true '����ie����֧���¼���������Ӧ�ģ�'
wscript.sleep 200 '�ű�������ѭ�����ȴ������¼���'
loop

sub event_onquit 'ie�˳��¼��������'
wscript.quit '��ie�˳�ʱ���ű�Ҳ�˳�'
end sub

sub cancel '"ȡ��"�¼��������'
ie.quit '����ie��quit�������ر�IE����'
end sub '���ᴥ��event_onquit�����ǽű�Ҳ�˳���'

sub confirm '"ȷ��"�¼�������̣����ǹؼ�'
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
set logs=wmi.execquery(wql) 'ע�⣬logs�ĳ�Ա����ÿ����־��'
for each l in logs '����ָ����־���ļ�����'
if l.cleareventlog() then
wnd.alert("�����־"&name&"ʱ����")
ie.quit
wscript.quit
end if
next
end sub