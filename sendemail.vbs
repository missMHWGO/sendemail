''''''''''''''''''''''''''''''''''''''''''��ȡ����'''''''''''''''''''''''''''''''''''''''''''''''
set ie=wscript.createobject("internetexplorer.application","event_") '����ie����'
Set objShell = WScript.CreateObject("WScript.Shell")

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
.write "<p>����  ��<input id=theme type=text size=30 >" 
.write "<p>����  ��<input type=file name=fileField class=file id=text accept='.txt' >" 
.write "<p>����1 ��<input type=file name=fileField class=file id=attach1 >" 
.write "<p>����2 ��<input type=file name=fileField class=file id=attach2 >" 
.write "<p>����3 ��<input type=file name=fileField class=file id=attach3 >" 
.write "<p>�˺�  ��<input id=user type=text size=15 value=liuyu > @mail.hust.edu.cn" 
.write "<p>�����ˣ�<input id=username type=text size=12 value=���� >" 
.write "<p>����  ��<input id=password type=password size=30 >"
.write "<br><br>" 
.write "<input id=confirm type=button value=ȷ�� >"
.write "<input id=cancel type=button value=ȡ�� >"
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
dim textname
textfile = ie.document.getElementById("text").value
fakepath = left(textfile,12)
textname = replace(textfile,fakepath,"")
WSH.Echo textname

dim filename
attach1file = ie.document.getElementById("attach1").value
filename = replace(attach1file,fakepath,"")
WSH.Echo filename

''''''''''''''''''''''''''''''''''''''''''''' 
WSH.Echo "���Ͳ���"
Set oExcel=CreateObject("excel.application")
Set oWorkBook=oExcel.Workbooks.Open( "C:\Users\Wang Jiajing\Desktop\��\������\���������б�.xls" )
SendEmailALL oWorkBook, textname,filename
oExcel.Quit
''''''''''''''''''''''''''''''''
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



Class CdoMail
  ' ���幫�����������ʼ��
      Public fso, wso, objMsg
    Private Sub Class_Initialize()
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set wso = CreateObject("wscript.Shell")
        Set objMsg = CreateObject("CDO.Message")
    End Sub


' ���÷��������ԣ�4��������Ϊ��STMP�ʼ���������ַ��STMP�ʼ��������˿ڣ�STMP�ʼ�������STMP�û�����STMP�ʼ��������û�����
    ' ���ӣ�Set MyMail = New CdoMail : MyMail.MailServerSet "smtp.qq.com", 443, "yu2n", "P@sSW0rd"
    Public Sub MailServerSet( strServerName, strServerPort, strServerUsername, strServerPassword )
        NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
        With objMsg.Configuration.Fields
            .Item(NameSpace & "sendusing") = 2                      'Pickup = 1(Send message using the local SMTP service pickup directory.), Port = 2(Send the message using the network (SMTP over the network). )
            .Item(NameSpace & "smtpserver") = strServerName         'SMTP Server host name / ip address
            .Item(NameSpace & "smtpserverport") = strServerPort     'SMTP Server port
            .Item(NameSpace & "smtpauthenticate") = 1               'Anonymous = 0, basic (clear-text) authentication = 1, NTLM = 2
            .Item(NameSpace & "smtpusessl") = True          
            .Item(NameSpace & "sendusername") = strServerUsername   '<�������ʼ���ַ>
            .Item(NameSpace & "sendpassword") = strServerPassword   '<�������ʼ�����>
            .Update
        End With
    End Sub
  ' �����ʼ�������������ߵ�ַ��4��������Ϊ���ļ���(���ܿ�)���ռ���(���ܿ�)���������͡��ܼ�����
    Public Sub  MailFromTo( strMailFrom, strMailTo, strMailCc, strMailBCc)
        objMsg.From = strMailFrom   '<�������ʼ���ַ,������������ͬ>
        objMsg.To = strMailTo       '<�������ʼ���ַ>
        objMsg.Cc = strMailCc       '[��������]           
        objMsg.Bcc = strMailBcc     '[�ܼ�����]
    End Sub
' �ʼ��������ã�3���������ǣ��ʼ�����(text/html/url)����ּ���⡢��������(text�ı���ʽ/html��ҳ��ʽ/urlһ���ִ����ҳ�ļ���ַ)
     Public Function MailBody( strType, strMailSubjectStr, strMessage )
        objMsg.Subject = strMailSubjectStr          '<�ʼ���ּ����>
        Select Case LCase( strType )
            Case "text"
                objMsg.TextBody = strMessage        '<�ı���ʽ����>       
            Case "html"
                objMsg.HTMLBody = strMessage        '<html��ҳ��ʽ����>
            Case "url"
                objMsg.CreateMHTMLBody strMessage   '<��ҳ�ļ���ַ>
            Case Else
                objMsg.BodyPart.Charset = "gb2312"   '<�ʼ����ݱ��룬Ĭ��gb2312>   
                objMsg.TextBody = strMessage        '<�ʼ����ݣ�Ĭ��Ϊ�ı���ʽ����>
        End Select
    End Function
  ' ������и���������Ϊ�����б����飬�����ļ���ʹ�� arrPath = Split( strPath & "|", "|")����·����
    Public Function MailAttachment( arrAttachment )
        If Not IsArray( arrAttachment ) Then arrAttachment = Split( arrAttachment & "|", "|")
        For i = 0 To UBound( arrAttachment )
            If fso.FileExists( arrAttachment(i) ) = True Then
                objMsg.Addattachment arrAttachment(i)
            End If
        Next
    End Function  
    ' �����ʼ�
    Public Sub Send()
        'Delivery Status Notifications: Default = 0, Never = 1, Failure = 2, Success 4, Delay = 8, SuccessFailOrDelay = 14
        objMsg.DSNOptions = 0
        objMsg.Fields.update
        objMsg.Send
    End Sub

End Class

Function SendOneEmail(strSendAddr, strAcount, strAccountName, strPasswd,textname,filename)
    Set MyMail = New CdoMail
    '�ʼ����������ļ���ȡ
    TextBodyFileDir = "C:\Users\Wang Jiajing\Desktop\��\������\"&textname
    Set fso=CreateObject("Scripting.FileSystemObject")
    Set TextBodyFile=fso.OpenTextFile(TextBodyFileDir, 1, False, 0)
    TextBodyInfo = TextBodyFile.readall
    TextBodyFile.Close
    '���÷�����(*)����������ַ���������˿ڡ������û�������������
    MyMail.MailServerSet    "mail.hust.edu.cn", 25, strAccountName, strPasswd
    '���üļ������ռ��ߵ�ַ(*)���ļ��ߡ��ռ��ߡ����͸���(�Ǳ���)�����͸���(�Ǳ���)
    MyMail.MailFromTo       strAcount, "", "", strSendAddr
    '�����ʼ�����(*)����������(text/html/url)���ʼ���ּ���⡢�ʼ������ı�
    MyMail.MailBody         "text", ie.document.getElementById("theme").value, TextBodyInfo
    '��Ӹ���(�Ǳ���)������������һ���ļ�·����������һ����������ļ�·��������
    'MyMail.MailAttachment   Split("e:\2016ȫ��������˼ά��ʦ��ѵ֪ͨ.doc|e:\2016ȫ��������˼ά����֪ͨ.doc|e:\2016���һ��ȫ����������������˼ά��ʦ��ѵ�ࣨ20160227��.docx|e:\2016���һ��ȫ����������������˼ά���֪ͨ��20160227).docx", "|")
    MyMail.MailAttachment   Split("C:\Users\Wang Jiajing\Desktop\��\������\"&filename, "|")	
	WSH.Echo filename
	' �����ʼ�(*)
    MyMail.Send
End Function
Function SendEmailToOneSheetAddr(Sheet, uiSheetCnt,textname,filename)
    arrAccountName = array("dian",ie.document.getElementById("user").value)'�������п������ö���˺š�����
    arrAccount = array("dian@hust.edu.cn(Dian Group)",ie.document.getElementById("user").value&"@hust.edu.cn("&ie.document.getElementById("username").value&")")
    arrPasswd = array("diangroup1",ie.document.getElementById("password").value)
    uiCntAddrMax = 40 '��������ÿ���ʼ�������������������
    uiCntAddr = 0
    strSendAddr = ""
    uiRowMax = Sheet.UsedRange.Rows.Count
    WSH.Echo "sheet " & uiSheetCnt & "��������" & uiRowMax
    'wscript.sleep 1*60*1000  '��λms 1����  
    uiMyEmailCnt = 0
    For uiCntRow = 2 To uiRowMax '����ÿһ��
        strCurAddr = Sheet.cells(uiCntRow,3).value 'Email��Ϣ�ڵ�3��
        If  strCurAddr = "" Then
        Else
            strSendAddr = strSendAddr & strCurAddr & ","
            uiCntAddr = uiCntAddr + 1
        End If        
        If uiCntAddr = uiCntAddrMax Then
            '�����ʼ�
            SendOneEmail   strSendAddr, arrAccount(1), arrAccountName(1), arrPasswd(1),textname,filename'����ɸ����˺ŷ���,uiMyEmailCnt
            WSH.Echo "��ǰ�˻� :" & arrAccount(1)
            WSH.Echo "�ѷ����� :" & strSendAddr
            wscript.sleep 2*60*1000  '��λms 2����  
            uiMyEmailCnt = uiMyEmailCnt + 1
            If uiMyEmailCnt = 2 Then '���uiMyEmailCnt������¼�˺Ÿ�����Ҳ����������Ԫ�ظ���
                uiMyEmailCnt = 0
            End If
            strSendAddr = ""
            uiCntAddr = 0
        End If
    Next
    
    If uiCntAddr > 0 Then
        '�����ʼ�
        SendOneEmail   strSendAddr, arrAccount(1), arrAccountName(1), arrPasswd(1),textname,filename'����ɸ����˺ŷ���,uiMyEmailCnt
        WSH.Echo "��ǰ�˻� :" & arrAccount(1)
        WSH.Echo "�ѷ����� :" & strSendAddr
        wscript.sleep 2*60*1000  '��λms 2����  
        uiMyEmailCnt = uiMyEmailCnt + 1
        If uiMyEmailCnt = 2 Then '���uiMyEmailCnt������¼�˺Ÿ�����Ҳ����������Ԫ�ظ���
            uiMyEmailCnt = 0
        End If
        strSendAddr = ""
        uiCntAddr = 0 
    End If
End Function

Function SendEmailALL(Book,textname,filename)
    For uiSheetCnt = 1 To 2 'ע���޸������ֵ���ӵ�1�ű���22�ű�
        Set Sheet = Book.Sheets(uiSheetCnt)     
        SendEmailToOneSheetAddr Sheet,uiSheetCnt,textname,filename
    Next
End Function

''''''''''''''''''''''''''''''''''''''''''�����ʼ�'''''''''''''''''''''''''''''''''''''''''''''''
