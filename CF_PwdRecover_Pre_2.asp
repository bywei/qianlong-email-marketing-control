<%If Action="" Then%>
          <table class="tb_2">
                    <form name="form1" method="post" action="?Action=pwdget">
            <tr class="tr_1"> 
 
                        <td colspan="2" >�һ����ǵ�����</td>
                      </tr>

                      <tr> 
                        <td width="40%" class="td_3">����û����� 
                          </td>
                        <td ><input name="UserName" type="text" id="UserName" size="30"></td>
                      </tr>
                      <tr> 
                        <td  class="td_3">ȡ�ط�ʽ�� </td>
                        <td ><input name="assort" type="radio" value="1" checked>
                          ������ʾ���� 
                          <input type="radio" name="assort" value="2">
                          ���͵�Email</td>
                      </tr>
                      <tr> 
                        <td> </td>
						<td>
                            <input type="submit" name="Submit" value="ȡ������">
                          </td>
                      </tr>
                    </form>
                  </table> 
<%End if%>

<%If Action="pwdget" And CInt(Request("Assort"))=1 Then%>
<%
UserName=GoBack(ChkStr(request("UserName"),1),"�������û���")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CFCount_User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof and rs.bof then  Call AlertBack("û�д��û���!",1)
if Rs("PasswordAsk")="" Then Call AlertBack("����ǰû����д���뱣������!",1)
%>
<table class="tb_2">
<form name="form1" method="post" action="?Action=answermodifypwd&UserName=<%=UserName%>">

            <tr class="tr_1"> 
                        <td colspan="2" >��������Ĵ�</td>
                      </tr>
                      <tr> 
                        <td  width="40%" class="td_3"> ������ʾ���⣺ 
                          </td>
                        <td ><%=rs("passwordask")%></td>
                      </tr>
                      <tr> 
                        <td class="td_3">������ʾ�𰸣� </td>
                        <td ><input name="passwordanswer" type="text" size="20"></td>
                      </tr>
                      <tr> 
                        <td > </td>
						<td>
                            <input type="submit" name="Submit3" value="ȷ��">
                          </td>
                      </tr>
                    </form>
                  </table>
<%End if%>


<%If Action="pwdget" And CInt(Request("Assort"))=2 Then%>  
<%
UserName=GoBack(ChkStr(request("UserName"),1),"�������û���")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CFCount_User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof and rs.bof then  Call AlertBack("û�д��û���!",1)

EmailTo=Rs("email")
EmailFrom=RsSet("SmtpUser")
EmailFromName=RsSet("SysTitle")
EmailSubj="�һ����������"
EmailBody=RsSet("SysTitle")&"���͵��޸ļ����������������,����ʱ��"&Now()&"<br>�����������޸���ļ���������<br>"
EmailBody=EmailBody & HttpPath(2)&"PwdRecover.asp?Action=pwdmodify&UserCode="&Rs("UserCode")


If RsSet("emailsendtype")=1 Then
 Set objMail = Server.CreateObject("CDONTS.Newmail")
 objMail.To = EmailTo
 objMail.From = EmailFrom
 objMail.Subject = EmailSubj
 objMail.Body = EmailBody
 objMail.Send
 Set objMail = Nothing
ElseIf RsSet("emailsendtype")=2 Then
 Set msg = Server.CreateObject("JMail.Message")
 msg.silent = true
 msg.Logging = true
 msg.Charset = "gb2312"
 msg.ContentType = "text/HTML"
 msg.MailServerUserName = RsSet("SmtpUser")
 msg.MailServerPassword = RsSet("SmtpPassword")
 msg.From = EmailFrom
 msg.FromName = EmailFromName
 msg.AddRecipient EmailTo
 msg.Subject = EmailSubj
 msg.Body = EmailBody
 
 If msg.Send (RsSet("SmtpAddress"))=false then
  set msg = nothing
  Call AlertUrl("�ʼ�����ʧ��","?")
 End If 
 
 set msg = nothing

End if
%>

<table class="tb_2">
              <tr class="tr_1"> 
                <td class="td_2">�����Ѿ��ɹ����͵����<font color="#33CC00"><font color="#ff0000"><%=rs("email")%></font></font>,�����,���û���յ��ʼ�����ϵ������ϵͳ����Ա�����</td></tr> 
              <tr> 
                <td class="td_2"><input type="submit" name="Submit" value="������ҳ" onclick="location.href='Index.asp'"></tr> 
</table>
<%end if%>
  
<%If Action="pwdmodify" Then%>
<%UserCode=ChkStr(Request("UserCode"),1)%>
<table class="tb_2">
                    <form name="form1" method="post" action="?Action=pwdmodifysave&UserCode=<%=UserCode%>">
<tr class="tr_1"> 
                        <td colspan="2">�޸�����</td>
                      </tr>
                      <tr> 
                        <td   width="40%" class="td_3">�����룺</td>
                        <td > <input name="Pwd" type="password" id="password"> 
                        </td>
                      </tr>
                      <tr> 
                        <td  class="td_3">������ȷ�ϣ�</td>
                        <td ><input name="Pwd2" type="password" id="password2"></td>
                      </tr>
                      <tr> 
                        <td >&nbsp;</td>
                        <td ><input type="submit" name="Submit2" value="�޸�����"> 
                        </td>
                      </tr>
                    </form>
                  </table>
<%End If%>