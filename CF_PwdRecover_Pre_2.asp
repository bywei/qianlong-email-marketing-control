<%If Action="" Then%>
          <table class="tb_2">
                    <form name="form1" method="post" action="?Action=pwdget">
            <tr class="tr_1"> 
 
                        <td colspan="2" >找回忘记的密码</td>
                      </tr>

                      <tr> 
                        <td width="40%" class="td_3">你的用户名： 
                          </td>
                        <td ><input name="UserName" type="text" id="UserName" size="30"></td>
                      </tr>
                      <tr> 
                        <td  class="td_3">取回方式： </td>
                        <td ><input name="assort" type="radio" value="1" checked>
                          根据提示问题 
                          <input type="radio" name="assort" value="2">
                          发送到Email</td>
                      </tr>
                      <tr> 
                        <td> </td>
						<td>
                            <input type="submit" name="Submit" value="取回密码">
                          </td>
                      </tr>
                    </form>
                  </table> 
<%End if%>

<%If Action="pwdget" And CInt(Request("Assort"))=1 Then%>
<%
UserName=GoBack(ChkStr(request("UserName"),1),"请输入用户名")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CFCount_User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof and rs.bof then  Call AlertBack("没有此用户名!",1)
if Rs("PasswordAsk")="" Then Call AlertBack("你以前没有填写密码保护资料!",1)
%>
<table class="tb_2">
<form name="form1" method="post" action="?Action=answermodifypwd&UserName=<%=UserName%>">

            <tr class="tr_1"> 
                        <td colspan="2" >请输入你的答案</td>
                      </tr>
                      <tr> 
                        <td  width="40%" class="td_3"> 密码提示问题： 
                          </td>
                        <td ><%=rs("passwordask")%></td>
                      </tr>
                      <tr> 
                        <td class="td_3">密码提示答案： </td>
                        <td ><input name="passwordanswer" type="text" size="20"></td>
                      </tr>
                      <tr> 
                        <td > </td>
						<td>
                            <input type="submit" name="Submit3" value="确定">
                          </td>
                      </tr>
                    </form>
                  </table>
<%End if%>


<%If Action="pwdget" And CInt(Request("Assort"))=2 Then%>  
<%
UserName=GoBack(ChkStr(request("UserName"),1),"请输入用户名")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CFCount_User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof and rs.bof then  Call AlertBack("没有此用户名!",1)

EmailTo=Rs("email")
EmailFrom=RsSet("SmtpUser")
EmailFromName=RsSet("SysTitle")
EmailSubj="找回密码的链接"
EmailBody=RsSet("SysTitle")&"发送的修改计数器的密码的链接,发送时间"&Now()&"<br>打开下面链接修改你的计数器密码<br>"
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
  Call AlertUrl("邮件发送失败","?")
 End If 
 
 set msg = nothing

End if
%>

<table class="tb_2">
              <tr class="tr_1"> 
                <td class="td_2">密码已经成功发送到你的<font color="#33CC00"><font color="#ff0000"><%=rs("email")%></font></font>,请查收,如果没有收到邮件请联系计数器系统管理员解决！</td></tr> 
              <tr> 
                <td class="td_2"><input type="submit" name="Submit" value="返回首页" onclick="location.href='Index.asp'"></tr> 
</table>
<%end if%>
  
<%If Action="pwdmodify" Then%>
<%UserCode=ChkStr(Request("UserCode"),1)%>
<table class="tb_2">
                    <form name="form1" method="post" action="?Action=pwdmodifysave&UserCode=<%=UserCode%>">
<tr class="tr_1"> 
                        <td colspan="2">修改密码</td>
                      </tr>
                      <tr> 
                        <td   width="40%" class="td_3">新密码：</td>
                        <td > <input name="Pwd" type="password" id="password"> 
                        </td>
                      </tr>
                      <tr> 
                        <td  class="td_3">新密码确认：</td>
                        <td ><input name="Pwd2" type="password" id="password2"></td>
                      </tr>
                      <tr> 
                        <td >&nbsp;</td>
                        <td ><input type="submit" name="Submit2" value="修改密码"> 
                        </td>
                      </tr>
                    </form>
                  </table>
<%End If%>