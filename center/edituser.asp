<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>管理页面</title>
<!-- 调用CSS，JS -->
<link href="images/style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	font-family: "宋体";
	font-size: 12px;
	color: #333333;
	background-color: #2286C2;
}
#button{width:74px; height:31px; margin:0; border: 0 none;}
.subButton{background-image:url(images/index1_82.gif);}
.exitButton{ background-image:url(images/index1_82.gif);}
-->
</style>
<%
  dim userId,msg
  userId=request.QueryString("userId")
  msg=request.QueryString("msg")
  if userId="" then 
      response.Redirect("userlist.asp")
      response.End()
  end if
%>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="8" height="8"><img src="images/index1_28.gif" width="8" height="39" /></td>
    <td width="99%" background="images/index1_40.gif"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="90" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_35.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_36.gif"><a href="#" class="font3"><strong>用户列表</strong></a></td>
                <td width="5"><img src="images/index1_38.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>收件箱</strong></a></td>
                <td width="5"><img src="images/index1_33.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>发件箱</strong></a></td>
                <td width="5"><img src="images/index1_33.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>草稿箱</strong></a></td>
                <td width="5"><img src="images/index1_33.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="8" height="8"><img src="images/index1_43.gif" width="8" height="39" /></td>
  </tr>
  <tr>
    <td background="images/index1_45.gif"></td>
    <td bgcolor="#FFFFFF" style="height:490px; vertical-align:top;">
    <form action="usercmd.asp" method="post">
    <table width="100%" border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td>
          <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#C4E7FB">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
                    <tr>
                      <td>&nbsp;<a href="#" class="font1">首页</a> &gt; <a href="#" class="font1">用户管理</a> &gt;用户列表&gt; <a href="#" class="font1"><strong>编辑用户</strong></a></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td>
         
            <%
     dim chksql
        chksql="select * from Users where id="&userId
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	 %>
          <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#BBD3EB">
              <tr>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"></td>
                <td height="27" colspan="2" align="left" background="images/index1_72.gif" bgcolor="#FFFFFF">编辑用户信息  <span style="color:red; text-align:center"><%=msg%></span></td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">用户名</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                <input type="hidden" name="action" value="update"/>
                <input type="hidden" name="userId" id="userId" value="<%=chkrs("id")%>"/>
                <input type="text" name="UserName" id="UserName" value="<%=chkrs("UserName")%>" />
                <a href="loglist.asp?Px=CreateTime&userName=<%=chkrs("UserName")%>">查看日志</a>
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">密码</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="password" name="UserPwd" id="UserPwd" value="<%=chkrs("UserPwd")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">邮箱</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserEmail" id="UserEmail" value="<%=chkrs("UserEmail")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">积分</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserScore" id="UserScore" value="<%=chkrs("UserScore")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">级别</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserLevel" id="UserLevel"  value="<%=chkrs("UserLevel")%>"/>
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">发送量</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserSendNum" id="UserSendNum" value="<%=chkrs("UserSendNum")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">VIP</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserIsVip" id="UserIsVip" value="<%=chkrs("UserIsVip")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">状态</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserState" id="UserState" value="<%=chkrs("UserState")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">登录状态</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="LoginState" id="UserState" value="<%=chkrs("LoginState")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">类型</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UserType" id="UserType" value="<%=chkrs("UserType")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">创建时间</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="CreateTime" id="CreateTime" value="<%=chkrs("CreateTime")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">使用时间</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UseOfTime" id="UseOfTime" value="<%=chkrs("UseOfTime")%>" />
                </td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"></td>
                <td height="26" align="center" bgcolor="#FFFFFF">最后登录时间</td>
                <td height="26" align="left" bgcolor="#FFFFFF">
                  <input type="text" name="UseOfTime" id="UseOfTime" value="<%=chkrs("LoginTime")%>" readonly="true" />
                </td>
              </tr>
            </table>
               <% chkrs.movenext
         loop
	     chkrs.close()
	     conn.close
         set conn=nothing
      %>
         
          </td>
        </tr>
        <tr>
          <td><a href="#">
            <input type="submit" name="button" id="button" value="" class="subButton" />
          </a></td>
        </tr>
      </table>
       </form>
      <h3>&nbsp;</h3></td>
    <td background="images/index1_47.gif"></td>
  </tr>
  <tr>
    <td width="8" height="8"><img src="images/index1_91.gif" width="8" height="8" /></td>
    <td background="images/index1_92.gif"></td>
    <td width="8" height="8"><img src="images/index1_93.gif" width="8" height="8" /></td>
  </tr>
</table>
</body>
</html>
