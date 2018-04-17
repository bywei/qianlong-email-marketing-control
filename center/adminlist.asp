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
-->
</style>
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
                <td align="center" background="images/index1_36.gif"><a href="#" class="font3"><strong>管理员</strong></a></td>
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
    <td bgcolor="#FFFFFF" style="height:490px; vertical-align:top;"><table width="100%" border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#C4E7FB">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
                    <tr>
                      <td>&nbsp;<a href="#" class="font1">首页</a> &gt; <a href="#" class="font1">管理设置</a> &gt; <a href="#" class="font1"><strong>列表</strong></a></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#BBD3EB">
              <tr>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">&nbsp;</td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><label>
                    <input type="checkbox" name="checkbox4" value="checkbox" />
                  </label></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">用户名</td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">密码</td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">操作</td>
              </tr>
     <%
     dim chksql
        chksql="select * from admin where adminName='"&session("admin")&"'"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	 %>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"><img src="images/index1_75.gif" width="16" height="15" /></td>
                <td height="26" align="center" bgcolor="#FFFFFF">
                    <input type="checkbox" name="userId" value="<%=chkrs("id")%>" />
                </td>
                <td height="26" align="left" bgcolor="#FFFFFF"><a href="editadmin.asp?id=<%=chkrs("id")%>" target="_self"><%=chkrs("adminName")%></a></td>
                <td height="26" align="left" bgcolor="#FFFFFF"><input type="password" value="<%=chkrs("adminPwd")%>" readonly="true"/></td>
                <td height="26" align="left" bgcolor="#FFFFFF"><a href="editadmin.asp?id=<%=chkrs("id")%>" target="_self">编辑用户</a></td>
              </tr>
      <% chkrs.movenext
         loop
	     chkrs.close()
	     conn.close
         set conn=nothing
      %>
            </table></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#E1E5E8">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="5" bgcolor="#FDFDFF">
                    <tr>
                      <td>&nbsp;<img src="images/index1_78.gif" width="16" height="12" align="absmiddle" />&nbsp;已读信息&nbsp;&nbsp;<img src="images/index1_75.gif" width="16" height="15" align="absmiddle" />&nbsp;未读信息</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td><a href="#"><img src="images/index1_86.gif" width="74" height="31" border="0" /></a>&nbsp;<a href="#"><img src="images/index1_82.gif" width="74" height="31" border="0" /></a>&nbsp;<a href="#"><img src="images/index1_84.gif" width="74" height="31" border="0" /></a></td>
        </tr>
      </table>
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
