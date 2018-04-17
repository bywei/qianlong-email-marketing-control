<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>钱龙邮件群发系统后台服务管理</title>
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
<script type="text/javascript">
function setContainerSize(){
document.getElementById("mainFrameTd").style.height=(document.documentElement.clientHeight-120)+"px";
document.getElementById("leftFrameTd").style.height=(document.documentElement.clientHeight-150)+"px";
document.getElementById("mainFrame").style.height=(document.documentElement.clientHeight-120)+"px";
}
window.onload=setContainerSize;
window.onresize=setContainerSize;
</script> 
</head>
<body>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="74" colspan="2" background="images/index1_03.gif">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="33%" rowspan="2"><img src="images/index1_01.gif" width="253" height="74" /></td>
          <td width="6%" rowspan="2">&nbsp;</td>
          <td width="61%" height="38" align="right">
            <table width="120" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="center"><img src="images/index1_06.gif" width="16" height="16" /></td>
                <td align="center" class="font2"><a href="#" class="font2"><strong><%=DateDiff("d",date,dateadd("m",1,date))%>帮助</strong></a></td>
                <td align="center"><img src="images/index1_08.gif" width="16" height="16" /></td>
                <td align="center" class="font2"><a href="login.asp?action=loginout" class="font2"><strong>退出</strong></a></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td align="right">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="right" class="font2"><a href="#"><img src="images/index1_13.gif" width="84" height="24" border="0" align="absmiddle" /></a>&nbsp;|&nbsp;登陆用户：<%=session("admin")%>&nbsp;|&nbsp;身份：管理员&nbsp;|&nbsp;界面风格：<a href="#"><img src="images/index1_16.gif" width="48" height="24" border="0" align="absmiddle" /></a>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="2">
      <table width="100%" border="0" cellspacing="10" cellpadding="0">
        <tr>
          <td width="10%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="8" height="8"><img src="images/index1_23.gif" width="8" height="29" /></td>
                <td width="99%" background="images/index1_24.gif"></td>
                <td width="8" height="8"><img src="images/index1_26.gif" width="8" height="29" /></td>
              </tr>
              <tr>
                <td background="images/index1_45.gif"></td>
                <td id="leftFrameTd" bgcolor="#FFFFFF" style="padding:0 2px 0 2px; vertical-align:top;height:500px;">
                  <table width="200" border="0" cellpadding="0" cellspacing="5">
                    <tr>
                      <td width="16%" align="center"><img src="images/index1_54.gif" width="15" height="11" /></td>
                      <td width="84%">用户管理</td>
                    </tr>
                    <tr>
                      <td align="center">&nbsp;</td>
                      <td>
                        <table width="100%" border="0" cellspacing="5" cellpadding="0">
                          <tr>
                            <td width="18%" align="center"><img src="images/index1_68.gif" width="11" height="14" /></td>
                            <td width="82%"><a href="userlist.asp" target="mainFrame">用户列表</a></td>
                          </tr>
                          <tr>
                            <td align="center"><img src="images/index1_68.gif" width="11" height="14" /></td>
                            <td><a href="loglist.asp" target="mainFrame">操作日志</a></td>
                          </tr>
                          <tr>
                            <td align="center"><img src="images/index1_68.gif" width="11" height="14" /></td>
                            <td><a href="adminlist.asp" target="mainFrame">管理设置</a></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td align="center"><img src="images/index1_54.gif" width="15" height="11" /></td>
                      <td><a href="editconfig.asp" target="mainFrame">配置管理</a></td>
                    </tr>
                  </table>
                </td>
                <td background="images/index1_47.gif"></td>
              </tr>
              <tr>
                <td width="8" height="8"><img src="images/index1_91.gif" width="8" height="8" /></td>
                <td background="images/index1_92.gif"></td>
                <td width="8" height="8"><img src="images/index1_93.gif" width="8" height="8" /></td>
              </tr>
            </table>
          </td>
       
          <td id="mainFrameTd" width="100%"  valign="top">
             <iframe src="center.html" name="mainFrame" id="mainFrame" width="100%" height="100%" frameborder="0" />
            
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
