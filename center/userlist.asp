<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ҳ��</title>
<!-- ����CSS��JS -->
<link href="images/style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	font-family: "����";
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
                <td align="center" background="images/index1_36.gif"><a href="#" class="font3"><strong>�û��б�</strong></a></td>
                <td width="5"><img src="images/index1_38.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>�ռ���</strong></a></td>
                <td width="5"><img src="images/index1_33.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>������</strong></a></td>
                <td width="5"><img src="images/index1_33.gif" width="5" height="39" /></td>
              </tr>
            </table></td>
          <td><table width="80" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="5"><img src="images/index1_29.gif" width="5" height="39" /></td>
                <td align="center" background="images/index1_30.gif"><a href="#" class="font2"><strong>�ݸ���</strong></a></td>
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
                      <td>&nbsp;<a href="#" class="font1">��ҳ</a> &gt; <a href="#" class="font1">�û�����</a> &gt; <a href="userlist.asp" class="font1"><strong>�û��б�</strong></a></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
                       <%
Px=Request("Px")
UserName=Request("UserName")
If Px="" Then  Px="CreateTime"

Sql="Select * From Users "
If UserName<>"" then Sql="Select * From Users where userName='"&UserName&"'"
Sql=Sql&" Order By "&Px&" Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,conn,1,1

If Not rs.eof then
 IF not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
  page=1
 Else
  Page=Int(Abs(Request("page")))
 End if
 rs.pagesize=15
 totalrs=rs.RecordCount
 totalpage=rs.pageCount
 mypagesize=rs.pagesize
 rs.absolutepage=page
End If
%>
<script type="text/javascript">
  function search(){
    var userName = document.getElementById("userName").value;
    window.location.href='?Px=CreateTime&userName='+userName+'&Page=<%=page%>';
  }

</script>
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#BBD3EB">
              <tr>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">&nbsp;</td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><label>
                    <input type="checkbox" name="checkbox4" value="checkbox" />
                  </label></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">�û���<input type="text" id="userName" name="userName" size="10" /><input type="button" value="����" onClick="search();" /></td>
                <!--td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">����</td-->
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserEmail&Page=<%=page%>">����</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserScore&Page=<%=page%>">����</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserLevel&Page=<%=page%>">����</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserSendNum&Page=<%=page%>">������</a></td>
                <td align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserIsVip&Page=<%=page%>">VIP</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UserState&Page=<%=page%>">״̬</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">����</td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=CreateTime&Page=<%=page%>">����ʱ��</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF"><a href="?Px=UseOfTime&Page=<%=page%>">ʹ������</a></td>
                <td height="27" align="center" background="images/index1_72.gif" bgcolor="#FFFFFF">����</td>
              </tr>
     <%
     'dim chksql
      '  chksql="select * from Users "
	 'set chkrs=server.CreateObject("adodb.recordset")
	 'chkrs.open chksql,conn,1,3
	 
	 'do while not chkrs.eof
	 %>
    <%
     While Not Rs.Eof And MyPageSize>0
    %>
              <tr>
                <td height="26" align="center" bgcolor="#FFFFFF"><img src="images/index1_75.gif" width="16" height="15" /></td>
                <td height="26" align="center" bgcolor="#FFFFFF">
                    <input type="checkbox" name="userId" value="<%=Rs("id")%>" />
                </td>
                <td height="26" align="left" bgcolor="#FFFFFF"><a href="edituser.asp?userId=<%=Rs("id")%>" target="_self"><%=Rs("UserName")%></a></td>
                <!--td height="26" align="left" bgcolor="#FFFFFF"><%=Rs("UserPwd")%></td-->
                <td height="26" align="left" bgcolor="#FFFFFF"><%=Rs("UserEmail")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=Rs("UserScore")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=Rs("UserLevel")%></td>
                <td height="26" align="center" bgcolor="#FFFFFF"><%=Rs("UserSendNum")%></td>
                <td height="26" align="center" bgcolor="#FFFFFF"><%=Rs("UserIsVip")%></td>
                <td height="26" align="center" bgcolor="#FFFFFF"><%=Rs("UserState")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=Rs("UserType")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=Rs("CreateTime")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=Rs("UseOfTime")%></td>
                <td align="center" bgcolor="#FFFFFF"><a href="edituser.asp?userId=<%=Rs("id")%>" target="_self">�޸�</a>
                <a href="usercmd.asp?action=delUser&userId=<%=Rs("id")%>" target="_self">ɾ��</a>
                <a href="loglist.asp?Px=CreateTime&userName=<%=Rs("UserName")%>">��־</a></td>
              </tr>
      <% 'chkrs.movenext
         'loop
	     'chkrs.close()
	     'conn.close
         'set conn=nothing
      %>

          <%
		   rs.movenext
           mypagesize=mypagesize-1
           wend
		   %>

      <tr >

</tr>
            </table></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#E1E5E8">
              <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="5" bgcolor="#FDFDFF">
                    <tr>
                      <td><a href="?Px=<%=Px%>&Page=1">��һҳ</a>
            <%
if page>1 then%>
            <a href='?Px=<%=Px%>&Page=<%=Page-1%>'>��һҳ</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Px=<%=Px%>&Page=<%=Page+1%>'>��һҳ</a>
            <%

end if
%>
            <a href="?Px=<%=Px%>&Page=<%=totalpage%>">���һҳ</a> ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%>
<img src="images/index1_78.gif" width="16" height="12" align="absmiddle" />�Ѷ���Ϣ &nbsp;&nbsp;<img src="images/index1_75.gif" width="16" height="15" align="absmiddle" />&nbsp;δ����Ϣ   
                      </td>
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
