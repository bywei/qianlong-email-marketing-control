<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<%
dim id,adminName,adminPwd
  id=trim(request.Form("id"))
  adminName=trim(request.Form("adminName"))
  adminPwd=trim(request.Form("adminPwd"))
  
     if id="" then 
	    response.Redirect("editadmin.asp?msg=�޸�ʧ��idΪ��")
	  end if
	'�޸�����
     upsql="select * from admin where Id="&id
	 set uprs=server.CreateObject("ADODB.recordset")
	 uprs.open upsql,conn,1,3
	    uprs("adminName")=adminName
		uprs("adminPwd")=adminPwd
     uprs.update
     uprs.close
     set uprs=nothing
     conn.close
     set conn=nothing
     response.Redirect("editadmin.asp?msg=�޸ĳɹ�")

%>