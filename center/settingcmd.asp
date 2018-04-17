<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<%
dim id,isLogin,isEmailVer,isReg,isSend
  id=trim(request.Form("id"))
  isLogin=trim(request.Form("isLogin"))
  isEmailVer=trim(request.Form("isEmailVer"))
  isReg=trim(request.Form("isReg"))
  isSend=trim(request.Form("isSend"))
  
  
     if id="" then 
	    response.Redirect("editconfig.asp?msg=修改失败id为空")
	  end if
	'修改数据
     upsql="select * from setting where Id="&id
	 set uprs=server.CreateObject("ADODB.recordset")
	 uprs.open upsql,conn,1,3
	    uprs("isLogin")=isLogin
		uprs("isEmailVer")=isEmailVer
        uprs("isReg")=isReg
        uprs("isSend")=isSend
     uprs.update
     uprs.close
     set uprs=nothing
     conn.close
     set conn=nothing
     response.Redirect("editconfig.asp?msg=修改成功")

%>