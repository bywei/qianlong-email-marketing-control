<%@LANGUAGE="VBSCRIPT"%>
<% Ip_address=Request.ServerVariables ("HTTP_X_FORWARDED_FOR")
    If Ip_address=""Then  
       Ip_address= Request.ServerVariables ("REMOTE_ADDR")
    end if
	response.Write(Ip_address)
%>