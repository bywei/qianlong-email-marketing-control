<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<%

'ͳ����Ϣ
 dim action,userId,url
   action=trim(request.QueryString("action"))
   userId=trim(request.QueryString("userId"))
   url=GetUrl()
  
   dim sql,isPresence
   isPresence="false"
   
   if action="i" then
     chksql="select * from tongji"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("url")=url then
		   isPresence="true"
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
   
   '���ݿ��в����ھͲ�������
   if isPresence="false" then
      sql="select * from tongji"
      set rs=server.CreateObject("adodb.recordset")
	  rs.open sql,conn,1,3
	  rs.addnew
	  rs("userId")=userId
	  rs("url")=url
	  rs.update
	  rs.close
	  set rs=nothing
	  conn.close
      set conn=nothing
	  response.Write("http://rescdn.qqmail.com/zh_CN/logo//201110230_lanjingling/logo_0_201110230.gif")
	  response.End()	
	end if
	'������ݿ��д��ھ͸�������
	  response.Write("http://rescdn.qqmail.com/zh_CN/dy/maillogo/mail_20069.gif")
	  response.End()	
  end if
  
  '�õ���ǰҳ��ĵ�ַ
Function GetUrl()
On Error Resume Next
Dim strTemp
If LCase(Request.ServerVariables("HTTPS")) = "off" Then
strTemp = "http://"
Else
strTemp = "https://"
End If
strTemp = strTemp & Request.ServerVariables("SERVER_NAME")
If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT")
strTemp = strTemp & Request.ServerVariables("URL")
If Trim(Request.QueryString) <> "" Then strTemp = strTemp & "?" & Trim(Request.QueryString)
GetUrl = strTemp
End Function 
  
%>
