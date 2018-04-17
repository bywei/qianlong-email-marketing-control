<!--#include FILE="Conn.asp"-->
<%
 Uid =trim(request.QueryString("u"))
 Url=trim(request.QueryString("l"))
 Gid=trim(request.QueryString("g"))
 Mail=trim(request.QueryString("m"))
 
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Url_count Where Uid='"&Uid&"' and Url='"&Url&"' and Mail='"&Mail&"'"
 Rs.Open Sql,Conn,3,2
 If Rs.Eof And Rs.Bof Then
   Rs.AddNew
   Rs("CreateDate")=Now()
   Rs("Uid")=Uid
   Rs("Gid")=Gid
   Rs("Url")=Url
   Rs("Mail")=Mail
 end if
 Rs("LastDate")=Now()
 Rs("UrlCount")=cint(Rs("UrlCount"))+1
 Rs.Update
 set Rs=nothing
 conn.close()
 
 response.Redirect(Url)
%>