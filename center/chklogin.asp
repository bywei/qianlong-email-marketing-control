<%
 if session("admin") ="" then
   response.Redirect("index.html")
 end if
%>