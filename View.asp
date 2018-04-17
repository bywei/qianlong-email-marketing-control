
<!--#include file="Conn.asp"-->
<!--#include file="CF_MyFunction.asp"-->
<!--#include file="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="CF_Style.asp"-->
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<!--#Include File="CF_View_Pre.asp"-->

<!--#include file="Bottom.asp"-->
</BODY></HTML>

<%Call ConnClose()%>