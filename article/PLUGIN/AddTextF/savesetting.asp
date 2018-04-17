<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 Walle 其它版本的Z-blog未知
'// 插件制作:    kikin(http://www.2sos.net/)
'// 备    注:    在文章前增加固定内容，可用于挂广告和版权说明 - 挂口页
'// 最后修改：   2011-10-23
'// 最后版本:    0.5
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_function_md5.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="../../function/c_system_event.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6) 

	Dim strContent
	strContent=LoadFromFile(BlogPath & "PLUGIN/AddTextF/mysetting.asp","utf-8")
    
    Dim objRegExp,ReplaceStr
    Set objRegExp = new RegExp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.multiline = True 
  
   Dim TagInf
   Dim MyWordA,PosA,HPosA
   Dim MyWordB,PosB,HPosB
   Dim MyWordC,PosC,HPosC
   Dim MyWordD,PosD,HPosD
   Dim MyLeft,MyLeftEnd,MyRight,MyRightEnd,LocalNum,LocalPas
   Dim MySLeft,MySLeftEnd,MySRight,MySRightEnd
   
   ReplaceStr ="< br /  >"
   objRegExp.Pattern= "\r\n"
   MyWordA=Request.Form("MyWordA")
   MyWordA=objRegExp.Replace(MyWordA, ReplaceStr)	
   PosA=Request.Form( "PosA") 
   HPosA=Request.Form( "HPosA")
   MyWordB=Request.Form("MyWordB")
   MyWordB=objRegExp.Replace(MyWordB, ReplaceStr)	
   PosB=Request.Form( "PosB") 
   HPosB=Request.Form( "HPosB")
   MyWordC=Request.Form("MyWordC")
   MyWordC=objRegExp.Replace(MyWordC, ReplaceStr)	
   PosC=Request.Form( "PosC") 
   HPosC=Request.Form( "HPosC")
   MyWordD=Request.Form("MyWordD")
   MyWordD=objRegExp.Replace(MyWordD, ReplaceStr)	
   PosD=Request.Form( "PosD") 
   HPosD=Request.Form( "HPosD")

   TagInf=Request.Form( "TagInf")
   
   MyLeft=Request.Form( "MyLeft")
   MyLeftEnd=Request.Form( "MyLeftEnd")
   MyRight=Request.Form( "MyRight")
   MyRightEnd=Request.Form( "MyRightEnd")
   MySLeft=Request.Form( "MySLeft")
   MySLeftEnd=Request.Form( "MySLeftEnd")
   MySRight=Request.Form( "MySRight")
   MySRightEnd=Request.Form( "MySRightEnd")
   LocalNum=Request.Form( "LocalNum")
   LocalPast=Request.Form( "LocalPast")

   Call SaveValueForSetting(strContent,True,"String","sMyWordA",MyWordA)
   Call SaveValueForSetting(strContent,True,"String","sPosA",PosA)   
   Call SaveValueForSetting(strContent,True,"String","sHPosA",HPosA)

   Call SaveValueForSetting(strContent,True,"String","sMyWordB",MyWordB)
   Call SaveValueForSetting(strContent,True,"String","sPosB",PosB)   
   Call SaveValueForSetting(strContent,True,"String","sHPosB",HPosB)

   Call SaveValueForSetting(strContent,True,"String","sMyWordC",MyWordC)
   Call SaveValueForSetting(strContent,True,"String","sPosC",PosC)   
   Call SaveValueForSetting(strContent,True,"String","sHPosC",HPosC)

   Call SaveValueForSetting(strContent,True,"String","sMyWordD",MyWordD)
   Call SaveValueForSetting(strContent,True,"String","sPosD",PosD)   
   Call SaveValueForSetting(strContent,True,"String","sHPosD",HPosD)

   Call SaveValueForSetting(strContent,True,"String","sTagInf",TagInf)  
   Call SaveValueForSetting(strContent,True,"String","sMyLeft",MyLeft)    
   Call SaveValueForSetting(strContent,True,"String","sMyRight",MyRight)    
   Call SaveValueForSetting(strContent,True,"String","sMyLeftEnd",MyLeftEnd)    
   Call SaveValueForSetting(strContent,True,"String","sMyRightEnd",MyRightEnd)    
    Call SaveValueForSetting(strContent,True,"String","sMySLeft",MySLeft)    
   Call SaveValueForSetting(strContent,True,"String","sMySRight",MySRight)    
   Call SaveValueForSetting(strContent,True,"String","sMySLeftEnd",MySLeftEnd)    
   Call SaveValueForSetting(strContent,True,"String","sMySRightEnd",MySRightEnd)    
   Call SaveValueForSetting(strContent,True,"String","sLocalNum",LocalNum)    
   Call SaveValueForSetting(strContent,True,"String","sLocalPast",LocalPast)    


   Call SaveToFile(BlogPath & "/PLUGIN/AddTextF/mysetting.asp",strContent,"utf-8",False)



If Err.Number=0 then
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "setting.asp"
End If


Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
