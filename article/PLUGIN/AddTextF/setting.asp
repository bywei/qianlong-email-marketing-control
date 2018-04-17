<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 Walle 其它版本的Z-blog未知
'// 插件制作:    kikin(http://www.2sos.net/)
'// 备    注:    在文章前增加固定内容，可用于挂广告和版权说明 - 挂口页
'// 最后修改：   2011-10-23
'// 最后版本:    0.4
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
If CheckPluginState("AddTextF")=False Then Call ShowError(48)

BlogTitle="AddTextF"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../script/common.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>

<div id="divMain">
<div class="Header">AddTextF 插件设定</div>
<div id="divMain2">
<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<br />
<form id="edit" name="edit" method="post">
<%
	Dim tmpSng
	tmpSng=LoadFromFile(BlogPath & "/PLUGIN/AddTextF/mysetting.asp","utf-8")
    
    Dim objRegExp,ReplaceStr
    Set objRegExp = new RegExp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.multiline = True 
    Response.Write "<p>※<a href=""http://www.2sos.net/post/AddTextF.html"">AddTextF日志加加</a>:欲显示的文字, 可使用HTML代码控制样式。提交后记得""文件重建""完成更新!</p>"
   Dim TagInf,tmpstrTS 
	Dim MyWordA,PosA,HPosA,tmpstrA,tmpstrHA
   Dim MyWordB,PosB,HPosB,tmpstrB,tmpstrHB
   Dim MyWordC,PosC,HPosC,tmpstrC,tmpstrHC
   Dim MyWordD,PosD,HPosD,tmpstrD,tmpstrHD
   Response.Write "<table  border=""1""> <tr> <td  colspan=""2"">"      
   Response.Write "<p>关键字Tag设定：<input type=""radio"" checked=""checked"" name=""TagInf"" id=""TagInf0"" value=""0"" />首次出现"   
   Response.Write "<input type=""radio""  name=""TagInf"" id=""TagInf1"" value=""1"" />全部"
   Response.Write "<input type=""radio""  name=""TagInf"" id=""TagInf2"" value=""2"" />每两次 "
   Response.Write "<input type=""radio""  name=""TagInf"" id=""TagInf3"" value=""3"" />每三次"
   Response.Write "<input type=""radio""  name=""TagInf"" id=""TagInf4"" value=""4"" />每四次 "   
   Response.Write "<input type=""radio""  name=""TagInf"" id=""TagInf6"" value=""6"" />不链接 </p></td></tr>"   
   If LoadValueForSetting(tmpSng,True,"String","sTagInf",TagInf) Then 
      tmpstrTS="TagInf"&TagInf      
   End If   
   Response.Write "<tr> <td width=""50%"">"       
   If LoadValueForSetting(tmpSng,True,"String","sMyWordA",MyWordA) Then 
     ReplaceStr = VBCrlf 
     objRegExp.Pattern= "< br /  >"
     MyWordA=objRegExp.Replace(MyWordA, ReplaceStr)
     MyWordA = TransferHTML(MyWordA,"[html-format]")
     Response.Write "插入文本一:<br/><textarea rows=""7"" cols=""60"" name=""MyWordA"">" & MyWordA & "</textarea> " 
   End If
   Response.Write "<p>上下位置：<input type=""radio"" checked=""checked"" name=""PosA"" id=""PosAHead"" value=""Head"" />文头"
   Response.Write "<input type=""radio"" name=""PosA"" id=""PosABefore"" value=""Before"" />段前<input type=""radio"" name=""PosA"" id=""PosAAfter"" value=""After"" />段间"
   Response.Write "<input type=""radio"" name=""PosA"" id=""PosAFoot"" value=""Foot"" />文尾 <input type=""radio"" name=""PosA"" id=""PosANone"" value=""None"" />停用<br />"
   If LoadValueForSetting(tmpSng,True,"String","sPosA",PosA) Then 
      tmpstrA="PosA"&PosA      
   End If
   Response.Write"左右位置：<input type=""radio"" checked=""checked"" name=""HPosA"" id =""HPosANone"" value=""None"" />正常"
   Response.Write "<input type=""radio"" name=""HPosA""  id =""HPosALeft"" value=""Left"" />左悬浮<input type=""radio"" name=""HPosA""  id =""HPosARight"" value=""Right"" />右悬浮</p>"
   If LoadValueForSetting(tmpSng,True,"String","sHPosA",HPosA) Then 
      tmpstrHA="HPosA"&HPosA      
   End If     
  
   Response.Write "</td> <td width=""50%"">"
      If LoadValueForSetting(tmpSng,True,"String","sMyWordB",MyWordB) Then 
     ReplaceStr = VBCrlf 
     objRegExp.Pattern= "< br /  >"
     MyWordB=objRegExp.Replace(MyWordB, ReplaceStr)
     MyWordB = TransferHTML(MyWordB,"[html-format]")
     Response.Write "插入文本二:<br/><textarea rows=""7"" cols=""60"" name=""MyWordB"">" & MyWordB & "</textarea> " 
   End If
   Response.Write "<p>上下位置：<input type=""radio"" checked=""checked"" name=""PosB"" id=""PosBHead"" value=""Head"" />文头"
   Response.Write "<input type=""radio"" name=""PosB"" id=""PosBBefore"" value=""Before"" />段前<input type=""radio"" name=""PosB"" id=""PosBAfter"" value=""After"" />段间"
   Response.Write "<input type=""radio"" name=""PosB"" id=""PosBFoot"" value=""Foot"" />文尾<input type=""radio"" name=""PosB"" id=""PosBNone"" value=""None"" />停用 <br />"
   If LoadValueForSetting(tmpSng,True,"String","sPosB",PosB) Then 
      tmpstrB="PosB"&PosB      
   End If
   Response.Write"左右位置：<input type=""radio"" checked=""checked"" name=""HPosB"" id =""HPosBNone"" value=""None"" />正常"
   Response.Write "<input type=""radio"" name=""HPosB""  id =""HPosBLeft"" value=""Left"" />左悬浮<input type=""radio"" name=""HPosB""  id =""HPosBRight"" value=""Right"" />右悬浮</p>"
   If LoadValueForSetting(tmpSng,True,"String","sHPosB",HPosB) Then 
      tmpstrHB="HPosB"&HPosB      
   End If    
   
    Response.Write "</td></tr><tr> <td width=""50%"">"
    If LoadValueForSetting(tmpSng,True,"String","sMyWordC",MyWordC) Then 
     ReplaceStr = VBCrlf 
     objRegExp.Pattern= "< br /  >"
     MyWordC=objRegExp.Replace(MyWordC, ReplaceStr)
     MyWordC = TransferHTML(MyWordC,"[html-format]")
     Response.Write "插入文本三:<br/><textarea rows=""7"" cols=""60"" name=""MyWordC"">" & MyWordC & "</textarea> " 
   End If
   Response.Write "<p>上下位置：<input type=""radio"" checked=""checked"" name=""PosC"" id=""PosCHead"" value=""Head"" />文头"
   Response.Write "<input type=""radio"" name=""PosC"" id=""PosCBefore"" value=""Before"" />段前<input type=""radio"" name=""PosC"" id=""PosCAfter"" value=""After"" />段间"
   Response.Write "<input type=""radio"" name=""PosC"" id=""PosCFoot"" value=""Foot"" />文尾<input type=""radio"" name=""PosC"" id=""PosCNone"" value=""None"" />停用 <br />"
   If LoadValueForSetting(tmpSng,True,"String","sPosC",PosC) Then 
      tmpstrC="PosC"&PosC      
   End If
   Response.Write"左右位置：<input type=""radio"" checked=""checked"" name=""HPosC"" id =""HPosCNone"" value=""None"" />正常"
   Response.Write "<input type=""radio"" name=""HPosC""  id =""HPosCLeft"" value=""Left"" />左悬浮<input type=""radio"" name=""HPosC""  id =""HPosCRight"" value=""Right"" />右悬浮</p>"
   If LoadValueForSetting(tmpSng,True,"String","sHPosC",HPosC) Then 
      tmpstrHC="HPosC"&HPosC      
   End If    
   
      Response.Write "</td> <td width=""50%"">"
      If LoadValueForSetting(tmpSng,True,"String","sMyWordD",MyWordD) Then 
     ReplaceStr = VBCrlf 
     objRegExp.Pattern= "< br /  >"
     MyWordD=objRegExp.Replace(MyWordD, ReplaceStr)
     MyWordD = TransferHTML(MyWordD,"[html-format]")
     Response.Write "插入文本四:<br/><textarea rows=""7"" cols=""60"" name=""MyWordD"">" & MyWordD & "</textarea> " 
   End If
   Response.Write "<p>上下位置：<input type=""radio"" checked=""checked"" name=""PosD"" id=""PosDHead"" value=""Head"" />文头"
   Response.Write "<input type=""radio"" name=""PosD"" id=""PosDBefore"" value=""Before"" />段前<input type=""radio"" name=""PosD"" id=""PosDAfter"" value=""After"" />段间"
   Response.Write "<input type=""radio"" name=""PosD"" id=""PosDFoot"" value=""Foot"" />文尾 <input type=""radio"" name=""PosD"" id=""PosDNone"" value=""None"" />停用<br />"
   If LoadValueForSetting(tmpSng,True,"String","sPosD",PosD) Then 
      tmpstrD="PosD"&PosD      
   End If
   Response.Write"左右位置：<input type=""radio"" checked=""checked"" name=""HPosD"" id =""HPosDNone"" value=""None"" />正常"
   Response.Write "<input type=""radio"" name=""HPosD""  id =""HPosDLeft"" value=""Left"" />左悬浮<input type=""radio"" name=""HPosD""  id =""HPosDRight"" value=""Right"" />右悬浮</p>"
   If LoadValueForSetting(tmpSng,True,"String","sHPosD",HPosD) Then 
      tmpstrHD="HPosD"&HPosD      
   End If    
   
   
   Response.Write "</td></tr>"
   '系统设定区域
   Dim MyLeft,MyLeftEnd,MyRight,MyRightEnd,LocalPast,LocalNum 
   Dim MySLeft,MySLeftEnd,MySRight,MySRightEnd
   Response.Write " <tr> <td  colspan=""2"">"         
   Response.Write "以下为系统设定,如无必要请勿调整"   
   If LoadValueForSetting(tmpSng,True,"String","sMyLeft",MyLeft) Then 
      MyLeft = TransferHTML(MyLeft,"[html-format]")
      Response.Write "<p>左悬浮前置代码<input name=""MyLeft"" style=""width:40%"" type=""text"" value=""" & MyLeft & """ />  "   
   End If
    If LoadValueForSetting(tmpSng,True,"String","sMyLeftEnd",MyLeftEnd) Then 
      MyLeftEnd = TransferHTML(MyLeftEnd,"[html-format]")
      Response.Write "左悬浮后置代码<input name=""MyLeftEnd"" style=""width:10%"" type=""text"" value=""" & MyLeftEnd & """ /> <br/>"    
   End If
    If LoadValueForSetting(tmpSng,True,"String","sMyRight",MyRight) Then 
      MyRight = TransferHTML(MyRight,"[html-format]")
      Response.Write "右悬浮前置代码<input name=""MyRight"" style=""width:40%"" type=""text"" value=""" & MyRight & """ />  "   
   End If
    If LoadValueForSetting(tmpSng,True,"String","sMyRightEnd",MyRightEnd) Then 
      MyRightEnd = TransferHTML(MyRightEnd,"[html-format]")
      Response.Write "右悬浮后置代码<input name=""MyRightEnd"" style=""width:10%"" type=""text"" value=""" & MyRightEnd & """ /> <br />"    
   End If
   
   If LoadValueForSetting(tmpSng,True,"String","sMySLeft",MySLeft) Then 
      MySLeft = TransferHTML(MySLeft,"[html-format]")
      Response.Write "文尾特定:右悬浮前置代码<input name=""MySLeft"" style=""width:40%"" type=""text"" value=""" & MySLeft & """ />  "   
   End If
    If LoadValueForSetting(tmpSng,True,"String","sMySLeftEnd",MySLeftEnd) Then 
      MySLeftEnd = TransferHTML(MySLeftEnd,"[html-format]")
      Response.Write "右悬浮后置代码<input name=""MySLeftEnd"" style=""width:10%"" type=""text"" value=""" & MySLeftEnd & """ /> <br />"    
   End If
       If LoadValueForSetting(tmpSng,True,"String","sMySRight",MySRight) Then 
      MySRight = TransferHTML(MySRight,"[html-format]")
      Response.Write "文尾特定:右悬浮前置代码<input name=""MySRight"" style=""width:40%"" type=""text"" value=""" & MySRight & """ />  "   
   End If
    If LoadValueForSetting(tmpSng,True,"String","sMySRightEnd",MySRightEnd) Then 
      MySRightEnd = TransferHTML(MySRightEnd,"[html-format]")
      Response.Write "右悬浮后置代码<input name=""MySRightEnd"" style=""width:10%"" type=""text"" value=""" & MySRightEnd & """ /> <br/>"    
   End If
   
    If LoadValueForSetting(tmpSng,True,"String","sLocalNum",LocalNum) Then       
      Response.Write "<p> 文件尾部调整参数<input name=""LocalNum"" style=""width:10"" type=""text"" value=""" & LocalNum & """ />  "   
   End If
    If LoadValueForSetting(tmpSng,True,"String","sLocalPast",LocalPast) Then       
      Response.Write "微调参数<input name=""LocalPast"" style=""width:10"" type=""text"" value=""" & LocalPast & """ /> </p>"    
   End If  
   
   Response.Write "</td></tr></table>"
   '调整页面Radio设定
   Response.Write "<script language=""javascript"">  document.getElementById('"&tmpstrA&"').checked=true; document.getElementById('"&tmpstrHA&"').checked=true;"
   Response.Write "  document.getElementById('"&tmpstrB&"').checked=true; document.getElementById('"&tmpstrHB&"').checked=true;"
   Response.Write "  document.getElementById('"&tmpstrC&"').checked=true; document.getElementById('"&tmpstrHC&"').checked=true;"
   Response.Write "  document.getElementById('"&tmpstrD&"').checked=true; document.getElementById('"&tmpstrHD&"').checked=true;"
   Response.Write "document.getElementById('"&tmpstrTS&"').checked=true;</script>"   
   Response.Write "<hr/>"
	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='document.getElementById(""edit"").action=""savesetting.asp"";' /></p>"
    Set objRegExp = nothing

  %>

       

</form>

</div>
<script language="JavaScript" type="text/javascript">
$(document).ready(function(){ 

try{
	//取消提示显示
	setTimeout("document.getElementById('ShowBlogHint').style.display='none'",2000);
	}
catch(e){};

});
</script>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
