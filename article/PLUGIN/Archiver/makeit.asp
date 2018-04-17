<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律
'// 备    注:    Archiver
'// 最后修改：   2011.8.30
'// 最后版本:    1.4
'///////////////////////////////////////////////////////////////////////////////
%>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_function_md5.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_event.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="zh-CN" />
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
</head>
<body>
	<div id="divMain">
		<div class="Header">生成Archiver (文章列表)</div>
<%
Call System_Initialize()
	
'检查非法链接
Call CheckReference("")

if BlogUser.Level>1 Then Call ShowError(6)
if CheckPluginState("archiver")=False Then Call ShowError(48)

Server.ScriptTimeOut=300

Dim strAct
'strAct=Request.QueryString("act")
strAct=Request("act")
Select Case strAct
	Case "StepO":Call StepO()
	Case "StepT":Call StepT()
	Case "StepTH":Call StepTH()
	Case "StepAll":Call StepAll()
	Case Else:Call Start()
End Select

'--释放内存缓存--
Function StepTH()
	Call FreeApplicationMemory
End Function

'--简单一页--
Function StepAll()
	Dim GroupBy,strSql,log_ViewNums
	GroupBy = Replace(Replace(Trim(Request("group")),"'","")," ","")
	If GroupBy = "PostTime" then 
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_PostTime desc"		
	ElseIf GroupBy = "Category" then	
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_CateID desc"
	Else 
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_ViewNums desc"
	End If	
	Response.write "<div id=""divMain2"">"
	Set objRS = CreateObject("ADODB.RECORDSET")		
		objRS.Open strSql, objConn, 1, 1
		IF not objRS.bof and not objRS.eof then
			Do while not objRS.eof
				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then 
					Set objCategory=New TCategory
					If objCategory.LoadInfoByID(objRS("log_CateID")) Then
						strnew = strnew & "<li>[<a href=""" & objCategory.Url & """ target=""_blank"">" & objCategory.Name & "</a>]&nbsp;&nbsp;<a href=""" & objArticle.Url & """  target=""_blank"">" & objRS("log_Title") & "</a> &nbsp;(" & objArticle.ViewNums & ")&nbsp;" & objArticle.PostTime & " </li>" & vbcrlf
					End if
					Set objCategory = Nothing
				End If
				Set objArticle = Nothing
				objRS.movenext
			Loop
			objRS.Close
		Else
			Response.write "<p>似乎还没有日志!</p>"
			Exit Function
		End If
	Set objRS = Nothing 
	
	Call GetReallyDirectory()
	
	'输出文件地址
	txtfilename =BlogPath & "archiver\index.html"
   
	Call System_Initialize

	LoadGlobeCache

	Dim ArtList
	Set ArtList=New TArticleList

	ArtList.LoadCache

	If LoadFromFile(BlogPath & "Themes/" & ZC_BLOG_THEME & "/Template/archiver_list.html","utf-8")<>"" Then

		ArtList.template="ARCHIVER_LIST"
		ArtList.Title=ARCHIVER
		ArtList.Build
		Dim Html
		Html=ArtList.html		
		Html=Replace(Html,"<#CUSTOM_TAGS#>",strnew)
		Call SaveToFile(txtfilename,html,"utf-8",False)
	
	Else
	
		templatefile =BlogPath & "PLUGIN\archiver\" & "template\list.html"
		txttemplate = LoadFromFile(templatefile,"utf-8")
		'模板文件替换
		txtout = txttemplate
		txtout = Replace(txtout, "<!--AllList-->", strnew)
		txtout = Replace(txtout, "<!--SiteName-->", ZC_BLOG_NAME)
		txtout = Replace(txtout, "<!--SiteURL-->", ZC_BLOG_HOST)
		txtout = Replace(txtout, "<!--ThePage-->", 1)
		txtout = Replace(txtout, "<!--AllPage-->", 1)
		txtout = Replace(txtout, "<!--ReCounts-->", 1)
		txtout = Replace(txtout, "<!--PostTime-->", Year(now)&"年"&month(now)&"月"&day(now)&"日 "&hour(now)&":"&minute(now)&":"&second(now))
		Call SaveToFile(txtfilename,txtout,"utf-8",False)

	End If
	
	Response.Write "<p>√任务完成! <a href=""makeit.asp"">返回插件</a> <a href="""& ZC_BLOG_HOST &"archiver/"" target=""_blank"">查看Archiver</a></p>"
	Response.write "<p>你可以在适当的位置添加下面的代码:</p>"
	Response.Write "<p><input type=""text"" style=""width:70%"" value="""&Server.HTMLEncode("<a href=""<"&Chr("37")&"=ZC_BLOG_HOST"&Chr("37")&">archiver.htm"">Archiver</a>")&"""></p>"	
	Response.write "</div>"
End Function

'--生成过程--
Function StepT()
	Dim PageCounts
	Set objRS = CreateObject("ADODB.RECORDSET")		
		objRS.Open Application("strSql"), objConn, 1, 1
		objRS.PageSize=Application("PageList")
		PageCounts = objRS.PageCount
		Dim pagenum,i,strnew
		pagenum=Request("p")
		If pagenum="" then pagenum=1 
		If not isInteger(pagenum) then
			Response.write "<div id=""divMain2"">参数错误!</div>"
			Exit Function
		End if
		pagenum=cint(pagenum)
		IF not objRS.bof and not objRS.eof then
			If pagenum<=1 then
				objRS.AbsolutePage=1
				pagenum=1
			ElseIf pagenum>objRS.pagecount then
				Response.write "<meta http-equiv=""refresh"" content=""3;url='?act=StepTH'""><div id=""divMain2"">已经完成更新任务!&nbsp;&nbsp;"&now()
				Response.write "<br>3秒后释放内存缓存</div>"
				Exit Function
			Else
				objRS.AbsolutePage=pagenum			
			End if
			i=1
			Do while not objRS.eof and i<=objRS.PageSize
				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then 
					Set objCategory=New TCategory
					If objCategory.LoadInfoByID(objRS("log_CateID")) Then
						strnew = strnew & "<li>[<a href=""" & objCategory.Url & """ target=""_blank"">" & objCategory.Name & "</a>]&nbsp;&nbsp;<a href=""view.asp?id=" & objRS("log_ID") & """>" & objRS("log_Title") & "</a>&nbsp;(" & objArticle.ViewNums & ")&nbsp; " & objArticle.PostTime & " </li>" & vbcrlf
					End if
					Set objCategory = Nothing
				End If
				Set objArticle = Nothing
				i=i+1
				objRS.movenext
			Loop
			objRS.Close
		Else
			Response.write "<div id=""divMain2"">似乎还没有日志!</div>"
			Exit Function
		End If
	Set objRS = Nothing 
	
	Call GetReallyDirectory()
	CreatDirectoryByCustomDirectory("ARCHIVER")
	'输出文件地址
	If pagenum = 1 then
		txtfilename = BlogPath & "archiver\index.html"
	Else
		txtfilename =BlogPath & "archiver\archiver"&pagenum&".html"
	End If
   
	Call System_Initialize

	LoadGlobeCache

	Dim ArtList
	Set ArtList=New TArticleList

	ArtList.LoadCache

	If LoadFromFile(BlogPath & "Themes/" & ZC_BLOG_THEME & "/Template/archiver_list.html","utf-8")<>"" Then

		ArtList.template="ARCHIVER_LIST"

		ArtList.Title=ARCHIVER

		ArtList.Build
		Dim Html
		Html=ArtList.html
		
		Html=Replace(Html,"<#CUSTOM_TAGS#>",strnew & "<p>" & Application("archiverList") & " 共" & PageCounts & "页&nbsp; 当前第" & pagenum & "页&nbsp;每页" & Application("PageList") & "篇日志&nbsp;&nbsp;</p>")
	   
		Call SaveToFile(txtfilename,html,"utf-8",False)
	
	Else
	
		templatefile =BlogPath & "PLUGIN\archiver\" & "template\list.html"

		txttemplate = LoadFromFile(templatefile,"utf-8")
		'模板文件替换
		txtout = txttemplate
		txtout = Replace(txtout, "<!--AllList-->", strnew)
		txtout = Replace(txtout, "<!--SiteName-->", ZC_BLOG_NAME)
		txtout = Replace(txtout, "<!--SiteURL-->", ZC_BLOG_HOST)
		txtout = Replace(txtout, "<!--ThePage-->", pagenum)
		txtout = Replace(txtout, "<!--AllPage-->", PageCounts)
		txtout = Replace(txtout, "<!--ReCounts-->", Application("PageList"))
		txtout = Replace(txtout, "<!--archiverList-->", Application("archiverList"))
		txtout = Replace(txtout, "<!--PostTime-->", Year(now)&"年"&month(now)&"月"&day(now)&"日 "&hour(now)&":"&minute(now)&":"&second(now))
		Call SaveToFile(txtfilename,txtout,"utf-8",False)
	
	End If
	
	Response.Write "<meta http-equiv=""refresh"" content=""1;url='?act=StepT&p="&pagenum+1&"'""><div id=""divMain2"">已生成第"&pagenum&"页,1秒后生成下一页!</div>"

End Function

'--初始化排序--
Function StepO()
	Dim PageList,GroupBy,OrderBy,strSql,log_ViewNums,archiverList,PageCounts
	PageList = Request("recount") 
	If not isInteger(PageList) then
		Response.write "<div id=""divMain2"">记录数参数填写错误!</div>"
		Exit Function
	End if
	GroupBy = Replace(Replace(Trim(Request("group")),"'","")," ","")
	If GroupBy = "PostTime" then 
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_PostTime desc"		
	ElseIf GroupBy = "Category" then	
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_CateID desc"
	Else 
		strSql="select * from blog_Article WHERE ([log_ID]>0) AND ([log_Level]>2) order by log_ViewNums desc"
	End If		
	Set objRS = CreateObject("ADODB.RECORDSET")
		objRS.Open strSql, objConn, 1, 1
		If not objRS.bof and not objRS.eof then
			objRS.PageSize=PageList
			PageCounts = objRS.PageCount
		Else
			Response.write "<div id=""divMain2"">似乎还没有日志!</div>"
			Exit Function
		End If
		objRS.close
	Set objRS = Nothing
	If PageCounts>1 then
		archiverList = "分页列表:<a href=index.html>1</a>&nbsp;"&VBcrlf		
		For i=2 to PageCounts
			archiverList = archiverList & "<a href=archiver"&i&".html>"&i&"</a>&nbsp;"&VBcrlf
		Next
	Else
		archiverList = "分页列表:<a href=index.html>1</a>"&VBcrlf
	End If
	Application.Lock() 
		Application("PageList") = PageList
		Application("strSql") = strSql
		Application("archiverList") = archiverList
	Application.UnLock() 
	Response.Write "<meta http-equiv=""refresh"" content=""3;url='?act=StepT&p=1'""><div id=""divMain2"">已经完成初始化,稍等3秒开始更新!</div>"
End Function

Function Start()
%>
		<div id="divMain2">
		<%Call GetBlogHint()%>
			<p>如果没有添加新日志,则不需要更新!  </p>
			<form id="edit" name="edit" method="get" action="">
				<input name="act" type="hidden" value="StepO" id="StepO"  />	
				<p>更新分页的Archiver! 此方式适用于内容较多的站点. 如果文章较多,更新可能需要较长时间.</p>
				<p>每页多少条记录:<input name="recount" type="text" value="80" id="recount" style="width: 30px;"/></p>
				<p><input name="group" type="radio" value="PostTime" id="PostTime" checked="checked" />按时间排序&nbsp;<input name="group" type="radio" value="log_ViewNums" id="log_ViewNums" />按访问量排序&nbsp;<input name="group" type="radio" value="Category" id="Category" />按分类排序</p>			
				<p><input class="button" type="submit" value="更新生成分页式列表" id="btnPost"/></p>
			</form>
			<form id="edit" name="All" method="get" action="">
				<input name="act" type="hidden" value="StepAll" id="StepAll"  />	
				<p>更新不分页的Archiver! 此方式适用于不超过500篇日志的站点. 如果文章较多,更新可能需要较长时间.</p>
				<p><input name="group" type="radio" value="PostTime" id="PostTime" checked="checked" />按时间排序&nbsp;<input name="group" type="radio" value="log_ViewNums" id="log_ViewNums" />按访问量排序&nbsp;<input name="group" type="radio" value="Category" id="Category" />按分类排序</p>
				<p><input class="button" type="submit" value="更新生成单页式列表" id="btnPost2"/></p>
			</form>
			<br />
			<p style="color:#666">插件生成的存档文件在BLOG的Archiver目录下；</p>
			<p style="color:#666">插件生成的存档页面底部有赞助商链接，不喜者请自行从模板中移除，模版文件:achiver_list.html和achiver_view.html。</p>
		</div>
<%
End Function

'*************************************
'检测是否有效的数字
'*************************************
Function IsInteger(Para) 
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

'*************************************
'释放内存缓存
'*************************************
Function FreeApplicationMemory
	Response.write "<div id=""divMain2"">"&VBcrlf
	Response.Write "<p>已经释放内存缓存.</p>"&VBcrlf
	IF isObject(Application("PageList")) Then
		Application("PageList").Close
		Set Application("PageList") = Nothing
		Application("PageList") = Null
		Response.Write "<p>PageList</p>"&VBcrlf
	Else
		Application("PageList") = Null
	End IF
	IF isObject(Application("strSql")) Then
		Application("strSql").Close
		Set Application("strSql") = Nothing
		Application("strSql") = Null
		Response.Write "<p>strSql</p>"&VBcrlf
	Else
		Application("strSql") = Null
	End IF
	IF isObject(Application("archiverList")) Then
		Application("archiverList").Close
		Set Application("archiverList") = Nothing
		Application("archiverList") = Null
		Response.Write "<p>archiverList</p>"&VBcrlf
	Else
		Application("archiverList") = Null
	End IF
	Response.Write "<p>√任务完成! <a href=""makeit.asp"">返回插件</a> <a href="""& ZC_BLOG_HOST &"archiver/"" target=""_blank"">查看Archiver</a></p>"
	Response.write "<p>你可以在适当的位置添加下面的代码:</p>"
	Response.Write "<p><input type=""text"" style=""width:70%"" value="""&Server.HTMLEncode("<a href=""<"&Chr("37")&"=ZC_BLOG_HOST"&Chr("37")&">archiver/"">Archiver</a>")&"""></p>"	
	Response.write "</div>"
End Function
%>
	</div>
</body>
</html>