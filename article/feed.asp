<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    sydication.asp
'// 开始时间:    2006.07.30
'// 最后修改:    
'// 备    注:    改名叫feed.asp
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="c_option.asp" -->
<!-- #include file="function/c_function.asp" -->
<!-- #include file="function/c_function_md5.asp" -->
<!-- #include file="function/c_system_lib.asp" -->
<!-- #include file="function/c_system_base.asp" -->
<!-- #include file="function/c_system_event.asp" -->
<!-- #include file="function/rss_lib.asp" -->
<!-- #include file="function/atom_lib.asp" -->
<!-- #include file="function/c_system_plugin.asp" -->
<!-- #include file="plugin/p_config.asp" -->
<%
Call System_Initialize()

'plugin node
For Each sAction_Plugin_Feed_Begin in Action_Plugin_Feed_Begin
	If Not IsEmpty(sAction_Plugin_Feed_Begin) Then Call Execute(sAction_Plugin_Feed_Begin)
Next


Dim strAct
strAct="rss"

'如果不是"接收引用"就要检查非法链接
If (strAct<>"tb") And (strAct<>"search") Then Call CheckReference("")

'权限检查
If Not CheckRights(strAct) Then Call ShowError(6)


'/////////////////////////////////////////////////////////////////////////////////
If Not IsEmpty(Request.QueryString("cate")) Then
	Call ExportRSSbyCate(Request.QueryString("cate"))
ElseIf Not IsEmpty(Request.QueryString("tags")) Then
	Call ExportRSSbyTags(Request.QueryString("tags"))
ElseIf Not IsEmpty(Request.QueryString("user")) Then
	Call ExportRSSbyUser(Request.QueryString("user"))
ElseIf Not IsEmpty(Request.QueryString("date")) Then
	Call ExportRSSbyDate(Request.QueryString("date"))
ElseIf Not IsEmpty(Request.QueryString("cmt")) Then
	Call ExportRSSbyCmt(Request.QueryString("cmt"))
ElseIf Not IsEmpty(Request.QueryString("gb")) Then
	Call ExportRSSbyGuestBook()
ElseIf Not IsEmpty(Request.QueryString("atom")) Then
	Call ExportATOM()
	Response.ContentType = "text/xml"
	Response.Write LoadFromFile(BlogPath & "rss.xml" ,"utf-8")
Else
	Response.ContentType = "text/xml"
	Response.Write LoadFromFile(BlogPath & "rss.xml" ,"utf-8")
End If




'/////////////////////////////////////////////////////////////////////////////////
Function ExportRSSbyCate(CateID)

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TRss2Export

	With Rss2Export

		Call CheckParameter(CateID,"int",0)

		Dim objRS,CateName,CateIntro

		.TimeZone=ZC_TIME_ZONE

		CateName=Categorys(CateID).Name

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE & " - " & CateName,"[html-format]")
		.AddChannelAttribute "link",TransferHTML(ZC_BLOG_HOST,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE & " - " & CateIntro,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "copyright",TransferHTML(ZC_BLOG_COPYRIGHT,"[nohtml][html-format]")
		.AddChannelAttribute "pubDate",Now

				Dim i

				Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2) AND ([log_CateID]="&CateID&") ORDER BY [log_PostTime] DESC")

				If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then

						If ZC_RSS_EXPORT_WHOLE Then
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						Else
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
				End If

	End With

	objRS.close
	Set objRS=Nothing

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function




Function ExportRSSbyUser(UserID)

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TRss2Export

	With Rss2Export

		Call CheckParameter(UserID,"int",0)

		Dim objRS,UserName,UserIntro

		.TimeZone=ZC_TIME_ZONE

		UserName=Users(UserID).Name

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE & " - " & UserName,"[html-format]")
		.AddChannelAttribute "link",TransferHTML(ZC_BLOG_HOST,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE & " - " & UserIntro,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "copyright",TransferHTML(ZC_BLOG_COPYRIGHT,"[nohtml][html-format]")
		.AddChannelAttribute "pubDate",Now

				Dim i

				Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2) AND ([log_AuthorID]="&UserID&") ORDER BY [log_PostTime] DESC")

				If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then

						If ZC_RSS_EXPORT_WHOLE Then
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						Else
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
				End If

	End With

	objRS.close
	Set objRS=Nothing

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function




Function ExportRSSbyDate(YearMonth)

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TRss2Export

	With Rss2Export

		Call CheckParameter(YearMonth,"dtm",Empty)

		Dim objRS,UserName,UserIntro

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE & " - " & UserName,"[html-format]")
		.AddChannelAttribute "link",TransferHTML(ZC_BLOG_HOST,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE & " - " & UserIntro,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "copyright",TransferHTML(ZC_BLOG_COPYRIGHT,"[nohtml][html-format]")
		.AddChannelAttribute "pubDate",Now

				Dim i

				Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2) AND (Year([log_PostTime])="&Year(YearMonth)&") AND (Month([log_PostTime])="&Month(YearMonth)&") ORDER BY [log_PostTime] DESC")

				If (Not objRS.bof) And (Not objRS.eof) Then
				Do While Not objRS.eof
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then

						If ZC_RSS_EXPORT_WHOLE Then
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						Else
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit Do
					Set objArticle=Nothing
				Loop
				End If

	End With

	objRS.close
	Set objRS=Nothing

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function




Function ExportRSSbyTags(TagsID)

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TRss2Export

	With Rss2Export

		Call CheckParameter(TagsID,"int",0)

		Dim objRS,TagName

		.TimeZone=ZC_TIME_ZONE

		TagName=Tags(TagsID).Name

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE,"[html-format]") & " - " & TagName
		.AddChannelAttribute "link",TransferHTML(ZC_BLOG_HOST,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]") & " - " & TagName
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "copyright",TransferHTML(ZC_BLOG_COPYRIGHT,"[nohtml][html-format]")
		.AddChannelAttribute "pubDate",Now

				Dim i

				Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2) AND ([log_Tag] LIKE '%{"&TagsID&"}%') ORDER BY [log_PostTime] DESC")

				If (Not objRS.bof) And (Not objRS.eof) Then
				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"))) Then

						If ZC_RSS_EXPORT_WHOLE Then
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						Else
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).Name & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro,Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
				End If

	End With

	objRS.close
	Set objRS=Nothing

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function




Function ExportRSSbyGuestBook()

	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TRss2Export

	With Rss2Export

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE & "-" & ZC_MSG275,"[html-format]")
		.AddChannelAttribute "link",ZC_BLOG_HOST & "guestbook.asp"
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
		.AddChannelAttribute "generator","Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "pubDate",Now

		Dim objRS
		Dim objComment
		Set objRS=objConn.Execute("SELECT TOP " & ZC_RSS2_COUNT & " * FROM [blog_Comment] WHERE ([comm_ID]>0) AND ([log_ID]=0) ORDER BY [comm_PostTime] DESC")

		Do While Not objRS.eof

			Set objComment=New TComment
			If objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"))) Then
				.AddItem objComment.Author & " Re:"&ZC_MSG275,objComment.Email & " (" & objComment.Author & ")",ZC_BLOG_HOST & "guestbook.asp" & "#cmt" & objComment.ID,objComment.PostTime,ZC_BLOG_HOST & "guestbook.asp" & "#cmt" & objComment.ID,objComment.HtmlContent,"","","","",""
			End If
			Set objComment=Nothing

			objRS.MoveNext

		Loop

		objRS.close
		Set objRS=Nothing

		Set objArticle=Nothing

	End With

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function




Function ExportRSSbyCmt(intID)

	Dim Rss2Export
	Dim objArticle
	Dim objRS
	Dim objComment

	Set Rss2Export = New TRss2Export

	With Rss2Export

		Call CheckParameter(intID,"int",0)

		Set objArticle=New TArticle
		If intID = 0 Then

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE & "-" & ZC_MSG027,"[html-format]")
		.AddChannelAttribute "link",ZC_BLOG_HOST
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
		.AddChannelAttribute "generator","Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "pubDate",Now


			Set objRS=objConn.Execute("SELECT TOP " & ZC_RSS2_COUNT & " * FROM [blog_Comment] WHERE ([comm_ID]>0) AND ([log_ID]>0) ORDER BY [comm_PostTime] DESC")

			Do While Not objRS.eof

				Set objComment=New TComment
				If objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"))) And objArticle.LoadInfoByID(objRS("log_ID")) Then
					.AddItem objComment.Author & " Re:"&objArticle.HtmlTitle,objComment.Email & " (" & objComment.Author & ")",objArticle.HtmlUrl & "#cmt" & objComment.ID,objComment.PostTime,objArticle.HtmlUrl & "#cmt" & objComment.ID,objComment.HtmlContent,"","","","",""
				End If
				Set objComment=Nothing

				objRS.MoveNext

			Loop

			objRS.close
			Set objRS=Nothing

		ElseIf objArticle.LoadInfoByID(intID) Then

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE,"[html-format]")& "-" & objArticle.HtmlTitle
		.AddChannelAttribute "link",objArticle.HtmlUrl
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "pubDate",objArticle.PostTime

			If objArticle.CommNums>0 Then

				Set objRS=objConn.Execute("SELECT * FROM [blog_Comment] WHERE ([comm_ID]>0) AND ([log_ID]="&intID&") ORDER BY [comm_PostTime] DESC")

				Do While Not objRS.eof

					Set objComment=New TComment
					If objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"))) Then
						.AddItem "Re:"&objArticle.HtmlTitle,objComment.Email & " (" & objComment.Author & ")",objArticle.HtmlUrl & "#cmt" & objComment.ID,objComment.PostTime,objArticle.HtmlUrl & "#cmt" & objComment.ID,objComment.HtmlContent,"","","","",""
					End If
					Set objComment=Nothing

					objRS.MoveNext

				Loop

				objRS.close
				Set objRS=Nothing
			End If
		End If
		Set objArticle=Nothing

	End With

	Rss2Export.Execute

	Set Rss2Export = Nothing

End Function
'/////////////////////////////////////////////////////////////////////////////////


'plugin node
For Each sAction_Plugin_Feed_End in Action_Plugin_Feed_End
	If Not IsEmpty(sAction_Plugin_Feed_End) Then Call Execute(sAction_Plugin_Feed_End)
Next

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>