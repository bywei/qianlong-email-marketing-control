<%
'注册插件
Call Registerplugin("Archiver", "Activeplugin_Archiver")

'安装插件
Function Installplugin_Archiver()
    'On Error Resume Next
    Call Archiver_Addto_Navbar()
    Call Archiver_Copy_Files()
    Call SetBlogHint_Custom("? 提示:[Archiver]已启用,你可以现在更新Archivers列表.")
    Response.Redirect "../plugin/Archiver/makeit.asp"
    Err.Clear
End Function

'卸载插件
Function UnInstallplugin_Archiver()
    Call Archiver_Delfrom_Navbar
    Call SetBlogHint_Custom("? 提示:[Archiver]已停用,由本插件生成的文件仍保留,请手动删除.")
End Function

Function Activeplugin_Archiver()
    '网站管理加上二级菜单项
    Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu",MakeSubMenu("[Archiver管理]","../plugin/Archiver/makeit.asp","m-left",False))

    'Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_TemplateTags","Archiver_Add_Code")
End Function

'添加到导航栏
Function Archiver_Addto_Navbar()
    Dim strContent
    strContent = LoadFromFile(BlogPath & "/include/navbar.asp", "utf-8")
    strContent = strContent & Chr(13)&Chr(10)& "<li><a href='"&ZC_BLOG_HOST&"Archiver/'>Archiver</a></li>"
    Call SaveToFile(BlogPath & "/include/navbar.asp", strContent, "utf-8", TRUE)
End Function

'插入底部版权前的代码
Function Archiver_Add_Code(ByRef aryTemplateTagsName, ByRef aryTemplateTagsValue)
Dim i,j
	j=UBound(aryTemplateTagsName)
	For i=1 to j
		If aryTemplateTagsName(i)="ZC_BLOG_COPYRIGHT" Then
		aryTemplateTagsValue(i)="<span id=""archiver""><a href="""&ZC_BLOG_HOST&"archiver/"">Archiver</a> </span> " & aryTemplateTagsValue(i)
		End If
	Next
End Function

'从导航删除
Function Archiver_Delfrom_Navbar()
    Dim strContent
    strContent = LoadFromFile(BlogPath & "/include/navbar.asp", "utf-8")
    strContent = Replace(strContent, Chr(13)&Chr(10)& "<li><a href='"&ZC_BLOG_HOST&"Archiver/'>Archiver</a></li>", "")
    Call SaveToFile(BlogPath & "/include/navbar.asp", strContent, "utf-8", TRUE)
End Function

'copy模板和文件到相应的文件夹
Function Archiver_Copy_Files()
    Dim strContent,strContent1,strContent2
    strContent = LoadFromFile(BlogPath & "plugin/archiver/template/archiver_list.html", "utf-8")
    Call SaveToFile(BlogPath & "Themes/" & ZC_BLOG_THEME & "/Template/archiver_list.html", strContent, "utf-8", TRUE)
    strContent1 = LoadFromFile(BlogPath & "plugin/archiver/template/archiver_view.html", "utf-8")
    Call SaveToFile(BlogPath & "Themes/" & ZC_BLOG_THEME & "/Template/archiver_view.html", strContent1, "utf-8", TRUE)
    Call GetReallyDirectory()
    CreatDirectoryByCustomDirectory("ARCHIVER")
    strContent2 = LoadFromFile(BlogPath & "plugin/archiver/source/view.asp", "utf-8")
    Call SaveToFile(BlogPath & "archiver/view.asp", strContent2, "utf-8", TRUE)
End Function
%>