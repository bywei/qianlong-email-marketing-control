
<%
DbPath="data/cfcount.mdb" '数据库名称设置，使用默认值的话至少要设置上面的SysCode以保证安全,最好把qqcf_data里的默认数据库名称换一个名称，换好后在这个填写。


On Error Resume Next
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(DbPath)
Set Conn = Server.CreateObject ("Adodb.Connection")
Conn.Open ConnStr

If Err Then
 err.Clear
 Set Conn = Nothing
 Response.Write "数据库连接出错，请检查连接字串。"
 Response.End
End If

InjGuard=0 '是否开启通用防注入，0为不开启，1为开启，注程序中已经做过防注入处理了,这里一般设置关闭就可以了
AdminOnlineExecSql=0 '是否允许超级管理员后台数据库管理里在线执行Sql，0为否，1为是，默认为否

If InjGuard=1 Then
SQL_injdata = "'|'"
SQL_inj = split(SQL_Injdata,"|") 

If Request.QueryString<>"" Then 
 For Each SQL_Get In Request.QueryString 
  For SQL_Data=0 To Ubound(SQL_inj) 
   if instr(Lcase(Request.QueryString(SQL_Get)),Sql_Inj(Sql_DATA))>0 Then 
    Response.Write "网址中包含非法字符"&Sql_Inj(Sql_DATA)
    Response.End 
   End If
  Next
 Next
End If

If Request.Form<>"" Then 
 For Each Sql_Post In Request.Form 
  For SQL_Data=0 To Ubound(SQL_inj) 
   if instr(Lcase(Request.Form(Sql_Post)),Sql_Inj(Sql_DATA))>0 Then 
    Response.Write "表单中包含非法字符"&Sql_Inj(Sql_DATA)
    Response.end 
   End If 
  Next 
 Next 
End If

If Request.Cookies<>"" Then
 For Each Sql_Cookies In Request.Cookies
  For SQL_Data=0 To Ubound(SQL_inj)
   if instr(Lcase(Request.Cookies(Sql_Cookies)),Sql_Inj(Sql_DATA))>0 Then
    Response.Write "Cookies中包含非法字符"&Sql_Inj(Sql_DATA)&",请在 IE选项-常规 里删除Cookiess再执行操作!"
    Response.end
   End If
  Next
 Next
End If
End If

Set RsSet=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_Admin"
RsSet.Open Sql,Conn,1,1
%>