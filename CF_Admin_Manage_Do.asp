
<%
If Action="lyreset" Then
 Sql="Update CFCount_Admin Set Store_TotalLy=0"
 Conn.ExeCute Sql
 Call AlertBack("复原成功",1)
End If

If Action="beforelydel" Then
 Sql="Delete From CFCount_Ly_Day Where DateDiff('d',AddDate,Now())>0"
 Conn.ExeCute Sql
 Call AlertBack("删除成功",1)
End If

If Action="beforewebdel" Then
 Sql="Delete From CFCount_Web_Day Where DateDiff('d',AddDate,Now())>0"
 Conn.ExeCute Sql
 Call AlertBack("删除成功",1)
End If

If Action="userdel" Then
 UserName=ChkStr(Request("UserName"),1)

 Filename=UserName&".jpg"
 Filename_2=UserName&".gif"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 If Fs.FileExists(Server.mappath("Upload/"&Filename_2)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename_2))
  Os.Delete
 End If

 Sql="Delete From CFCount_Alexa_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Back Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Hour Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Ly_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_User Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql

 Call AlertBack("删除成功!",1)
End If

If Action="usermodifysave" Then
 ID=ChkStr(Request("ID"),2)
 Pwd_New=ChkStr(Request("Pwd_New"),1)
 PasswordAnswer_New=ChkStr(Request("PasswordAnswer_New"),1)
 UserState=ChkStr(Request("UserState"),2)
 AdminDesc=ChkStr(Request("AdminDesc"),1)


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where ID="&ID
 Rs.Open Sql,Conn,3,2
 If Pwd_New<>"" Then Rs("Pwd")=Md5(Pwd_New,1)
 If PasswordAnswer_New<>"" Then
  Rs("PassWordAsk")="你的答案是?"
  Rs("PasswordAnswer")=Md5(PasswordAnswer_New,1)
 End If
 Rs("UserState")=UserState
 Rs("AdminDesc")=AdminDesc
 Rs.Update

 Call AlertBack("修改用户资料成功!",2)
End If

If Action="syssetmodifysave" Then
 SysTitle=GoBack(ChkStr(Request("SysTitle"),1),"请输入计数器名称!")
 SysTitle=Replace(SysTitle,"""","")'替换掉双引号
 SysTitle=Replace(SysTitle,"'","")'替换掉单引号
 
 TjTextName=GoBack(Left(ChkStr(Request("TjTextName"),1),8),"请输入计数器名称!")
 TjTextName=Replace(TjTextName,"""","")'替换掉双引号
 TjTextName=Replace(TjTextName,"'","")'替换掉单引号
 
 RegState=ChkStr(Request("RegState"),2)
 LyKeep=ChkStr(Request("LyKeep"),2)
 OnlineKeep=ChkStr(Request("OnlineKeep"),2)
 WebKeep=ChkStr(Request("WebKeep"),2)
 StyleTotal=GoBack(ChkStr(Request("StyleTotal"),2),"请输入图片样式数量!")
 LogoUrl=GoBack(ChkStr(Request("LogoUrl"),1),"请输入计数器系统上的Logo图片地址!")
 IpArea=ChkStr(Request("IpArea"),2)
 SkinType=ChkStr(Request("SkinType"),2)
 SkinColor=ChkStr(Request("SkinColor"),1)
 emailsendtype=ChkStr(Request("emailsendtype"),2)
 SmtpAddress=ChkStr(Request("SmtpAddress"),1)
 SmtpUser=ChkStr(Request("SmtpUser"),1)
 SmtpPassword=ChkStr(Request("SmtpPassword"),1)
 


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Admin"
 Rs.Open Sql,conn,3,2
 Rs("SysTitle")=SysTitle
 Rs("TjTextName")=TjTextName
 Rs("RegState")=RegState
 Rs("LyKeep")=LyKeep
 Rs("OnlineKeep")=OnlineKeep
 Rs("WebKeep")=WebKeep
 Rs("StyleTotal")=StyleTotal
 Rs("LogoUrl")=LogoUrl
 Rs("IpArea")=IpArea
 Rs("SkinType")=SkinType
 Rs("SkinColor")=SkinColor
 Rs("emailsendtype")=emailsendtype
 Rs("SmtpAddress")=SmtpAddress
 Rs("SmtpUser")=SmtpUser
 Rs("SmtpPassword")=SmtpPassword
 Rs.Update

 Call AlertBack("修改管理员资料成功!",1)
End if

If Action="othersetmodifysave" Then

 MyStr="SysCode="&Goback(ChkStr(Request("SysCode"),1),"请输入系统安全码")&"|||"
 MyStr=MyStr&"HourCountKeepDay="&Goback(ChkStr(Request("HourCountKeepDay"),1),"请输入小时统计保留天数")&"|||"
 MyStr=MyStr&"WebKeepDay="&Goback(ChkStr(Request("WebKeepDay"),1),"请输入网页浏览记录保留天数")&"|||"
 MyStr=MyStr&"LyKeepDay="&Goback(ChkStr(Request("LyKeepDay"),1),"请输入来源记录保留天数")&"|||"
 MyStr=MyStr&"SearchKeepDay="&Goback(ChkStr(Request("SearchKeepDay"),1),"请输入搜索引擎记录保留天数")&"|||"
 MyStr=MyStr&"KeywordKeepDay="&Goback(ChkStr(Request("KeywordKeepDay"),1),"请输入KeywordKeepDay")&"|||"
 MyStr=MyStr&"CountKeepDay="&Goback(ChkStr(Request("CountKeepDay"),1),"请输入天数记录保留天数")&"|||"

 Sql="Update CFCount_Admin set OtherSet='"&Mystr&"'"
 Conn.Execute Sql
 
 Application("CFCount_MySet") = Empty'清空Asp里的配置

 Call AlertBack("修改成功",1)
End If

If Action="searchaddsave" Then
 SiteDesc=GoBack(ChkStr(Lcase(Request("SiteDesc")),1),"请输入搜索引擎中文名")
 SiteFlag=GoBack(ChkStr(Lcase(Request("SiteFlag")),1),"请输入搜索引擎英文标识")
 KeywordFlag=GoBack(ChkStr(Lcase(Request("KeywordFlag")),1),"请输入搜索关键字的参数")
 SiteFlag=Replace(SiteFlag,",")
 SiteFlag=Replace(SiteFlag,"|")
 KeywordFlag=Replace(KeywordFlag,",")
 KeywordFlag=Replace(KeywordFlag,"|")


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"'"
 Rs.Open Sql,Conn,3,2
 If Not.Rs.Eof Then
  Call AlertBack("已经加入,不能重名!",1)
 Else
  Rs.AddNew
  Rs("SiteDesc")=SiteDesc
  Rs("SiteFlag")=SiteFlag
  Rs("KeywordFlag")=KeywordFlag
  Rs.Update
 End If

 Call AlertBack("加入成功!",1)
End If


If Action="searchmodifysave" Then
 ID=ChkStr(Request("ID"),2)
 SiteDesc=GoBack(ChkStr(Lcase(Request("SiteDesc")),1),"请输入搜索引擎中文名")
 SiteFlag=GoBack(ChkStr(Lcase(Request("SiteFlag")),1),"请输入搜索引擎英文标识")
 KeywordFlag=GoBack(ChkStr(Lcase(Request("KeywordFlag")),1),"请输入搜索关键字的参数")

 Sql="Select Count(*) From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"' And ID<>"&ID
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Rs.Open Sql,Conn,1,1
 If Rs(0)>0 Then Call AlertBack("已经加入,不能重名!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Search_Set Where ID="&ID
 Rs.Open Sql,Conn,3,2
 Rs("SiteDesc")=SiteDesc
 Rs("SiteFlag")=SiteFlag
 Rs("KeywordFlag")=KeywordFlag
 Rs.Update

 Call AlertBack("修改成功!",2)
End If

If Action="searchdel" Then
 SiteFlag=ChkStr(Request("SiteFlag"),1)
 Sql="Delete From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where SiteFlag='"&SiteFlag&"'"
 Conn.ExeCute Sql

 Call AlertBack("删除成功",1)
End If  '自定义删除

If Action="picdel" Then
 Filename=ChkStr(Request("Filename"),1)

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 Set Rs = Server.CreateObject("ADODB.Recordset")
 Sql="delete from CFCount_Upfile where Filename='"&Filename&"'"
 Rs.open sql,conn,3,2

 Call AlertBack("成功删除!",1)
End if

If Action="pwdmodifysave" Then
 Pwd_Old=GoBack(ChkStr(Trim(Request("Pwd_Old")),1),"请输入原管理员密码")
 Admin=GoBack(ChkStr(Trim(Request("Admin")),1),"请输入管理员名称")
 Pwd=GoBack(ChkStr(Trim(Request("Pwd")),1),"请输入密码")
 Pwd2=GoBack(ChkStr(Trim(Request("Pwd2")),1),"请输入重复密码")

 If Pwd<>Pwd2 Then Call AlertBack("输入的密码不一致，请重新输入一遍!",1)
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Admin"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd_Old,1)<>Rs("Pwd") Then Call AlertBack("旧密码输入有错误",1)

 Rs("Admin")=Admin
 Rs("Pwd")=Md5(Pwd,1)
 Rs.Update

 Call AlertBack("修改成功",1)
End If

If Action="alllydel" Then
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Delete From CFCount_Ly_Day"
 Rs.Open Sql,Conn,3,2
 Call AlertBack("删除成功!",1)
End if

If Action="adminlogout" Then
 Session("CFCountAdmin")=""
 Response.Cookies("CFCountAdminCookie")=""
 Response.Redirect "Admin.asp"
End If

If Action="usergo" Then
 Session("CFCountUser")=Request("UserName")	
 Response.Redirect("Manage.asp")
End If

If Action="sqlexesave" Then
 If AdminOnlineExecSql=0 Then Call AlertBack("Conn.asp里设置了系统禁止在线执行Sql语句",1)
 
 Pwd=GoBack(ChkStr(Trim(Request("Pwd")),1),"请输入确认密码")

 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select Pwd From CFCount_Admin"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd,1)<>Rs("Pwd") Then Call AlertBack("确认密码输入有错误",1)
 
 Sql = Request.Form("Sql")
 Conn.Execute(Sql)
  
 Call AlertBack("执行成功",1)
End If

If Action="dbys" Then
 YsDbName="backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_"&Md5(timer(),1)&".mdb"
 TgDbName="backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&(second(now)+1)&"_"&Md5(timer(),1)&".mdb"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Fs.copyfile Server.MapPath(DbPath),Server.MapPath("data\"&YsDbName)
 
  Const JET_3X = 4
  Set Engine = CreateObject("JRO.JetEngine")
  on error resume next
  Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data\"&YsDbName) , "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data\"&TgDbName)

  Fs.CopyFile Server.MapPath("data\"&TgDbName),Server.MapPath(DbPath)
  
 Call AlertBack("压缩成功",1)
End If

If Action="dbbackup" Then
 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Fs.copyfile Server.MapPath(DbPath),Server.MapPath("data\backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_"&Md5(timer(),1)&".mdb")
 Call AlertBack("备份数据库成功!",1)
End If '数据库备份结束


If Action="dbdel" Then
 DbMc=ChkStr(Request("DbMc"),1)
 
If Lcase(StrReverse(Left(StrReverse(DbMc),4)))<>".mdb" Then Call Alertback("只能删除扩展名为mdb的数据库文件",1)

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Set Os = Fs.GetFile(Server.mappath("data/"&DbMc))
 Os.Delete


 Call AlertBack("成功删除!",1)
End If
%>