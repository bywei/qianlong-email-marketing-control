
<%
  dim UserId
If Action="userloginsave" Then
 UserName=Trim(GoBack(ChkStr(Request("UserName"),1),"请输入用户名称"))
 Pwd=Trim(GoBack(ChkStr(Request("Pwd"),1),"请输入用户密码"))
 Cookies_Time=CLng(ChkStr(Request("Cookies_Time"),2))
 CheckCode=ChkStr(Trim(Request("CheckCode")),1)

 CFCountUserCookie=Md5(1&UserName&Pwd&GetMySet("SysCode"),1)

 If Session("ValidCode")<>CheckCode Then Call AlertUrl("四位数字的验证码输入错误!","Index.asp")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Rs.Eof And Rs.Bof Then Call AlertUrl("没有此注册用户!","Index.asp")
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertUrl("密码错误!","Index.asp")
 If cstr(Rs("EmaState"))<>cstr("true") Then Call AlertUrl("帐号未审核，请联系客服审核后使用","Index.asp")
 
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2 
 Rs("LoginTotal")=Rs("LoginTotal")+1
 Rs("LastLoginTime")=Now()
 Rs("UserCode")=CFCountUserCookie
 Rs.Update

 Session("CFCountUser")=UserName

 If Cookies_Time>0 Then
  Response.Cookies("CFCountUserCookie")=CFCountUserCookie
  Response.Cookies("CFCountUserCookie").expires=Dateadd("n",Cookies_Time,Now())
 End If

 Response.Cookies("CFCountGGCookie")="yes"
 Response.Cookies("CFCountGGCookie").expires=Dateadd("d",365,Now())

 Response.Redirect "Manage.asp"
End If


IF Action="adminloginsave" Then

Admin=GoBack(ChkStr(Request("Admin"),1),"请输入管理员名称")
Pwd=GoBack(ChkStr(Request("Pwd"),1),"请输入管理员密码")
Cookies_Time=CLng(ChkStr(Request("Cookies_Time"),2))
CheckCode=ChkStr(Trim(Request("CheckCode")),1)

Ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip = "" Then Ip=Request.ServerVariables("REMOTE_ADDR")

CFCountAdminCookie=Md5(0&Admin&Pwd&GetMySet("SysCode"),1)

If Session("ValidCode")<>CheckCode Then Call AlertUrl("四位数字的验证码输入错误!","Admin.asp")

Set Rs = Server.CreateObject("ADODB.Recordset")
Sql="Select * From CFCount_Admin where Admin='"&Admin&"'"
rs.open sql,conn,3,2

IF Rs.eof and Rs.bof Then Call AlertUrl("没有此管理员名称!","Admin.asp")
IF Rs("Pwd")<>Md5(Pwd,1) Then Call AlertUrl("密码错误!","Admin.asp")

 Rs("LoginTotal")=Rs("LoginTotal")+1
 Rs("LastLoginTime")=Now()
 Rs("AdminCode")=CFCountAdminCookie
 Rs.Update

 Session("CFCountAdmin")="ok"

 If Cookies_Time>0 Then
  Response.Cookies("CFCountAdminCookie")=CFCountAdminCookie
  Response.Cookies("CFCountAdminCookie").expires=Dateadd("n",Cookies_Time,Now())
 End If

 Response.Redirect "CF_Admin_Manage.asp"
End If



If Action="userchecksave" Then

 Response.Charset= "gb2312"

 UserName=ChkStr(Request("UserName"),1)
 If UserName="" Then
  Response.Write "请输入要检测的用户名"
  Response.End
 End if

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="Select Count(*) From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs(0)=0 Then
  Response.Write "恭喜，此用户还没有被注册!"
 Else
  Response.Write "对不起，此用户名已经被别人注册了!"
 End if  

 Response.End
End if

If Action="userregsave" Then

 If RsSet("RegState")=0 Then Call AlertBack("此系统已经暂停新用户注册!",1)

 UserName=GoBack(ChkStr(Trim(Request("UserName")),1),"请填入注册用户名!")
 Call CheckInput_Letter(UserName)
 Pwd=GoBack(ChkStr(Request("Pwd"),1),"请填入密码!")
 Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"请再次填入密码!")

 PasswordAsk=ChkStr(Request("PasswordAsk"),1)
 PasswordAnswer=ChkStr(Request("PasswordAnswer"),1)
 
 Email=GoBack(ChkStr(Request("Email"),1),"请输入E-mail!")
 
 QQ=ChkStr(Request("QQ"),1)
 PageName=GoBack(ChkStr(Request("PageName"),1),"请填入主页名称!")
 PageUrl=GoBack(ChkStr(Request("PageUrl"),1),"请填入主页网址!")
 
 QuickStyle=CInt(ChkStr(Request("QuickStyle"),2))
 TjOpen=ChkStr(Request("TjOpen"),2)
 ShowTotal=CLng(ChkStr(Request("ShowTotal"),2))
 
 If TjOpen="" Then TjOpen=0

 
 CheckCode=Trim(Request("CheckCode"))
 
 If Session("ValidCode")<>CheckCode Then Call AlertUrl("四位数字的验证码输入错误!","?Reg.asp")
 
 If Pwd<>Pwd2 Then Call AlertBack("填入的密码不一致，请重新输入一遍!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select Count(*) From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs(0)>0 Then Call AlertBack("此用户名已经被别人注册,请换名!",1)

'添加用户名到CFCount_user,以后不会把用户名存储到该表中
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs.AddNew
 Rs("UserName")=UserName
 Rs("Pwd")=Md5(Pwd,1)
 Rs("UserPwd")=Pwd
 Rs("UserEmail")=Email
 Rs("UseRscore")=0
 Rs("UserLevel")=1
 Rs("useRsendNum")="0"
 Rs("UserIsVip")="false"
 Rs("UseRstate")=1
 Rs("UserType")=1
 Rs("CreateTime")=Date()
 Rs("UseOfTime")=dateadd("M",1,date)
 Rs("EmaState")="true"
 Rs.Update
 
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 UserId= Rs("id")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserId="&UserId&""
 Rs.Open Sql,Conn,3,2
 If Not Rs.Eof Then
  Call AlertBack("此用户名已经被别人注册,请换名!",1)
 Else
 Rs.AddNew
 Rs("UserId")=UserId
 Rs("UserName")=UserName
 Rs("Pwd")=Md5(Pwd,1)
 Rs("Pwd_View")=Md5(Pwd,1)
 Rs("PasswordAsk")=PasswordAsk
 Rs("PasswordAnswer")=Md5(PasswordAnswer,1)
 Rs("Email")=Email
 Rs("QQ")=QQ
 Rs("PageName")=PageName
 Rs("PageUrl")=PageUrl
 Rs("ShowTotal")=ShowTotal
 Rs("TjOpen")=TjOpen
 Rs("UserCode")=Md5(1&UserName&Pwd&GetMySet("SysCode"),1)
 Rs.Update
 End If
 


 If QuickStyle=1 Or QuickStyle=4 Then
  CounterShow=-1
  ShowType=1
  CounterHiddenPic=2
  CounterSite=3
 End if

 If QuickStyle=2 Or QuickStyle=5 Then
  CounterShow=0
  ShowType=1
  CounterHiddenPic=2
  CounterSite=3
 End if
 
 If QuickStyle=3 Or QuickStyle=6 Then
  CounterShow=0
  ShowType=2
  CounterHiddenPic=1
  CounterSite=3
 End if
 
 If QuickStyle=7 Then
  CounterShow=0
  ShowType=2
  CounterHiddenPic=0
  CounterSite=3
 End if
 
 If QuickStyle=4 Or QuickStyle=5 Or QuickStyle=6 Then
  TjOpen=-1
  TjOpen_2=-1
 Else
  TjOpen_2=0
 End If
 
  Set Rs=Server.CreateObject("Adodb.RecordSet")
  Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
  Rs.Open Sql,Conn,3,2
  Rs("CounterShow")=CounterShow
  Rs("ShowType")=ShowType
  Rs("CounterHiddenPic")=CounterHiddenPic
  Rs("CounterSite")=CounterSite
  Rs("OnlineShow")=TjOpen_2
  Rs("TodayShow")=TjOpen_2
  Rs("TodayIpShow")=TjOpen_2
  Rs("IpShow")=TjOpen_2
  Rs("VisitShow")=TjOpen_2
  Rs("TjOpen")=TjOpen
  Rs("SurveyOpen")=TjOpen
  Rs("TodayLyOpen")=TjOpen
  Rs("OnlineOpen")=TjOpen
  Rs("TodayHourOpen")=TjOpen
  Rs("EveryDayOpen")=TjOpen
  Rs.Update


 Response.Cookies("CFCountGGCookie")="yes"
 Response.Cookies("CFCountGGCookie").expires=Dateadd("d",365,Now())

 Session("CFCountUser")=UserName

Call AlertUrl("注册成功,邮件营销效果跟踪系统后台!","Manage.asp")

End If




If Action="answermodifypwd" Then
UserName=ChkStr(request("UserName"),1)
PasswordAnswer=ChkStr(request("PasswordAnswer"),1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 UserId= Rs("id")

Set Rs= Server.CreateObject("ADODB.Recordset")
Sql="select * from CFCount_User where UserId="&UserId&""
Rs.Open Sql,Conn,1,1
If Not Rs.Eof Then
 If Rs("PasswordAnswer")<>Md5(PasswordAnswer,1) Then
  Call AlertUrl("回答答案有误！","?")
 End If
Else
 Call AlertUrl("用户名不存在！","Index.asp")
End If

Response.Redirect "?Action=pwdmodify&UserCode="&Rs("UserCode")

End If

If Action="pwdmodifysave" Then

Pwd=GoBack(ChkStr(Request("Pwd"),1),"请输入密码!")
Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"请输入重复密码!")
UserCode=GoBack(ChkStr(Request("UserCode"),1),"请输入较验码!")


If Pwd<>Pwd2 Then Call AlertBack("两次输入的密码不一致！",1)

Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserCode='" &UserCode& "'"
Rs.Open Sql,Conn,3,2

If Rs.Eof And Rs.Bof Then
 Call AlertBack("较验码不存在!",1)
Else
 Rs("Pwd")=Md5(Pwd,1)
 Rs.Update
End If

'Set Rs=Server.CreateObject("Adodb.RecordSet")
'Sql="Select * From Users Where Pwd='" &Pwd& "'"
'Rs.Open Sql,Conn,3,2
'Rs("Pwd")=Md5(Pwd,1)
'Rs("UserPwd")=Pwd
'Rs.Update

Call AlertUrl("修改成功，请牢记你的新密码","Index.asp")
End If



If Action="gbookaddsave" Then
 UserName=ChkStr(Request("UserName"),1)
 Content=ChkStr(Request("Content"),1)
 Contact=ChkStr(Request("Contact"),1)
 Ly=Left(ChkStr(URLDecode(Request("Ly")),1),255)
 CurrWeb=Left(ChkStr(URLDecode(Request("CurrWeb")),1),255)

  
 If Content="" Then Call AlertClose("请输入留言内容")
 If Contact="" Then Call AlertClose("请输入联系方式")

Sql="Select * From CFCount_Gbook Where UserName='"&UserName&"' And Content='"&Content&"' And Contact='"&Contact&"' And DateDiff('d',AddTime,Now())=0"
Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,Conn,3,2
If Not Rs.Eof Then
 Call AlertClose("请不要留相同的留言")
Else
 Rs.AddNew
 Rs("UserName")=UserName
 Rs("Content")=Content
 Rs("Contact")=Contact
 Rs("Ly")=Ly
 Rs("CurrWeb")=CurrWeb
 Rs.Update
End If

Call AlertClose("留言成功")

End If


'查看用户登录
If Action="viewuserlogin" Then
 UserName=GoBack(ChkStr(Request("UserName"),1),"请输入用户名称")
 Pwd_View=GoBack(ChkStr(Request("Pwd_View"),1),"请输入独立查看密码")
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 UserId= Rs("id")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserId="&UserId&""
 Rs.Open Sql,Conn,1,1

 If Rs.Eof And Rs.Bof Then Call AlertBack("没有此注册用户!",1)

 If Rs("Pwd_View")<>Md5(Pwd_View,1) Then Call AlertBack("独立查看密码错误!",1)
 
 Session("CFCountUser_View")=UserName
 Response.Redirect "Manage.asp"
End If
%>