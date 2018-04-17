
<%
If Action="subaddsave" Or Action="subgoto" Or Action="submodifysave" Or Action="subdel" Or Action="jsqsetsave" Or Action="jsqquicksetsave" Or Action="datamodifysave" Or Action="pwdmodifysave" Or Action="passwordanswermodifysave" Or Action="stylemodifysave" Or Action="imgmodifysave" Or Action="jsqresetsave" Or Action="userdelsave" Or Action="gbookmodifysave" Or Action="gbookdel" Then
 If Session("CFCountUser_View")<>"" Then'为独立查看者时
  Call AlertBack("独立查看者无法对此操作",1)
 End If
End If

If Action="subaddsave" Then
 SubUserName=GoBack(ChkStr(Request("SubUserName"),1),"请填入子账号用户名!")
 Call CheckInput_Letter(SubUserName)
 
 Pwd=GoBack(ChkStr(Request("Pwd"),1),"请填入子账号密码!")
 Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"请填入重复子账号密码!")
 
 If Pwd<>Pwd2 Then Call AlertBack("重复密码有误!",1)
 
 PageName=GoBack(Server.HtmlEncode(ChkStr(Request("PageName"),1)),"请填入站点名称!")
 PageUrl=GoBack(Server.HtmlEncode(ChkStr(Request("PageUrl"),1)),"请填入域名!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,3,2
 If Not Rs.Eof Then
  Call AlertBack("此用户名已经被别人注册!",1)
 Else
  Rs.AddNew
  Rs("UserName")=SubUserName
  Rs("Pwd")=Md5(Pwd,1)
  Rs("ParentName")=UserName
  Rs("PageName")=PageName
  Rs("PageUrl")=PageUrl
  Rs.Update
 End If
 
 Call AlertBack("注册成功!",2)
End If


If Request("Action")="subgoto" Then
 SubUserName=ChkStr(Request("SubUserName"),1)
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("ParentName")<>UserName Then Call AlertBack("没有权限",1)
 
 Session("CFCountSubAdmin")=UserName
 Session("CFCountUser")=SubUserName
 Response.Redirect("Manage.asp")
End If

If Request("Action")="parentgoto" Then
 Session("CFCountUser")=Session("CFCountSubAdmin")
 Session("CFCountSubAdmin")=""
 Response.Redirect("Manage.asp")
End If

If Action="submodifysave" Then

 SubUserName=ChkStr(Request("SubUserName"),1)
 
 PageName=GoBack(Server.HtmlEncode(ChkStr(Request("PageName"),1)),"请填入站点名称!")
 PageUrl=GoBack(Server.HtmlEncode(ChkStr(Request("PageUrl"),1)),"请填入域名!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&SubUserName&"' And ParentName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PageName")=PageName
 Rs("PageUrl")=PageUrl
 Rs.Update
 
 Call AlertBack("修改成功!",2)
End If

If Action="subdel" Then '账号删除
 SubUserName=ChkStr(Request("SubUserName"),1)
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("ParentName")<>UserName Then Call AlertBack("没有权限",1)

 Filename=SubUserName&".jpg"
 Filename_2=SubUserName&".gif"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 If Fs.FileExists(Server.mappath("Upload/"&Filename_2)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename_2))
  Os.Delete
 End If

 Sql="Delete From CFCount_Alexa_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Back Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Hour Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Ly_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_User Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql

 Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql

 Call AlertBack("这个子账号已经完全删除！",1)
End If
 

If Action="jsqsetsave" Then '计数器设置
 ShowTotal=GoBack(ChkStr(Request("ShowTotal"),2),"请输入计数显示数字!")
 PicNum=GoBack(ChkStr(Request("PicNum"),2),"请输入计数器显示位数!")
 Adddate=GoBack(ChkStr(Request("Adddate"),3),"请输入计数器开始计数时间!")
 CounterShow=CInt(Request("CounterShow"))
 ShowType=Int(Request("ShowType"))
 CounterHiddenPic=Int(Request("CounterHiddenPic"))
 CounterSite=Int(Request("CounterSite"))
 ImgCounterShow=Int(Request("ImgCounterShow"))
 ImgShowType=Int(Request("ImgShowType"))
 Online=Int(Request("Online"))
 OnlineTime=Request("OnlineTime")
 OnlineShow=Int(Request("OnlineShow"))
 TodayShow=Int(Request("TodayShow"))
 TodayIpShow=Int(Request("TodayIpShow"))
 IpShow=Int(Request("IpShow"))
 VisitShow=Int(Request("VisitShow"))

 TjOpen=Int(Request("TjOpen"))
 SurveyOpen=Int(Request("SurveyOpen"))
 TodayLyOpen=Int(Request("TodayLyOpen"))
 OnlineOpen=Int(Request("OnlineOpen"))
 TodayHourOpen=Int(Request("TodayHourOpen"))
 EveryDayOpen=Int(Request("EveryDayOpen"))
 
 GbookState=Int(Request("GbookState"))
 
 If Int(Picnum)<1 Or Int(Picnum)>8 Then Call AlertBack("请输入计数器显示位数在1-8之间!",1)
 If Onlinetime="" Then Call AlertBack("请输入记录在线人数的间隔时间!",1)
 If onlinetime<1 Or onlinetime>180 Then Call AlertBack("只能设置180分钟以内,一般设置在30分钟!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PicNum")=PicNum
 Rs("ShowTotal")=ShowTotal
 Rs("AddDate")=AddDate
 Rs("CounterShow")=CounterShow
 Rs("ShowType")=ShowType
 Rs("CounterHiddenPic")=CounterHiddenPic
 Rs("CounterSite")=CounterSite
 Rs("ImgCounterShow")=ImgCounterShow
 Rs("ImgShowType")=ImgShowType
 Rs("Online")=Online
 Rs("OnlineTime")=OnlineTime
 Rs("OnlineShow")=OnlineShow
 Rs("TodayShow")=TodayShow
 Rs("TodayIpShow")=TodayIpShow
 Rs("IpShow")=IpShow
 Rs("VisitShow")=VisitShow
 Rs("TjOpen")=TjOpen
 Rs("SurveyOpen")=SurveyOpen
 Rs("TodayLyOpen")=TodayLyOpen
 Rs("OnlineOpen")=OnlineOpen
 Rs("TodayHourOpen")=TodayHourOpen
 Rs("EveryDayOpen")=EveryDayOpen
 Rs("GbookState")=GbookState
 Rs.Update
 
 Response.Cookies("CFCountCookie")=""
 Call AlertUrl("修改成功!","?Action=jsqset")
End If

If Action="jsqquicksetsave" Then '计数器基本样式设置
 QuickStyle=Int(Request("QuickStyle"))
 TjOpen=Int(Request("TjOpen"))
 
 If QuickStyle=0 Then Call AlertBack("请选择一种基本样式!",1)

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
 
 Call AlertUrl("设置成功，转到预览计数的页面!","?Action=getcode")
End If

If Action="datamodifysave" Then
 Pwd=Goback(ChkStr(Request("Pwd"),1),"请输入密码确认")
 Email=GoBack(ChkStr(Request("Email"),1),"请输入E-mail!")
 QQ=ChkStr(Request("QQ"),1)
 PageName=GoBack(ChkStr(Request("PageName"),1),"请填入主页名称!")
 PageUrl=GoBack(ChkStr(Request("PageUrl"),1),"请填入主页网址!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("密码输入有误",1)

 Rs("Email")=Email
 Rs("QQ")=QQ
 Rs("PageName")=PageName
 Rs("PageUrl")=PageUrl
 Rs.Update
 
 Call AlertBack("修改成功!",1)
End If

If Action="pwdmodifysave" Then '用户修改密码
 Pwd=goback(ChkStr(Request("Pwd"),1),"请输入新管理密码")
 Pwd2=goback(ChkStr(Request("Pwd2"),1),"请输入确认新管理密码")

 Pwd_View=goback(ChkStr(Request("Pwd_View"),1),"请输入新查看密码")
 Pwd_View2=goback(ChkStr(Request("Pwd_View2"),1),"请输入确认新查看密码")
 If Pwd_View<>Pwd_View2 Then Call AlertBack("新查看密码不一致",1)
 If Pwd<>Pwd2 Then Call AlertBack("新管理密码不一致",1)
 Pwd_Old=goback(ChkStr(Request("Pwd_Old"),1),"请输入旧管理密码确认")
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd_Old,1)<>Rs("Pwd") Then Call AlertBack("输入的旧管理密码有错误:",1)
 Rs("Pwd")=Md5(Pwd,1)
 Rs("Pwd_View")=Md5(Pwd_View,1)
 Rs.update
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("Pwd")=Md5(Pwd,1)
 Rs("UserPwd")=Pwd
 Rs.update

 Call AlertBack("密码修改成功",1)
End If

If Action="passwordanswermodifysave" Then '用户修改提示问题和答案
 
 Set rs = Server.CreateObject("ADODB.Recordset")
 sql="select * from CFCount_User where UserName='"&UserName&"'"
 rs.open sql,conn,1,1
 If Rs("PasswordAsk")<>"" Then
  PasswordAnswer_Old=GoBack(ChkStr(Request("PasswordAnswer_Old"),1),"请输入原密码回答答案")
  if Rs("PasswordAnswer")<>Md5(PasswordAnswer_Old,1) Then Call AlertBack("原密码回答答案有错误",1)
 End If

 PasswordAsk_New=GoBack(ChkStr(Request("PasswordAsk_New"),1),"请输入新密码提示问题")
 PasswordAnswer_New=GoBack(ChkStr(Request("PasswordAnswer_New"),1),"请输入新密码回答答案")

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PasswordAsk")=PasswordAsk_New
 Rs("PasswordAnswer")=Md5(PasswordAnswer_New,1)
 Rs.update
 Call AlertBack("修改成功,请牢记！",1)
End If

If Action="stylemodifysave" Then
 Style=GoBack(ChkStr(Request("Style"),2),"请选择一种你喜欢样式!")
 CounterShow=Int(Request("CounterShow"))

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("CounterShow")=CounterShow
 Rs("style")=style
 Rs.Update

 Call AlertUrl("修改成功!","?Action=stylelist")
End If


If Action="imgmodifysave" Then
 Assort=Int(Request("Assort"))
 FileName_2=ChkStr(Request("FileName_2"),1)

 If Assort=2 And FileName_2="" Then
  Call AlertBack("请选择图片",1)
 Else
  FileName=FileName_2
 End If

 If Assort=3 Then FileName="default.jpg"

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("ImgFileName")=FileName
 Rs.Update

 Call AlertUrl("修改成功!","Manage.asp?Action=imglist")
End If

If Action="jsqresetsave" Then '计数器重置
 Pwd=Goback(ChkStr(Request("Pwd"),1),"请输入密码确认")

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("密码输入有误",1)

 If Int(Request("ly"))=-1 Then
  Sql="Delete From CFCount_Ly_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("daytj"))=-1 Then
  Sql="Delete From CFCount_Count_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("hourtj"))=-1 Then
  Sql="Delete From CFCount_Count_Hour Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("backtj"))=-1 Then
  Sql="Delete From CFCount_Back Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("alexatj"))=-1 Then
  Sql="Delete From CFCount_Alexa_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("searchtj"))=-1 Then
  Sql="Delete From CFCount_Search_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("keywordtj"))=-1 Then
  Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("webtj"))=-1 Then
  Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("RealShowTotal"))=-1 Then
  Sql="Update CFCount_User Set RealShowTotal=0 Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("RealIpTotal"))=-1 Then
  Sql="Update CFCount_User Set RealIpTotal=0 Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 Call Alertback("你的已经重置成功！",1)
End If

If Action="userdelsave" Then '账号删除
 Pwd=Goback(ChkStr(Request("Pwd"),1),"请输入密码确认")
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("密码输入有误",1)

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

 Call AlertUrl("你的账号已经完全删除！","?Action=logout")
End If


If Action="gbookmodifysave" Then
 ID=ChkStr(Request("ID"),1)
 
 Content=GoBack(ChkStr(Request("Content"),1),"请填入留言内容!")
 Contact=GoBack(ChkStr(Request("Contact"),1),"请填入联系方式!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Gbook Where UserName='"&UserName&"' And ID="&ID
 Rs.Open Sql,Conn,3,2
 Rs("Content")=Content
 Rs("Contact")=Contact
 Rs.Update
 
 Call AlertBack("修改成功!",2)
End If

If Action="gbookdel" Then
 ID=ChkStr(Request("ID"),1)
 
 Sql="Delete From CFCount_Gbook Where UserName='"&UserName&"' And ID="&ID
 Conn.Execute Sql
 
 Call AlertBack("删除成功!",1)
End If

If Action="urlgo" Then
 Url=URLDecode(Request("Url"))
 Call GotoUrl(Url)
End If


If Action="userlogout" Then
 Session("CFCountUser")=""
 Session("CFCountUser_View")=""
 Response.Cookies("CFCountUserCooKie")=""
 Response.Cookies("CFCountUserCooKie").Expires=Dateadd("n",-60,Now())
 Response.Redirect "Index.asp"
End If
%>
