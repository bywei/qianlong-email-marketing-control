<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<%
'注册用户
'http://www.35so.bywei.cn/EmailUsers/index.asp?action=insert&userName=username&userPwd=userpwd&userEmail=useremail
 dim action,userId,userName,userPwd,userEmail,userSendNum
   action=trim(request.Form("action"))
   'action=trim(request.QueryString("action"))
   'userId=trim(request.QueryString("userId"))
   userId=trim(request.Form("userId"))
   userName=trim(request.Form("userName"))
   userPwd=trim(request.Form("userPwd"))
   userEmail=trim(request.Form("userEmail"))
   userSendNum=request.Form("UserSendNum")
   dim sql,chksql,upsql,domianSql
   
   if action="insert" then
     chksql="select * from Users"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName then
		   '该群发账号已经花有名主了啦！
		   response.Write("false")
		   response.End()
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
   
      sql="select * from Users"
      set rs=server.CreateObject("adodb.recordset")
	  rs.open sql,conn,1,3
	  rs.addnew
	  rs("UserName")=userName
	  rs("Pwd")="965eb72c92a549dd"
	  rs("UserPwd")=userPwd
	  rs("UserEmail")=userEmail
	  rs("UserScore")=0
	  rs("UserLevel")=1
	  rs("userSendNum")="0"
	  rs("UserIsVip")="false"
	  rs("UserState")=1
	  rs("UserType")=1
	  rs("CreateTime")=Date()
	  rs("UseOfTime")=dateadd("M",1,date)
	  rs("EmaState")="true"
	  rs.update
	  
	  
	   Set Rs=Server.CreateObject("Adodb.RecordSet")
       Sql="Select * From Users Where UserName='"&userName&"'"
       Rs.Open Sql,Conn,1,1
       userId= Rs("id")

       Set Rs=Server.CreateObject("Adodb.RecordSet")
       Sql="Select * From CFCount_User"
       Rs.Open Sql,Conn,3,2
      Rs.AddNew
      Rs("UserId")=userId
      Rs("UserName")=userName
      Rs("Pwd")="965eb72c92a549dd"
      Rs("Pwd_View")="965eb72c92a549dd"
      Rs("Email")=userEmail
      Rs("PageName")="邮件营销效果跟踪"
      Rs("PageUrl")="http://ema.qianlongsoft.com/"
      Rs.Update
	  rs.close
	  set rs=nothing
	   conn.close
    set conn=nothing
	  'response.Write("<script>alert('邮件群发账号注册成功!');</script>")
	  response.Write("true")
	  response.End()
  end if
  
  if action="" then
     response.Write("none")
     response.redirect("../")
	 response.End()
  end if
 
 '登录并返回数据 
  if action="login" then
  	'测试默认永远返回登录成功
  	response.write("{""islogin"":true,""userId"":""1""}")
		response.End()
		     
     chksql="select * from Users where userName='"&userName&"'"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName and chkrs("UserPwd")=userPwd  and chkrs("UserState")=0 then
		  if DateDiff("d",chkrs("CreateTime"),dateadd("m",1,chkrs("UseOfTime")))>0 then
                   if chkrs("loginState")="true" then
                     chkrs("loginState")="false"
                     chkrs("LoginTime")=now()
                     chkrs.update
		     response.write("{""islogin"":true,""userId"":"""&chkrs("Id")&"""}")
		     response.End()
                   end if
                   if Datediff("n",chkrs("LoginTime"),dateadd("n",5,now()))>0  then
                     chkrs("loginState")="true"
                     chkrs("LoginTime")=now()
                     chkrs.update
		     response.write("{""islogin"":true,""userId"":"""&chkrs("Id")&"""}")
		     response.End()
                   end if
		   else
		     response.write("{""islogin"":false}")
	         response.End()
		   end if
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
	   conn.close
    set conn=nothing
	 response.write("{""islogin"":false}")
	 response.End()
 end if
 
 
 '获取用户信息
 if action="getUserInfo" then
        chksql="select * from Users where id="&userId
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
		    response.write( "{ ""userScore"":"""&chkrs("UserScore")&""",""userLevel"":"""&chkrs("UserLevel")&""",""userSendNum"":"""&chkrs("userSendNum")&""",""userIsVip"":"&chkrs("userIsVip")&",""createTime"":"""&chkrs("createTime")&""",""useOfTime"":"""&chkrs("useOfTime")&""",""loginTime"":"""&chkrs("loginTime")&"""}")
		   response.End()
	     chkrs.movenext
	loop
	  chkrs.close()
	   conn.close
    set conn=nothing
	response.write( "{ ""userScore"":""0"",""userLevel"":""0"",""userSendNum"":""0"",""userIsVip"":false}")
	 response.End()
 end if
   
   
   '修改数据    '?update=access&userScore="&chkrs("UserScore")&"&userLevel="&chkrs("UserLevel")&"&userSendNum="&chkrs("userSendNum")&"&userIsVip="&chkrs("userIsVip")&"
   '更新数据
   if action="update" then
	'修改数据
     upsql="select * from Users where Id="&userId&""
	 set uprs=server.CreateObject("ADODB.recordset")
	 uprs.open upsql,conn,1,3
        if cint(userSendNum) >0 then
	 uprs("userSendNum")=uprs("userSendNum")+cint(userSendNum)
	 uprs("userScore")=uprs("userScore")+cint(cint(userSendNum) * 0.4)
        end if
          uprs("loginState")="true"
          uprs("LoginTime")=now()
     uprs.update
     uprs.close
     set uprs=nothing
     conn.close
     set conn=nothing
     response.write( "{ ""UserSendNum"":"""&userId&"}")
	response.End()
   end if


  '获取用户域名后缀信息
  if action ="getUserDomain" then 
    dim result
	result=""
    domianSql="select * from userDomain where userId="&userId
    set domainRs=server.CreateObject("adodb.recordset")
	domainRs.open domianSql,conn,1,3
    do while not domainRs.eof
	 result=result&""""&domainRs("id")&""":"""&domainRs("userDomainName")&""","
         domainRs.movenext
   loop
	set domainRs=nothing
	conn.close()
	
	if result="" then 
            response.Write("{""0"":""qianlongsoft.com""}")
	   response.End()
	else
	  response.write("{"&result&"""0"":""bywei.cn""}")
	   response.End()
	end if
  end if
    
   
   
   '广告设置
   if action="ad" then 
      response.Write("<script>window.location.href='Http://www.bywei.cn/blog/'</script>")
   end if
%>
