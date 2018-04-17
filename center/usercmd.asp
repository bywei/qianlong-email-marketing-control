<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<!--#include file="chklogin.asp"-->
<%
 dim action,userId,userName,userPwd,userEmail,userScore,userLevel
 dim userSendNum,userIsVip,userState,loginState,userType,createTime,useOfTime
   action=trim(request.Form("action"))
   if action ="" then 
    action=trim(request.QueryString("action"))
   end if
   userId=trim(request.Form("userId"))
   if userId="" then
     userId=trim(request.QueryString("userId"))
   end if
   userName=trim(request.Form("userName"))
   userPwd=trim(request.Form("userPwd"))
   userEmail=trim(request.Form("userEmail"))
   userScore=trim(request.Form("userScore"))
   userLevel=trim(request.Form("userLevel"))
   
   userSendNum=trim(request.Form("UserSendNum"))
    userIsVip=trim(request.Form("userIsVip"))
	userState=trim(request.Form("userState"))
        loginState=trim(request.Form("loginState"))
	userType=trim(request.Form("userType"))
	createTime=trim(request.Form("createTime"))
	useOfTime=trim(request.Form("useOfTime"))


  
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
	  rs("UserPwd")=userPwd
	  rs("UserEmail")=userEmail
	  rs("UserScore")=0
	  rs("UserLevel")=1
	  rs("userSendNum")="0"
	  rs("UserIsVip")="false"
	  rs.update
	  rs.close
	  set rs=nothing
	   conn.close
    set conn=nothing
	  'response.Write("<script>alert('邮件群发账号注册成功!');/script>")
	  response.Write("true")
	  response.End()
  end if
  
  if action="" then
     response.Write("none")
     response.redirect("login.html")
	 response.End()
  end if
 
 '登录并返回数据 
  if action="login" then
     chksql="select * from Users "
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName and chkrs("UserPwd")=userPwd  and chkrs("UserState")=0 then
		   response.write("{""islogin"":true,""userId"":"""&chkrs("Id")&"""}")
		   response.End()
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
	   conn.close
    set conn=nothing
	 response.write("{""islogin"":false}")
	 response.End()
 end if

  
   if action="update" then
     if userId="" then 
	    response.Redirect("userlist.asp")
	  end if
	'修改数据
     upsql="select * from Users where Id="&userId&""
	 set uprs=server.CreateObject("ADODB.recordset")
	 uprs.open upsql,conn,1,3
	    uprs("userSendNum")=userSendNum
		uprs("id")=userId
        uprs("userName")=userName
        uprs("userPwd")=userPwd
        uprs("userEmail")=userEmail
        uprs("userScore")=userScore
        uprs("userLevel")=userLevel
   
        uprs("UserSendNum")=userSendNum
        uprs("userIsVip")=userIsVip
	    uprs("userState")=userState
        uprs("loginState")=loginState
	    uprs("userType")=userType
	    uprs("createTime")=createTime
	    uprs("useOfTime")=useOfTime
     uprs.update
     uprs.close
     set uprs=nothing
	 
	 Ip_address=Request.ServerVariables ("HTTP_X_FORWARDED_FOR")
    If Ip_address=""Then  
       Ip_address= Request.ServerVariables ("REMOTE_ADDR")
    end if
	 '记录日志
	 upsql="select * from comLog "
	 set rs=server.CreateObject("ADODB.recordset")
	 rs.open upsql,conn,1,3
	 rs.addnew()
	   rs("logTime")=date()
	   rs("logUser")=session("admin")
	   rs("logIp")=Ip_address
	    rs("userSendNum")=userSendNum
		rs("id")=userId
        rs("userName")=userName
        rs("userPwd")=userPwd
        rs("userEmail")=userEmail
        rs("userScore")=userScore
        rs("userLevel")=userLevel
   
        rs("UserSendNum")=userSendNum
        rs("userIsVip")=userIsVip
	    rs("userState")=userState
        rs("loginState")=loginState
	    rs("userType")=userType
	    rs("createTime")=createTime
	    rs("useOfTime")=useOfTime
     rs.update
     rs.close
     set rs=nothing
	 
     conn.close
     set conn=nothing
     response.Redirect("userlist.asp?msg=修改成功&userId="&userId)
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
            response.Write("{""0"":""bywei.cn""}")
	   response.End()
	else
	  response.write("{"&result&"""0"":""bywei.cn""}")
	   response.End()
	end if
  end if
  
 if action = "delUser" then
    if userId="" then
	  response.Redirect("userlist.asp?msg=修改失败")
	 end if
    delsql="delete from Users where id="&userId
	 set delrs=server.CreateObject("adodb.recordset")
	 delrs.open delsql,conn,1,3
	  delrs.update()
	  delrs.close()
	  conn.close
    set conn=nothing
	response.Redirect("userlist.asp?msg=修改成功&userId="&userId)
  end if

%>