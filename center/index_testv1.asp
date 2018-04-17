<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<%
'注册用户
'http://www.35so.bywei.cn/EmailUsers/index.asp?action=insert&userName=username&userPwd=userpwd&userEmail=useremail
 dim action,userId,userName,userPwd,userEmail,userSendNum
   action=trim(request.Form("action"))
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
	 response.End()
  end if
 
 '登录并返回数据 
  if action="login" then
     chksql="select * from Users"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName and chkrs("UserPwd")=userPwd  then
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
 
 
 '获取用户信息
 if action="getUserInfo" then
        chksql="select * from Users where id="&userId
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
		    response.write( "{ ""userScore"":"""&chkrs("UserScore")&""",""userLevel"":"""&chkrs("UserLevel")&""",""userSendNum"":"""&chkrs("userSendNum")&""",""userIsVip"":"&chkrs("userIsVip")&"}")
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
	 uprs("userSendNum")=userSendNum+uprs("userSendNum")
	  if userSendNum > 0 then 
	     uprs("userScore")=uprs("userScore")+int(userSendNum * 0.4)
	  end if
     uprs.update
     uprs.close
     set uprs=nothing
     conn.close
     set conn=nothing
     response.write( "{ ""UserSendNum"":"""&userId&"}")
	response.End()
   end if


 
    
   
   
   '广告设置
   if action="ad" then 
      response.Write("<script>window.location.href='Http://www.bywei.cn/blog/'</script>")
   end if
%>
