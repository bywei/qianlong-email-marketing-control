<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn.asp"-->
<%
dim action,userId,userName,userPwd,userEmail
action = trim(request.QueryString("action"))
adminName=trim(request.Form("adminName"))
adminPwd=trim(request.Form("adminPwd"))


'登陆
if action="login" then
     chksql="select * from admin"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("adminName")=adminName and chkrs("adminPwd")=adminPwd  then
		   session("admin")=adminName
		   response.Redirect("admin.asp")
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
	   conn.close
    set conn=nothing
end if

if action="loginout" and session("admin") <> "" then
   session("admin")=""
   response.Redirect("index.html")
end if 

if action ="reg" then
  regUser()
end if

'获取用户信息
function getUserInfo()
  userId=request.Form("userId")
  chksql="select * from Users"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userId")=userId  then
		   response.Write("<script>window.location.href='?login=access&userScore="&chkrs("UserScore")&"&userLevel="&chkrs("UserLevel")&"&userSendNum="&chkrs("userSendNum")&"&userIsVip="&chkrs("userIsVip")&"'</script>")
		   response.End()
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
	   conn.close
    set conn=nothing
end function

'注册用户
function reg(userName,userPwd,userEmail)
'action=insert&userName=username&userPwd=userpwd&userEmail=useremail
userEmail=request.Form("userEmail")
     sql="select * from Users"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName then
		   response.Write("<script>alert('该群发账号已经花有名主了啦!');</script>")
		   response.End()
		end if
	     chkrs.movenext
	loop
	  chkrs.close()
   
      set rs=server.CreateObject("adodb.recordset")
	  rs.open sql,conn,1,3
	  rs.addnew
	  rs("UserName")=userName
	  rs("UserPwd")=userPwd
	  rs("UserEmail")=userEmail
	  rs.update
	  rs.close
	  set rs=nothing
	   conn.close
    set conn=nothing
	  response.Write("<script>alert('邮件群发账号注册成功!');</script>")
	  response.End()
end function


function updateInfo()
    '修改数据    '?update=access&userScore="&chkrs("UserScore")&"&userLevel="&chkrs("UserLevel")&"&userSendNum="&chkrs("userSendNum")&"&userIsVip="&chkrs("userIsVip")&"
   dim userScore,userLevel,userSendNum,userIsVip
   userScore=request.QueryString("userScore")
   userLevel=request.QueryString("userLevel")
   userSendNum=request.QueryString("userSendNum")
   userIsVip=request.QueryString("userIsVip")
   
   dim oldUserScore,oldUserLevel,oldUserSendNum,oldUserIsVip

	'查询原来的数据
	chksql="select * from Users where UserName='"&userName&"'"
	 set chkrs=server.CreateObject("adodb.recordset")
	 chkrs.open chksql,conn,1,3
	 
	 do while not chkrs.eof
	    if chkrs("userName")=userName and chkrs("UserPwd")=userPwd  then
		   oldUserScore= chkrs("UserScore")
		   oldUserLevel=chkrs("UserLevel")
		   oldUserSendNum=chkrs("UserSendNum")
		   oldUserIsVip=chkrs("UserIsVip")
		end if
	     chkrs.movenext
	loop
	   chkrs.close()
	   conn.close
    set conn=nothing
	
	'修改数据
     upsql="select * from Users where UserName='"&userName&"'"
	 set uprs=server.CreateObject("ADODB.recordset")
	 uprs.open upsql,conn,1,3
	 uprs("userScore")=userScore+oldUserScore
	 uprs("userLevel")=userLevel
	 uprs("userSendNum")=userSendNum+oldUserSendNum
	 uprs("userIsVip")=userIsVip
     uprs.update
     uprs.close
     set uprs=nothing
     conn.close
    set conn=nothing
end function

response.Redirect("index.html")
response.End()
%>