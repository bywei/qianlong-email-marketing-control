
<%
  dim UserId
If Action="userloginsave" Then
 UserName=Trim(GoBack(ChkStr(Request("UserName"),1),"�������û�����"))
 Pwd=Trim(GoBack(ChkStr(Request("Pwd"),1),"�������û�����"))
 Cookies_Time=CLng(ChkStr(Request("Cookies_Time"),2))
 CheckCode=ChkStr(Trim(Request("CheckCode")),1)

 CFCountUserCookie=Md5(1&UserName&Pwd&GetMySet("SysCode"),1)

 If Session("ValidCode")<>CheckCode Then Call AlertUrl("��λ���ֵ���֤���������!","Index.asp")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Rs.Eof And Rs.Bof Then Call AlertUrl("û�д�ע���û�!","Index.asp")
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertUrl("�������!","Index.asp")
 If cstr(Rs("EmaState"))<>cstr("true") Then Call AlertUrl("�ʺ�δ��ˣ�����ϵ�ͷ���˺�ʹ��","Index.asp")
 
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

Admin=GoBack(ChkStr(Request("Admin"),1),"���������Ա����")
Pwd=GoBack(ChkStr(Request("Pwd"),1),"���������Ա����")
Cookies_Time=CLng(ChkStr(Request("Cookies_Time"),2))
CheckCode=ChkStr(Trim(Request("CheckCode")),1)

Ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip = "" Then Ip=Request.ServerVariables("REMOTE_ADDR")

CFCountAdminCookie=Md5(0&Admin&Pwd&GetMySet("SysCode"),1)

If Session("ValidCode")<>CheckCode Then Call AlertUrl("��λ���ֵ���֤���������!","Admin.asp")

Set Rs = Server.CreateObject("ADODB.Recordset")
Sql="Select * From CFCount_Admin where Admin='"&Admin&"'"
rs.open sql,conn,3,2

IF Rs.eof and Rs.bof Then Call AlertUrl("û�д˹���Ա����!","Admin.asp")
IF Rs("Pwd")<>Md5(Pwd,1) Then Call AlertUrl("�������!","Admin.asp")

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
  Response.Write "������Ҫ�����û���"
  Response.End
 End if

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="Select Count(*) From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs(0)=0 Then
  Response.Write "��ϲ�����û���û�б�ע��!"
 Else
  Response.Write "�Բ��𣬴��û����Ѿ�������ע����!"
 End if  

 Response.End
End if

If Action="userregsave" Then

 If RsSet("RegState")=0 Then Call AlertBack("��ϵͳ�Ѿ���ͣ���û�ע��!",1)

 UserName=GoBack(ChkStr(Trim(Request("UserName")),1),"������ע���û���!")
 Call CheckInput_Letter(UserName)
 Pwd=GoBack(ChkStr(Request("Pwd"),1),"����������!")
 Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"���ٴ���������!")

 PasswordAsk=ChkStr(Request("PasswordAsk"),1)
 PasswordAnswer=ChkStr(Request("PasswordAnswer"),1)
 
 Email=GoBack(ChkStr(Request("Email"),1),"������E-mail!")
 
 QQ=ChkStr(Request("QQ"),1)
 PageName=GoBack(ChkStr(Request("PageName"),1),"��������ҳ����!")
 PageUrl=GoBack(ChkStr(Request("PageUrl"),1),"��������ҳ��ַ!")
 
 QuickStyle=CInt(ChkStr(Request("QuickStyle"),2))
 TjOpen=ChkStr(Request("TjOpen"),2)
 ShowTotal=CLng(ChkStr(Request("ShowTotal"),2))
 
 If TjOpen="" Then TjOpen=0

 
 CheckCode=Trim(Request("CheckCode"))
 
 If Session("ValidCode")<>CheckCode Then Call AlertUrl("��λ���ֵ���֤���������!","?Reg.asp")
 
 If Pwd<>Pwd2 Then Call AlertBack("��������벻һ�£�����������һ��!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select Count(*) From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs(0)>0 Then Call AlertBack("���û����Ѿ�������ע��,�뻻��!",1)

'����û�����CFCount_user,�Ժ󲻻���û����洢���ñ���
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
  Call AlertBack("���û����Ѿ�������ע��,�뻻��!",1)
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

Call AlertUrl("ע��ɹ�,�ʼ�Ӫ��Ч������ϵͳ��̨!","Manage.asp")

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
  Call AlertUrl("�ش������","?")
 End If
Else
 Call AlertUrl("�û��������ڣ�","Index.asp")
End If

Response.Redirect "?Action=pwdmodify&UserCode="&Rs("UserCode")

End If

If Action="pwdmodifysave" Then

Pwd=GoBack(ChkStr(Request("Pwd"),1),"����������!")
Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"�������ظ�����!")
UserCode=GoBack(ChkStr(Request("UserCode"),1),"�����������!")


If Pwd<>Pwd2 Then Call AlertBack("������������벻һ�£�",1)

Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserCode='" &UserCode& "'"
Rs.Open Sql,Conn,3,2

If Rs.Eof And Rs.Bof Then
 Call AlertBack("�����벻����!",1)
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

Call AlertUrl("�޸ĳɹ������μ����������","Index.asp")
End If



If Action="gbookaddsave" Then
 UserName=ChkStr(Request("UserName"),1)
 Content=ChkStr(Request("Content"),1)
 Contact=ChkStr(Request("Contact"),1)
 Ly=Left(ChkStr(URLDecode(Request("Ly")),1),255)
 CurrWeb=Left(ChkStr(URLDecode(Request("CurrWeb")),1),255)

  
 If Content="" Then Call AlertClose("��������������")
 If Contact="" Then Call AlertClose("��������ϵ��ʽ")

Sql="Select * From CFCount_Gbook Where UserName='"&UserName&"' And Content='"&Content&"' And Contact='"&Contact&"' And DateDiff('d',AddTime,Now())=0"
Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,Conn,3,2
If Not Rs.Eof Then
 Call AlertClose("�벻Ҫ����ͬ������")
Else
 Rs.AddNew
 Rs("UserName")=UserName
 Rs("Content")=Content
 Rs("Contact")=Contact
 Rs("Ly")=Ly
 Rs("CurrWeb")=CurrWeb
 Rs.Update
End If

Call AlertClose("���Գɹ�")

End If


'�鿴�û���¼
If Action="viewuserlogin" Then
 UserName=GoBack(ChkStr(Request("UserName"),1),"�������û�����")
 Pwd_View=GoBack(ChkStr(Request("Pwd_View"),1),"����������鿴����")
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 UserId= Rs("id")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserId="&UserId&""
 Rs.Open Sql,Conn,1,1

 If Rs.Eof And Rs.Bof Then Call AlertBack("û�д�ע���û�!",1)

 If Rs("Pwd_View")<>Md5(Pwd_View,1) Then Call AlertBack("�����鿴�������!",1)
 
 Session("CFCountUser_View")=UserName
 Response.Redirect "Manage.asp"
End If
%>